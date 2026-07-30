# VP daily report field calculator
#
# Usage:
#   Rscript scripts/calc_vp_daily_report.R "C:/path/to/source.xlsx" "2026-07-01" "outputs/vp_daily_fields.csv"
#
# Args:
#   1. input workbook path, required
#   2. run date, optional, yyyy-mm-dd. The script calculates T-1.
#   3. output csv path, optional.
#
# Required packages: readxl, dplyr, tidyr, stringr, glue, readr

user_libraries <- c(
  file.path(Sys.getenv("USERPROFILE"), "Documents", "R", "win-library", paste0(R.version$major, ".", strsplit(R.version$minor, "\\.")[[1]][1])),
  file.path(Sys.getenv("USERPROFILE"), "Documents", "R", "win-library", paste0(R.version$major, ".", R.version$minor))
)
existing_user_libraries <- user_libraries[dir.exists(user_libraries)]
if (length(existing_user_libraries) > 0) {
  .libPaths(c(existing_user_libraries, .libPaths()))
}
required_packages <- c("readxl", "dplyr", "tidyr", "stringr", "glue", "readr")
missing_packages <- required_packages[!vapply(required_packages, requireNamespace, logical(1), quietly = TRUE)]
if (length(missing_packages) > 0) {
  stop(
    "Missing R packages: ",
    paste(missing_packages, collapse = ", "),
    ". Please install them before running this script.",
    call. = FALSE
  )
}

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(stringr)
  library(glue)
  library(readr)
})

args <- commandArgs(trailingOnly = TRUE)
input_path <- if (length(args) >= 1) args[[1]] else stop("Please provide the input workbook path as the first argument.", call. = FALSE)
run_date <- if (length(args) >= 2) as.Date(args[[2]]) else Sys.Date()
output_path <- if (length(args) >= 3) args[[3]] else "outputs/vp_daily_fields.csv"
target_date <- run_date - 1

as_report_date <- function(x) {
  if (inherits(x, "Date")) return(x)
  if (inherits(x, "POSIXt")) return(as.Date(x))

  out <- rep(as.Date(NA), length(x))

  is_posix <- vapply(x, inherits, logical(1), what = "POSIXt")
  if (any(is_posix, na.rm = TRUE)) {
    out[is_posix] <- as.Date(as.POSIXct(x[is_posix], origin = "1970-01-01", tz = "UTC"))
  }

  is_num <- !is.na(x) & suppressWarnings(!is.na(as.numeric(x)))
  if (any(is_num, na.rm = TRUE)) {
    num <- suppressWarnings(as.numeric(x[is_num]))
    idx <- which(is_num)

    # Excel serial dates, e.g. 45838 for 2025-07-01 style dates.
    is_excel_serial <- num >= 20000 & num <= 80000
    out[idx[is_excel_serial]] <- as.Date(num[is_excel_serial], origin = "1899-12-30")

    # yyyymmdd integers, e.g. 20260630.
    is_yyyymmdd_num <- num >= 19000101 & num <= 29991231
    out[idx[is_yyyymmdd_num]] <- as.Date(as.character(as.integer(num[is_yyyymmdd_num])), format = "%Y%m%d")

    # POSIX seconds accidentally produced when POSIXct cells are unlisted.
    is_unix_seconds <- num > 1000000000 & num < 3000000000
    out[idx[is_unix_seconds]] <- as.Date(as.POSIXct(num[is_unix_seconds], origin = "1970-01-01", tz = "UTC"))
  }

  x_chr <- as.character(x)
  valid <- is.na(out) & !is.na(x_chr) & x_chr != "" & x_chr != "NA"

  is_yyyymmdd <- rep(FALSE, length(x_chr))
  is_yyyymmdd[valid] <- str_detect(x_chr[valid], "^\\d{8}$")
  out[is_yyyymmdd] <- as.Date(x_chr[is_yyyymmdd], format = "%Y%m%d")

  is_datetime <- rep(FALSE, length(x_chr))
  is_datetime[valid] <- str_detect(x_chr[valid], "^\\d{4}-\\d{2}-\\d{2}")
  out[is_datetime] <- as.Date(substr(x_chr[is_datetime], 1, 10))

  out
}

to_num <- function(x) suppressWarnings(as.numeric(x))

fmt_num <- function(x) {
  out <- ifelse(is.na(x), "/", format(round(x, 0), big.mark = ",", trim = TRUE, scientific = FALSE))
  as.character(out)
}

fmt_pct <- function(x, space_before = TRUE) {
  if (length(x) == 0 || is.na(x) || is.infinite(x)) return("/")
  sign <- ifelse(x >= 0, "+", "")
  spacer <- ifelse(space_before, " ", "")
  paste0(spacer, sign, sprintf("%.1f%%", x * 100))
}

fmt_pct_value <- function(x) {
  if (length(x) == 0 || is.na(x) || is.infinite(x)) return("/")
  sprintf("%.1f%%", x * 100)
}

trend_word <- function(x, up = "有增长", down = "下降", flat = "基本持平") {
  if (is.na(x)) return("变化")
  if (x > 0.001) return(up)
  if (x < -0.001) return(down)
  flat
}

safe_rate <- function(current, baseline) {
  if (is.na(current) || is.na(baseline) || baseline == 0) return(NA_real_)
  current / baseline - 1
}

zero_if_na <- function(x) {
  ifelse(is.na(x), 0, x)
}

weekday_id <- function(date) {
  # POSIXlt wday: Sunday = 0, Monday = 1, ... Saturday = 6
  as.POSIXlt(date)$wday
}

get_compare_rule <- function(date) {
  # Holiday/special-date rules can be inserted here later.
  w <- weekday_id(date)
  dplyr::case_when(
    w == 1 ~ "prev_week_weekday_avg",
    w %in% 2:5 ~ "previous_day",
    w == 6 ~ "prev_weekend_avg",
    w == 0 ~ "weekly_vs_last_week",
    TRUE ~ "previous_day"
  )
}

get_compare_label <- function(rule) {
  dplyr::case_when(
    rule == "prev_week_weekday_avg" ~ "上一周周中日均",
    rule == "previous_day" ~ "前一日",
    rule == "prev_weekend_avg" ~ "上周末日均",
    rule == "weekly_vs_last_week" ~ "上周",
    TRUE ~ "基准"
  )
}

get_date_sets <- function(date, rule) {
  if (rule == "prev_week_weekday_avg") {
    list(current = date, baseline = seq(date - 7, date - 3, by = "day"))
  } else if (rule == "previous_day") {
    list(current = date, baseline = date - 1)
  } else if (rule == "prev_weekend_avg") {
    list(current = date, baseline = seq(date - 7, date - 6, by = "day"))
  } else if (rule == "weekly_vs_last_week") {
    list(current = seq(date - 6, date, by = "day"), baseline = seq(date - 13, date - 7, by = "day"))
  } else {
    list(current = date, baseline = date - 1)
  }
}

summarise_value <- function(data, scope_name, metric_name, dates, agg = c("sum", "mean")) {
  agg <- match.arg(agg)
  values_by_date <- data %>%
    filter(.data$scope == scope_name, .data$metric == metric_name, .data$date %in% dates) %>%
    group_by(.data$date) %>%
    summarise(value = dplyr::last(.data$value[!is.na(.data$value)]), .groups = "drop") %>%
    pull(.data$value)
  if (length(values_by_date) == 0 || all(is.na(values_by_date))) return(NA_real_)
  if (agg == "mean") mean(values_by_date, na.rm = TRUE) else sum(values_by_date, na.rm = TRUE)
}

calc_metric_compare <- function(data, scope_name, metric_name, date_sets, rule) {
  baseline_agg <- ifelse(rule %in% c("prev_week_weekday_avg", "prev_weekend_avg"), "mean", "sum")
  current <- summarise_value(data, scope_name, metric_name, date_sets$current, "sum")
  baseline <- summarise_value(data, scope_name, metric_name, date_sets$baseline, baseline_agg)
  tibble(scope = scope_name, metric = metric_name, current = current, baseline = baseline, rate = safe_rate(current, baseline))
}

read_sales_trend <- function(path, scope_names = c("全系", "P7", "G7", "M03", "P7+", "G6", "G9", "X9", "GX")) {
  raw <- readxl::read_excel(path, sheet = "销售过程数据走势", col_names = FALSE, .name_repair = "minimal")
  all_scope_names <- c("全系", "P7", "G7", "M03", "P7+", "G6", "G9", "X9", "GX", "L03")
  metric_names <- c("客流", "线索", "线上线索", "线下线索", "试驾", "锁单")
  labels <- as.character(raw[[1]])
  all_scope_rows <- which(labels %in% all_scope_names)
  scope_rows <- all_scope_rows[labels[all_scope_rows] %in% scope_names]

  bind_rows(lapply(scope_rows, function(scope_row) {
    scope_name <- labels[[scope_row]]
    date_row <- scope_row + 1L
    next_scope_row <- all_scope_rows[all_scope_rows > scope_row]
    metric_end <- if (length(next_scope_row) > 0) min(next_scope_row) - 1L else nrow(raw)
    metric_indices <- seq(scope_row + 2L, metric_end)
    metric_indices <- metric_indices[labels[metric_indices] %in% metric_names]
    dates <- as_report_date(unlist(raw[date_row, -1], use.names = FALSE))

    bind_rows(lapply(metric_indices, function(metric_row) {
      tibble(
        scope = scope_name,
        metric = labels[[metric_row]],
        date = dates,
        value = to_num(unlist(raw[metric_row, -1], use.names = FALSE))
      )
    }))
  })) %>%
    filter(!is.na(.data$date), !is.na(.data$metric), .data$metric != "")
}

read_online_official_traffic <- function(path) {
  raw <- readxl::read_excel(path, sheet = "线上官渠潜客客流", col_names = FALSE, .name_repair = "minimal")
  tibble(
    scope = "官渠",
    metric = "线上官渠潜客客流",
    date = as_report_date(unlist(raw[1, -1], use.names = FALSE)),
    value = to_num(unlist(raw[2, -1], use.names = FALSE))
  ) %>%
    filter(!is.na(.data$date))
}
read_data_bottom_model_metric <- function(path, block_name, metric_name, scope_names, scope_name = "L03") {
  raw <- readxl::read_excel(path, sheet = "数据底表", col_names = FALSE, .name_repair = "minimal")
  labels <- stringr::str_squish(as.character(raw[[1]]))
  block_names <- c("客流", "线索（分渠道）", "线索（分车系）", "试驾", "锁单")
  block_header_rows <- which(labels %in% block_names)
  block_rows <- which(labels == block_name)

  for (block_row in block_rows) {
    next_header_row <- block_header_rows[block_header_rows > block_row]
    next_header_row <- if (length(next_header_row) > 0) min(next_header_row) else nrow(raw) + 1L
    rows_in_block <- seq(block_row + 1L, next_header_row - 1L)
    scope_rows <- rows_in_block[!is.na(labels[rows_in_block]) & labels[rows_in_block] %in% scope_names]
    if (length(scope_rows) == 0) next
    scope_row <- scope_rows[[1]]

    candidate_date_rows <- seq(block_row + 1L, scope_row - 1L)
    date_counts <- vapply(candidate_date_rows, function(row) {
      sum(!is.na(as_report_date(unlist(raw[row, -1], use.names = FALSE))))
    }, integer(1))
    if (length(date_counts) == 0 || max(date_counts) == 0) next
    date_row <- candidate_date_rows[[which.max(date_counts)]]

    return(tibble(
      scope = scope_name,
      metric = metric_name,
      date = as_report_date(unlist(raw[date_row, -1], use.names = FALSE)),
      value = to_num(unlist(raw[scope_row, -1], use.names = FALSE))
    ) %>%
      filter(!is.na(.data$date)))
  }

  stop(
    "Could not find 数据底表/", block_name, "/", paste(scope_names, collapse = "|"),
    " for metric ", metric_name, ".",
    call. = FALSE
  )
}

read_order_metrics_from_data_bottom <- function(path) {
  bind_rows(
    read_data_bottom_model_metric(path, "锁单", "锁单", c("总计"), scope_name = "全系"),
    read_data_bottom_model_metric(path, "锁单", "锁单", c("小鹏GX", "GX"), scope_name = "GX"),
    read_data_bottom_model_metric(path, "锁单", "锁单", c("L03"), scope_name = "L03")
  )
}
read_l03_metrics_from_month_compare <- function(path) {
  raw <- readxl::read_excel(path, sheet = "月累环比 (2)", range = "A1:J15", col_names = FALSE, .name_repair = "minimal")
  labels <- as.character(raw[[1]])
  l03_row <- which(labels == "L03")[[1]]
  dates <- as_report_date(c(raw[[2]][2], raw[[3]][2]))

  bind_rows(
    tibble(scope = "L03", metric = "线索", date = dates, value = to_num(c(raw[[2]][l03_row], raw[[3]][l03_row]))),
    tibble(scope = "L03", metric = "试驾", date = dates, value = to_num(c(raw[[5]][l03_row], raw[[6]][l03_row])))
  ) %>%
    filter(!is.na(.data$date))
}

read_l03_metrics <- function(path) {
  trend_data <- read_sales_trend(path, scope_names = "L03") %>%
    filter(.data$metric %in% c("线索", "试驾", "锁单"))
  month_compare_data <- read_l03_metrics_from_month_compare(path)
  order_data <- read_order_metrics_from_data_bottom(path) %>% filter(.data$scope == "L03")

  bind_rows(trend_data, month_compare_data, order_data) %>%
    arrange(.data$scope, .data$metric, .data$date) %>%
    group_by(.data$scope, .data$metric, .data$date) %>%
    summarise(value = dplyr::last(.data$value[!is.na(.data$value)]), .groups = "drop")
}
build_calc_table <- function(sales_data, online_data, l03_data, date_sets, rule) {
  all_data <- bind_rows(sales_data, online_data, l03_data)
  base <- bind_rows(
    calc_metric_compare(all_data, "官渠", "线上官渠潜客客流", date_sets, rule),
    calc_metric_compare(all_data, "全系", "客流", date_sets, rule),
    calc_metric_compare(all_data, "全系", "线索", date_sets, rule),
    calc_metric_compare(all_data, "全系", "线上线索", date_sets, rule),
    calc_metric_compare(all_data, "全系", "线下线索", date_sets, rule),
    calc_metric_compare(all_data, "GX", "线索", date_sets, rule),
    calc_metric_compare(all_data, "L03", "线索", date_sets, rule),
    calc_metric_compare(all_data, "全系", "试驾", date_sets, rule),
    calc_metric_compare(all_data, "L03", "试驾", date_sets, rule),
    calc_metric_compare(all_data, "GX", "试驾", date_sets, rule),
    calc_metric_compare(all_data, "全系", "锁单", date_sets, rule),
    calc_metric_compare(all_data, "L03", "锁单", date_sets, rule),
    calc_metric_compare(all_data, "GX", "锁单", date_sets, rule)
  )

  get_row <- function(scope_name, metric_name) {
    base %>% filter(.data$scope == scope_name, .data$metric == metric_name) %>% slice(1)
  }

  all_leads <- get_row("全系", "线索")
  gx_leads <- get_row("GX", "线索")
  l03_leads <- get_row("L03", "线索")
  all_drives <- get_row("全系", "试驾")
  l03_drives <- get_row("L03", "试驾")
  gx_drives <- get_row("GX", "试驾")
  all_orders <- get_row("全系", "锁单")
  l03_orders <- get_row("L03", "锁单")
  gx_orders <- get_row("GX", "锁单")

  bind_rows(
    base,
    tibble(scope = "其他车型", metric = "线索", current = all_leads$current - gx_leads$current - l03_leads$current, baseline = all_leads$baseline - gx_leads$baseline - l03_leads$baseline, rate = safe_rate(all_leads$current - gx_leads$current - l03_leads$current, all_leads$baseline - gx_leads$baseline - l03_leads$baseline)),
    tibble(scope = "全系（非L03）", metric = "试驾", current = all_drives$current - l03_drives$current, baseline = all_drives$baseline - l03_drives$baseline, rate = safe_rate(all_drives$current - l03_drives$current, all_drives$baseline - l03_drives$baseline)),
    tibble(scope = "其他车型", metric = "试驾", current = all_drives$current - gx_drives$current - l03_drives$current, baseline = all_drives$baseline - gx_drives$baseline - l03_drives$baseline, rate = safe_rate(all_drives$current - gx_drives$current - l03_drives$current, all_drives$baseline - gx_drives$baseline - l03_drives$baseline)),
    tibble(scope = "其他车型", metric = "锁单", current = all_orders$current - zero_if_na(l03_orders$current) - gx_orders$current, baseline = all_orders$baseline - zero_if_na(l03_orders$baseline) - gx_orders$baseline, rate = safe_rate(all_orders$current - zero_if_na(l03_orders$current) - gx_orders$current, all_orders$baseline - zero_if_na(l03_orders$baseline) - gx_orders$baseline))
  )
}

pick_value <- function(calc_table, scope_name, metric_name, col_name) {
  calc_table %>%
    filter(.data$scope == scope_name, .data$metric == metric_name) %>%
    slice(1) %>%
    pull(all_of(col_name))
}

build_overview_text <- function(calc_table, compare_label) {
  official_rate <- pick_value(calc_table, "官渠", "线上官渠潜客客流", "rate")
  traffic_rate <- pick_value(calc_table, "全系", "客流", "rate")
  leads_rate <- pick_value(calc_table, "全系", "线索", "rate")
  online_leads_rate <- pick_value(calc_table, "全系", "线上线索", "rate")
  offline_leads_rate <- pick_value(calc_table, "全系", "线下线索", "rate")
  gx_leads_rate <- pick_value(calc_table, "GX", "线索", "rate")
  l03_leads_rate <- pick_value(calc_table, "L03", "线索", "rate")
  other_leads_rate <- pick_value(calc_table, "其他车型", "线索", "rate")
  drives_current <- pick_value(calc_table, "全系", "试驾", "current")
  drives_rate <- pick_value(calc_table, "全系", "试驾", "rate")
  l03_drives_current <- pick_value(calc_table, "L03", "试驾", "current")
  l03_drives_rate <- pick_value(calc_table, "L03", "试驾", "rate")
  l03_drives_share <- ifelse(is.na(drives_current) || drives_current == 0, NA_real_, l03_drives_current / drives_current)
  non_l03_drives_rate <- pick_value(calc_table, "全系（非L03）", "试驾", "rate")
  gx_drives_rate <- pick_value(calc_table, "GX", "试驾", "rate")
  other_drives_rate <- pick_value(calc_table, "其他车型", "试驾", "rate")
  orders_current <- pick_value(calc_table, "全系", "锁单", "current")
  orders_rate <- pick_value(calc_table, "全系", "锁单", "rate")
  l03_orders_current <- pick_value(calc_table, "L03", "锁单", "current")
  l03_orders_rate <- pick_value(calc_table, "L03", "锁单", "rate")
  gx_orders_current <- pick_value(calc_table, "GX", "锁单", "current")
  gx_orders_rate <- pick_value(calc_table, "GX", "锁单", "rate")
  other_orders_current <- pick_value(calc_table, "其他车型", "锁单", "current")
  other_orders_rate <- pick_value(calc_table, "其他车型", "锁单", "rate")

  glue(
    "线上官渠潜客客流：官渠客流{trend_word(official_rate)}，VS {compare_label}{fmt_pct(official_rate)}；\n\n",
    "门店客流：门店客流{trend_word(traffic_rate)}，VS {compare_label}{fmt_pct(traffic_rate)}；\n\n",
    "线索总量：线索总量{trend_word(leads_rate)}，VS {compare_label}{fmt_pct(leads_rate)}，",
    "其中线上线索总量{fmt_pct(online_leads_rate)}，",
    "线下线索总量{fmt_pct(offline_leads_rate)}；",
    "其中 GX{fmt_pct(gx_leads_rate)}，",
    "L03{fmt_pct(l03_leads_rate)}，",
    "其他车型合计{fmt_pct(other_leads_rate)}；\n\n",
    "试驾总量：试驾总量{trend_word(drives_rate)}，VS {compare_label}{fmt_pct(drives_rate)}；",
    "L03 试驾{fmt_num(l03_drives_current)}（占试驾总量比例{fmt_pct_value(l03_drives_share)}），VS {compare_label}{fmt_pct(l03_drives_rate)}；",
    "全系（非 L03）试驾 VS {compare_label}{fmt_pct(non_l03_drives_rate)}；",
    "GX{fmt_pct(gx_drives_rate)}，",
    "其他车型合计{fmt_pct(other_drives_rate)}；\n\n",
    "锁单总量：全系锁单总量{fmt_num(orders_current)}，VS {compare_label}{fmt_pct(orders_rate)}；",
    "其中 L03 净锁单{fmt_num(l03_orders_current)}台，{fmt_pct(l03_orders_rate)}；",
    "GX 净锁单{fmt_num(gx_orders_current)}台，{fmt_pct(gx_orders_rate)}；",
    "其他车型合计{fmt_num(other_orders_current)}台，{fmt_pct(other_orders_rate)}。"
  )
}

read_order_conversion_summary <- function(path) {
  raw <- readxl::read_excel(path, sheet = "月累环比 (2)", range = "T1:Y3", col_names = FALSE, .name_repair = "minimal")
  list(
    order_month_value = to_num(raw[[1]][3]),
    order_month_mom = to_num(raw[[2]][3]),
    order_month_yoy = to_num(raw[[3]][3]),
    conversion_month_value = to_num(raw[[4]][3]),
    conversion_month_mom = to_num(raw[[5]][3]),
    conversion_month_yoy = to_num(raw[[6]][3])
  )
}

read_daily_field_summary <- function(path, target_date) {
  raw <- readxl::read_excel(path, sheet = "各车系当日数据", col_names = FALSE, .name_repair = "minimal")
  target_date <- as.Date(target_date, origin = "1970-01-01")
  l03_orders_daily <- to_num(raw[[23]][6])
  l03_orders <- read_l03_metrics(path) %>%
    filter(
      .data$metric == "锁单",
      !is.na(.data$date),
      format(.data$date, "%Y-%m") == format(target_date, "%Y-%m"),
      .data$date <= target_date
    )
  order_conversion_summary <- read_order_conversion_summary(path)
  order_daily_value <- read_order_metrics_from_data_bottom(path) %>%
    filter(.data$scope == "全系", .data$metric == "锁单", .data$date == target_date) %>%
    slice(1) %>%
    pull(.data$value)
  order_month_value <- order_conversion_summary$order_month_value
  order_month_mom <- order_conversion_summary$order_month_mom
  order_month_yoy <- order_conversion_summary$order_month_yoy

  tibble(
    metric = c("客流", "线索", "试驾", "锁单", "转化率"),
    daily_value = c(to_num(raw[[2]][4]), to_num(raw[[4]][4]), to_num(raw[[13]][4]), zero_if_na(order_daily_value), NA_real_),
    month_value = c(to_num(raw[[2]][14]), to_num(raw[[4]][14]), to_num(raw[[13]][14]), order_month_value, order_conversion_summary$conversion_month_value),
    month_mom = c(to_num(raw[[2]][17]), to_num(raw[[4]][17]), to_num(raw[[13]][17]), order_month_mom, order_conversion_summary$conversion_month_mom),
    month_yoy = c(to_num(raw[[2]][60]), to_num(raw[[4]][60]), to_num(raw[[13]][60]), order_month_yoy, order_conversion_summary$conversion_month_yoy)
  )
}

build_model_text <- function(daily_summary, target_date) {
  target_date <- as.Date(target_date, origin = "1970-01-01")
  day_label <- as.integer(format(target_date, "%d"))
  row_for <- function(metric_name) daily_summary %>% filter(.data$metric == metric_name) %>% slice(1)
  traffic <- row_for("客流")
  leads <- row_for("线索")
  drives <- row_for("试驾")
  orders <- row_for("锁单")
  conversion <- row_for("转化率")

  glue(
    "客流：{day_label}日客流{fmt_num(traffic$daily_value)}，月累客流{fmt_num(traffic$month_value)}，月累环比{fmt_pct(traffic$month_mom, FALSE)}，月累同比{fmt_pct(traffic$month_yoy, FALSE)}\n",
    "线索：{day_label}日线索{fmt_num(leads$daily_value)}，月累线索{fmt_num(leads$month_value)}，月累环比{fmt_pct(leads$month_mom, FALSE)}，月累同比{fmt_pct(leads$month_yoy, FALSE)}\n",
    "试驾：{day_label}日试驾{fmt_num(drives$daily_value)}，月累试驾{fmt_num(drives$month_value)}，月累环比{fmt_pct(drives$month_mom, FALSE)}，月累同比{fmt_pct(drives$month_yoy, FALSE)}\n",
    "锁单：{day_label}日锁单{fmt_num(orders$daily_value)}，月累锁单{fmt_num(orders$month_value)}，月累环比{fmt_pct(orders$month_mom, FALSE)}，月累同比{fmt_pct(orders$month_yoy, FALSE)}\n",
    "转化率：月累转化率{fmt_pct_value(conversion$month_value)}，月累环比{fmt_pct(conversion$month_mom, FALSE)}，月累同比{fmt_pct(conversion$month_yoy, FALSE)}"
  )
}

main <- function() {
  if (!file.exists(input_path)) stop("Input workbook does not exist: ", input_path, call. = FALSE)

  rule <- get_compare_rule(target_date)
  compare_label <- get_compare_label(rule)
  date_sets <- get_date_sets(target_date, rule)

  sales_data <- bind_rows(read_sales_trend(input_path), read_order_metrics_from_data_bottom(input_path))
  online_data <- read_online_official_traffic(input_path)
  l03_data <- read_l03_metrics(input_path)

  if (!target_date %in% sales_data$date) {
    available_dates <- sort(unique(sales_data$date[!is.na(sales_data$date)]))
    stop(
      "Target date ", as.character(target_date), " was not found in 销售过程数据走势. Available date range: ",
      as.character(min(available_dates)), " to ", as.character(max(available_dates)),
      call. = FALSE
    )
  }
  if (!target_date %in% online_data$date) {
    available_dates <- sort(unique(online_data$date[!is.na(online_data$date)]))
    stop(
      "Target date ", as.character(target_date), " was not found in 线上官渠潜客客流. Available date range: ",
      as.character(min(available_dates)), " to ", as.character(max(available_dates)),
      call. = FALSE
    )
  }
  if (!target_date %in% l03_data$date) {
    available_dates <- sort(unique(l03_data$date[!is.na(l03_data$date)]))
    stop(
      "Target date ", as.character(target_date), " was not found in 销售过程数据走势/L03 metrics. Available date range: ",
      as.character(min(available_dates)), " to ", as.character(max(available_dates)),
      call. = FALSE
    )
  }

  calc_table <- build_calc_table(sales_data, online_data, l03_data, date_sets, rule)
  daily_summary <- read_daily_field_summary(input_path, target_date)

  overview_content <- paste0(as.character(build_overview_text(calc_table, compare_label)), collapse = "")
  model_content <- paste0(as.character(build_model_text(daily_summary, target_date)), collapse = "")

  result <- tibble(
    field = c("概览解析-AI颜色版", "车系数据解析-AI颜色版"),
    target_date = as.character(target_date),
    compare_rule = rule,
    compare_label = compare_label,
    content = c(overview_content, model_content)
  )

  dir.create(dirname(output_path), recursive = TRUE, showWarnings = FALSE)
  readr::write_excel_csv(result, output_path)

  cat("Target date: ", as.character(target_date), "\n", sep = "")
  cat("Compare rule: ", rule, " (VS ", compare_label, ")\n", sep = "")
  cat("Output: ", output_path, "\n\n", sep = "")
  cat("概览解析-AI颜色版\n")
  cat(result$content[[1]], "\n\n", sep = "")
  cat("车系数据解析-AI颜色版\n")
  cat(result$content[[2]], "\n", sep = "")
}

main()
