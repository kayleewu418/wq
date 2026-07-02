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

trend_word <- function(x, up = "有上升", down = "有下降", flat = "基本持平") {
  if (is.na(x)) return("变化")
  if (x > 0.001) return(up)
  if (x < -0.001) return(down)
  flat
}

safe_rate <- function(current, baseline) {
  if (is.na(current) || is.na(baseline) || baseline == 0) return(NA_real_)
  current / baseline - 1
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
    w == 6 ~ "last_saturday",
    w == 0 ~ "weekly_vs_last_week",
    TRUE ~ "previous_day"
  )
}

get_compare_label <- function(rule) {
  dplyr::case_when(
    rule == "prev_week_weekday_avg" ~ "上一周周中日均",
    rule == "previous_day" ~ "前一日",
    rule == "last_saturday" ~ "上周六",
    rule == "weekly_vs_last_week" ~ "上周",
    TRUE ~ "基准"
  )
}

get_date_sets <- function(date, rule) {
  if (rule == "prev_week_weekday_avg") {
    list(current = date, baseline = seq(date - 7, date - 3, by = "day"))
  } else if (rule == "previous_day") {
    list(current = date, baseline = date - 1)
  } else if (rule == "last_saturday") {
    list(current = date, baseline = date - 7)
  } else if (rule == "weekly_vs_last_week") {
    list(current = seq(date - 6, date, by = "day"), baseline = seq(date - 13, date - 7, by = "day"))
  } else {
    list(current = date, baseline = date - 1)
  }
}

summarise_value <- function(data, scope_name, metric_name, dates, agg = c("sum", "mean")) {
  agg <- match.arg(agg)
  values <- data %>%
    filter(.data$scope == scope_name, .data$metric == metric_name, .data$date %in% dates) %>%
    pull(.data$value)
  if (length(values) == 0 || all(is.na(values))) return(NA_real_)
  if (agg == "mean") mean(values, na.rm = TRUE) else sum(values, na.rm = TRUE)
}

calc_metric_compare <- function(data, scope_name, metric_name, date_sets, rule) {
  baseline_agg <- ifelse(rule == "prev_week_weekday_avg", "mean", "sum")
  current <- summarise_value(data, scope_name, metric_name, date_sets$current, "sum")
  baseline <- summarise_value(data, scope_name, metric_name, date_sets$baseline, baseline_agg)
  tibble(scope = scope_name, metric = metric_name, current = current, baseline = baseline, rate = safe_rate(current, baseline))
}

read_sales_trend <- function(path) {
  raw <- readxl::read_excel(path, sheet = "销售过程数据走势", col_names = FALSE, .name_repair = "minimal")
  blocks <- tibble::tribble(
    ~scope, ~date_row, ~metric_start, ~metric_end,
    "全系", 2L, 3L, 8L,
    "P7", 14L, 15L, 17L,
    "G7", 20L, 21L, 23L,
    "M03", 26L, 27L, 29L,
    "P7+", 32L, 33L, 35L,
    "G6", 38L, 39L, 41L,
    "G9", 44L, 45L, 47L,
    "X9", 50L, 51L, 53L,
    "GX", 56L, 57L, 59L
  )

  bind_rows(lapply(seq_len(nrow(blocks)), function(i) {
    block <- blocks[i, ]
    dates <- as_report_date(unlist(raw[block$date_row, -1], use.names = FALSE))
    metric_rows <- raw[block$metric_start:block$metric_end, , drop = FALSE]
    bind_rows(lapply(seq_len(nrow(metric_rows)), function(j) {
      tibble(
        scope = block$scope,
        metric = as.character(metric_rows[[1]][j]),
        date = dates,
        value = to_num(unlist(metric_rows[j, -1], use.names = FALSE))
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

build_calc_table <- function(sales_data, online_data, date_sets, rule) {
  all_data <- bind_rows(sales_data, online_data)
  base <- bind_rows(
    calc_metric_compare(all_data, "官渠", "线上官渠潜客客流", date_sets, rule),
    calc_metric_compare(all_data, "全系", "客流", date_sets, rule),
    calc_metric_compare(all_data, "全系", "线索", date_sets, rule),
    calc_metric_compare(all_data, "全系", "线上线索", date_sets, rule),
    calc_metric_compare(all_data, "全系", "线下线索", date_sets, rule),
    calc_metric_compare(all_data, "GX", "线索", date_sets, rule),
    calc_metric_compare(all_data, "全系", "试驾", date_sets, rule),
    calc_metric_compare(all_data, "GX", "试驾", date_sets, rule),
    calc_metric_compare(all_data, "全系", "锁单", date_sets, rule),
    calc_metric_compare(all_data, "GX", "锁单", date_sets, rule)
  )

  get_row <- function(scope_name, metric_name) {
    base %>% filter(.data$scope == scope_name, .data$metric == metric_name) %>% slice(1)
  }

  all_leads <- get_row("全系", "线索")
  gx_leads <- get_row("GX", "线索")
  all_drives <- get_row("全系", "试驾")
  gx_drives <- get_row("GX", "试驾")
  all_orders <- get_row("全系", "锁单")
  gx_orders <- get_row("GX", "锁单")

  bind_rows(
    base,
    tibble(scope = "其他车型", metric = "线索", current = all_leads$current - gx_leads$current, baseline = all_leads$baseline - gx_leads$baseline, rate = safe_rate(all_leads$current - gx_leads$current, all_leads$baseline - gx_leads$baseline)),
    tibble(scope = "其他车型", metric = "试驾", current = all_drives$current - gx_drives$current, baseline = all_drives$baseline - gx_drives$baseline, rate = safe_rate(all_drives$current - gx_drives$current, all_drives$baseline - gx_drives$baseline)),
    # Current workbook convention: all-system order row already excludes GX.
    tibble(scope = "其他车型", metric = "锁单", current = all_orders$current, baseline = all_orders$baseline, rate = all_orders$rate),
    tibble(scope = "含GX全系", metric = "锁单", current = all_orders$current + gx_orders$current, baseline = all_orders$baseline + gx_orders$baseline, rate = safe_rate(all_orders$current + gx_orders$current, all_orders$baseline + gx_orders$baseline))
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
  other_leads_rate <- pick_value(calc_table, "其他车型", "线索", "rate")
  drives_rate <- pick_value(calc_table, "全系", "试驾", "rate")
  gx_drives_rate <- pick_value(calc_table, "GX", "试驾", "rate")
  other_drives_rate <- pick_value(calc_table, "其他车型", "试驾", "rate")
  orders_current <- pick_value(calc_table, "其他车型", "锁单", "current")
  orders_rate <- pick_value(calc_table, "其他车型", "锁单", "rate")
  gx_orders_current <- pick_value(calc_table, "GX", "锁单", "current")
  gx_orders_rate <- pick_value(calc_table, "GX", "锁单", "rate")
  other_orders_current <- pick_value(calc_table, "其他车型", "锁单", "current")
  other_orders_rate <- pick_value(calc_table, "其他车型", "锁单", "rate")

  glue(
    "线上官渠潜客客流：官渠客流{trend_word(official_rate)}，VS{compare_label}{fmt_pct(official_rate, FALSE)}；\n",
    "门店客流：门店客流{trend_word(traffic_rate)}，VS {compare_label}{fmt_pct(traffic_rate)}；\n",
    "线索总量：线索总量{trend_word(leads_rate)}，VS {compare_label}{fmt_pct(leads_rate)}，",
    "其中线上线索总量VS {compare_label}{fmt_pct(online_leads_rate)}，",
    "线下线索总量VS {compare_label}{fmt_pct(offline_leads_rate)}；",
    "其中GX线索VS {compare_label}{fmt_pct(gx_leads_rate)}，",
    "其他车型合计VS {compare_label}{fmt_pct(other_leads_rate)}；\n",
    "试驾总量：试驾总量{trend_word(drives_rate)}，VS {compare_label}{fmt_pct(drives_rate)}；",
    "其中GX试驾VS {compare_label}{fmt_pct(gx_drives_rate)}，",
    "其他车型合计VS {compare_label}{fmt_pct(other_drives_rate)}；\n",
    "锁单总量：全系锁单总量{fmt_num(orders_current)}，VS {compare_label}{fmt_pct(orders_rate)}；",
    "其中GX锁单{fmt_num(gx_orders_current)}台，VS {compare_label}{fmt_pct(gx_orders_rate)}，",
    "其他车型合计{fmt_num(other_orders_current)}台，VS {compare_label}{fmt_pct(other_orders_rate)}。"
  )
}

read_daily_field_summary <- function(path) {
  raw <- readxl::read_excel(path, sheet = "各车系当日数据", col_names = FALSE, .name_repair = "minimal")
  tibble(
    metric = c("客流", "线索", "试驾", "锁单", "转化率"),
    daily_value = c(to_num(raw[[2]][4]), to_num(raw[[4]][4]), to_num(raw[[13]][4]), to_num(raw[[22]][4]), NA_real_),
    month_value = c(to_num(raw[[2]][14]), to_num(raw[[4]][14]), to_num(raw[[13]][14]), to_num(raw[[22]][14]), to_num(raw[[33]][14])),
    month_mom = c(to_num(raw[[2]][17]), to_num(raw[[4]][17]), to_num(raw[[13]][17]), to_num(raw[[22]][17]), to_num(raw[[33]][17])),
    month_yoy = c(to_num(raw[[2]][60]), to_num(raw[[4]][60]), to_num(raw[[13]][60]), to_num(raw[[22]][60]), to_num(raw[[33]][60]))
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
    "转化率：月累转化率{fmt_pct(conversion$month_value, FALSE)}，月累环比{fmt_pct(conversion$month_mom, FALSE)}，月累同比{fmt_pct(conversion$month_yoy, FALSE)}"
  )
}

main <- function() {
  if (!file.exists(input_path)) stop("Input workbook does not exist: ", input_path, call. = FALSE)

  rule <- get_compare_rule(target_date)
  compare_label <- get_compare_label(rule)
  date_sets <- get_date_sets(target_date, rule)

  sales_data <- read_sales_trend(input_path)
  online_data <- read_online_official_traffic(input_path)

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

  calc_table <- build_calc_table(sales_data, online_data, date_sets, rule)
  daily_summary <- read_daily_field_summary(input_path)

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
