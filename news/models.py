from django.db import models

class Student(models.Model):
    full_name = models.CharField(max_length=70)

    class Sex(models.IntegerChoices):
        MALE = 1, '男'
        FEMALE = 2, '女'
        OTHER = 3,'其他'
    sex = models.IntegerField(choices = Sex.choices)
    #设置性别选项



    def __str__(self):
        return self.full_name

class Homework(models.Model):
    commmit_date = models.DateField(auto_now=True)
    headline = models.CharField(max_length=200)
    attach = models.FileField()
    remark = models.TextField()
    student = models.ForeignKey(Student, on_delete=models.CASCADE, null = True)
    #设置作业选项

    def __str__(self):
        return self.headline

