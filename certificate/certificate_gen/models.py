from django.db import models
from django.contrib.auth.models import AbstractUser


# Create your models here.

class Signs(models.Model):
    Signatories_name = models.CharField(max_length=100,blank=True, null=True)
    Signatories_no = models.CharField(max_length=15,null=True,blank=True)
    digital_Signatures = models.CharField(max_length=15,null=True,blank=True)
    certificate_type = models.CharField(primary_key=True)

class User(AbstractUser):
    Name  = models.CharField(max_length=150,null=True,blank=True)
    Registration_number = models.CharField(max_length=150,null=True,blank=True)
    Category  = models.CharField(max_length=150,null=True,blank=True)
    Year_of_study =models.CharField(max_length=30,null=True,blank=True)
    Course = models.CharField(max_length=100, null=True,blank=True)
    institution = models.CharField(max_length=100,null=True,blank=True)
    department = models.CharField(max_length=100, null=True,blank=True)
    location = models.CharField(max_length=15,null=True,blank=True)
    result_rank = models.ForeignKey(Signs,on_delete=models.CASCADE)
    event_name = models.CharField(max_length=100,blank=True, null=True)

    def __str__(self):
        return str(self.Name)

class Event_details(models.Model):
    event_name = models.ForeignKey(User,on_delete=models.CASCADE,blank=True, null=True,related_name='event_name')
    event_date = models.CharField(max_length=15,null=True,blank=True)
    organizer = models.CharField(max_length=15,null=True,blank=True)









