from django.db import models
from django.contrib.auth.models import AbstractUser




# Create your models here.
class Files(models.Model):
    time = models.CharField(max_length=100,blank=False, null=False)
    documents = models.FileField(upload_to="processed_files/", null=True,blank=True)
