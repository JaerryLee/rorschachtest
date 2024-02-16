from django.db import models
from django.core.validators import RegexValidator
from django.contrib.auth.models import AbstractUser


# Create your models here.
class User(AbstractUser):
    phoneNumberRegex = RegexValidator(regex=r'^01([0|1|6|7|8|9]?)-?([0-9]{3,4})-?([0-9]{4})$')
    phone = models.CharField(validators=[phoneNumberRegex], max_length=11, unique=True)
    group = models.CharField(max_length=20, choices=[('beginner', '초급'), ('intermediate', '중급')],
                             default='beginner')
    consent = models.BooleanField(default=False)
