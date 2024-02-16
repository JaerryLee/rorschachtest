from django.contrib import admin
from .models import Post, Comment, Notice

admin.site.register([Post, Comment, Notice])
