from django.urls import path
from . import views

urlpatterns = [
    path('client_info/', views.add_client, name='client_info'),
]