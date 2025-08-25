from django.urls import path
from . import views

app_name = "board"

urlpatterns = [
    path('beginner_board/', views.beginner_board, name='beginner_board'),
    path('intermediate_board/', views.intermediate_board, name='intermediate_board'),
    path('advanced_board/', views.advanced_board, name='advanced_board'),
    path('post/<int:post_id>/', views.post_detail, name='post_detail'),
    path('create_post/<str:group>/', views.create_post, name='create_post'),
    path('notice/', views.notice, name='notice'),
    path('notice/<int:notice_id>/', views.notice_detail, name='notice_detail'),
    path('delete_post/<int:post_id>', views.delete_post, name='delete_post'),
]
