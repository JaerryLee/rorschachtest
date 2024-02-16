from django.contrib import admin
from django.urls import path, include
from config.views import greeting, about, register, plan, main, privacy, service
from scoring import views
from django.contrib.auth.decorators import login_required
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('admin/', admin.site.urls),
    path('accounts/', include('accounts.urls')),
    path('', main, name='main'),
    path('greeting/', greeting, name='greeting'),
    path('about/', about, name='about'),
    path('register/', register, name='register'),
    path('privacy/', privacy, name='privacy_policy'),
    path('service/', service, name='service'),
    path('plan/', plan, name='plan'),
    path('client_info/', login_required(views.add_client), name='client_info'),
    path('search/', login_required(views.search), name='search'),
    path('django_plotly_dash/', include('django_plotly_dash.urls')),
    path('search_results/', views.search_results),
    path('clients/', views.client_list, name='client_list'),
    path('clients/<int:client_id>/', views.client_detail, name='client_detail'),
    path('export_structural_summary/<int:client_id>/', views.export_structural_summary_xlsx,
         name='export_structural_summary_xlsx'),
    path('board/', include('board.urls', namespace='board')),
    path('edit_responses/<str:client_id>/', views.edit_responses, name='edit_responses'),
    path('update-response-codes/<int:client_id>/', views.update_response_codes, name='update_response_codes'),
]

urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)