from django.contrib import admin
from django.urls import path, include
from config.views import greeting, about, register, plan, main, privacy, service
from django.contrib.auth.decorators import login_required
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('admin/', admin.site.urls),
    path('accounts/', include('accounts.urls')),

    # 메인/정적 페이지
    path('', main, name='main'),
    path('greeting/', greeting, name='greeting'),
    path('about/', about, name='about'),
    path('register/', register, name='register'),
    path('privacy/', privacy, name='privacy_policy'),
    path('service/', service, name='service'),
    path('plan/', plan, name='plan'),
    
    path('', include(('scoring.urls', 'scoring'), namespace='scoring')),
    path('board/', include('board.urls', namespace='board')),

    path('django_plotly_dash/', include('django_plotly_dash.urls')),
]

urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
