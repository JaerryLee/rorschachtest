from django.urls import path
from . import views

app_name = 'scoring'

urlpatterns = [
    path('client_info/', views.add_client, name='client_info'),

    # 검색
    path('search/<int:client_id>/', views.search, name='search'),
    # 편집
    path('responses/<int:client_id>/edit/', views.edit_responses, name='edit_responses'),
    path('responses/<int:client_id>/update/', views.update_response_codes, name='update_response_codes'),
    path('search/results/', views.search_results, name='search_results'),

    # 템플릿
    path('responses/template/', views.download_response_template, name='download_response_template'),

    # 고급 업로드
    path('advanced/<int:client_id>/upload/', views.advanced_upload, name='advanced_upload'),

    path('advanced/<int:client_id>/summary.xlsx',
        views.export_structural_summary_xlsx_advanced,
        name='export_structural_summary_xlsx_advanced'
    ),
    # 수검자 관리/보고서
    path('clients/', views.client_list, name='client_list'),
    path('clients/<int:client_id>/', views.client_detail, name='client_detail'),
    path('clients/<int:client_id>/export-structural-summary.xlsx',
         views.export_structural_summary_xlsx, name='export_structural_summary_xlsx'),
    path('clients/<int:client_id>/summary.xlsx',
         views.export_structural_summary_xlsx_auto,
         name='export_structural_summary_xlsx_auto'),
    
]
