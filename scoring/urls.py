from django.urls import path
from . import views

app_name = 'scoring'

urlpatterns = [
    path('client_info/', views.add_client, name='client_info'),
    path('clients/', views.client_list, name='client_list'),
    path('clients/<int:client_id>/', views.client_detail, name='client_detail'),

    path('search/<int:client_id>/', views.search, name='search'),
    path('responses/<int:client_id>/update/', views.update_response_codes, name='update_response_codes'),
    path('search/results/', views.search_results, name='search_results'),

    path(
        'clients/<int:client_id>/export-structural-summary.xlsx',
        views.export_structural_summary_xlsx,
        name='export_structural_summary_xlsx',
    ),
    path(
        'templates/response/intermediate.xlsx',
        views.download_response_template_intermediate,
        name='download_response_template_intermediate',
    ),

    path('advanced/', views.advanced_entry, name='advanced_entry'),
    path('advanced/<int:client_id>/upload/', views.advanced_upload, name='advanced_upload'),
    path('advanced/<int:client_id>/edit/', views.advanced_edit_responses, name='advanced_edit_responses'),
    path(
        'advanced/<int:client_id>/summary.xlsx',
        views.export_structural_summary_xlsx_advanced,
        name='export_structural_summary_xlsx_advanced',
    ),
    path(
        'templates/response/advanced.xlsx',
        views.download_response_template_advanced,
        name='download_response_template_advanced',
    ),
]
