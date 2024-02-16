from django.contrib import admin
from import_export import resources
from import_export.admin import ImportExportModelAdmin
from .models import SearchReference, CardImages, PopularResponse, Client, ResponseCode, StructuralSummary


class SearchReferenceResource(resources.ModelResource):
    class Meta:
        model = SearchReference
        fields = ('id', 'Card', 'LOC', 'Cont', 'FQ', 'Determinants', 'Item', 'V')
        export_order = fields


class SearchReferenceAdmin(ImportExportModelAdmin):
    fields = ('Card', 'LOC', 'Cont', 'FQ', 'Determinants', 'Item', 'V')
    list_display = ('id', 'Card', 'LOC', 'Cont', 'FQ', 'Determinants', 'Item', 'V')
    resource_class = SearchReferenceResource


class ResponseCodeResource(resources.ModelResource):
    class Meta:
        model = ResponseCode
        fields = ('client', 'card', 'response_num', 'time', 'response', 'rotation', 'inquiry', 'location', 'loc_num',
                  'dev_qual', 'determinants', 'pair', 'form_qual', 'content', 'popular', 'Z', 'special', 'comment')
        export_order = fields


class ResponseCodeAdmin(ImportExportModelAdmin):
    fields = ('client', 'card', 'response_num', 'time', 'response', 'rotation', 'inquiry', 'location', 'loc_num',
              'dev_qual', 'determinants', 'pair', 'form_qual', 'content', 'popular', 'Z', 'special', 'comment')
    list_display = ('get_name', 'card', 'response_num')

    def get_name(self, obj):
        return obj.client.name

    get_name.admin_order_field = 'client'
    get_name.short_description = 'Client Name'
    resource_class = ResponseCodeResource


class StructuralSummaryResource(resources.ModelResource):
    class Meta:
        model = StructuralSummary
        exclude = ()


class StructuralSummaryAdmin(ImportExportModelAdmin):
    resource_class = StructuralSummaryResource


class CardImagesResource(resources.ModelResource):
    class Meta:
        model = CardImages
        fields = ('id', 'card_number', 'section', 'img_file', 'detail_img')
        export_order = fields


class CardImagesAdmin(ImportExportModelAdmin):
    resource_class = CardImagesResource
    list_display = ('id', 'card_number', 'section', 'img_file', 'detail_img')
    search_fields = ('card_number', 'section')
    list_filter = ('card_number', 'section')


admin.site.register(CardImages, CardImagesAdmin)

admin.site.register(SearchReference, SearchReferenceAdmin)
admin.site.register(Client)
admin.site.register(ResponseCode, ResponseCodeAdmin)
admin.site.register(StructuralSummary, StructuralSummaryAdmin)


class PopularResponseResource(resources.ModelResource):
    class Meta:
        model = PopularResponse
        fields = ('id', 'card_number', 'p', 'Z')
        export_order = fields


class PopularResponseAdmin(ImportExportModelAdmin):
    fields = ('card_number', 'p', 'Z')
    list_display = ('id', 'card_number', 'p', 'Z')
    resource_class = PopularResponseResource


admin.site.register(PopularResponse, PopularResponseAdmin)
