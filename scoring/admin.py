from django.contrib import admin
from django.utils.html import format_html
from django.urls import reverse, NoReverseMatch
from import_export import resources
from import_export.admin import ImportExportModelAdmin
from .models import (
    SearchReference,
    CardImages,
    PopularResponse,
    Client,
    ResponseCode,
    StructuralSummary,
)


class SearchReferenceResource(resources.ModelResource):
    class Meta:
        model = SearchReference
        fields = ("id", "Card", "LOC", "Cont", "FQ", "Determinants", "Item", "V")
        export_order = fields


class ResponseCodeResource(resources.ModelResource):
    class Meta:
        model = ResponseCode
        fields = (
            "client",
            "card",
            "response_num",
            "time",
            "response",
            "rotation",
            "inquiry",
            "location",
            "loc_num",
            "dev_qual",
            "determinants",
            "pair",
            "form_qual",
            "content",
            "popular",
            "Z",
            "special",
            "comment",
        )
        export_order = fields


class StructuralSummaryResource(resources.ModelResource):
    class Meta:
        model = StructuralSummary
        exclude = ()


class CardImagesResource(resources.ModelResource):
    class Meta:
        model = CardImages
        fields = ("id", "card_number", "section", "img_file", "detail_img")
        export_order = fields


class PopularResponseResource(resources.ModelResource):
    class Meta:
        model = PopularResponse
        fields = ("id", "card_number", "p", "Z")
        export_order = fields


@admin.register(SearchReference)
class SearchReferenceAdmin(ImportExportModelAdmin):
    resource_class = SearchReferenceResource
    fields = ("Card", "LOC", "Cont", "FQ", "Determinants", "Item", "V")
    list_display = ("id", "Card", "LOC", "Cont", "FQ", "Determinants", "Item", "V")
    search_fields = ("Card", "LOC", "Cont", "Determinants", "Item")
    list_filter = ("Card", "FQ", "V")
    ordering = ("Card", "LOC", "Cont")


@admin.register(ResponseCode)
class ResponseCodeAdmin(ImportExportModelAdmin):
    resource_class = ResponseCodeResource
    list_select_related = ("client",)

    fields = (
        "client",
        "card",
        "response_num",
        "time",
        "response",
        "rotation",
        "inquiry",
        "location",
        "loc_num",
        "dev_qual",
        "determinants",
        "pair",
        "form_qual",
        "content",
        "popular",
        "Z",
        "special",
        "comment",
    )
    list_display = (
        "get_client_name",
        "card",
        "response_num",
        "form_qual",
        "popular",
        "Z",
        "short_response",
    )
    search_fields = (
        "client__name",
        "response",
        "inquiry",
        "determinants",
        "content",
        "special",
        "card",
    )
    list_filter = ("card", "form_qual", "popular", "Z", "client")
    ordering = ("client__name", "card", "response_num")

    @admin.display(ordering="client__name", description="Client")
    def get_client_name(self, obj):
        return obj.client.name if obj.client_id else "-"

    @admin.display(description="Response (short)")
    def short_response(self, obj):
        txt = (obj.response or "").strip()
        return txt if len(txt) <= 30 else f"{txt[:30]}…"


@admin.register(StructuralSummary)
class StructuralSummaryAdmin(ImportExportModelAdmin):
    resource_class = StructuralSummaryResource
    list_select_related = ("client",)

    list_display = (
        "client",
        "R",
        "L",
        "Zf",
        "Zsum",
        "Zd",
        "popular",
        "sum6",
        "wsum6",
        "PTI",
        "DEPI",
        "CDI",
        "SCON",
        "HVI",
        "OBS_posi",
    )
    search_fields = ("client__name",)
    list_filter = ("OBS_posi",)
    ordering = ("client__name",)

    readonly_fields = tuple(
        f.name for f in StructuralSummary._meta.fields if f.name not in ("id", "client")
    )

    actions = ["recalculate_selected"]

    @admin.action(description="선택한 항목 재계산")
    def recalculate_selected(self, request, queryset):
        
        for ss in queryset:
            ss.calculate_values()
            ss.save()


@admin.register(CardImages)
class CardImagesAdmin(ImportExportModelAdmin):
    resource_class = CardImagesResource
    list_display = ("id", "card_number", "section", "img_thumb", "detail_thumb")
    search_fields = ("card_number", "section")
    list_filter = ("card_number", "section")

    @admin.display(description="Image")
    def img_thumb(self, obj):
        try:
            if obj.img_file and hasattr(obj.img_file, "url"):
                return format_html('<img src="{}" style="height:40px;">', obj.img_file.url)
        except Exception:
            pass
        return "-"

    @admin.display(description="Detail")
    def detail_thumb(self, obj):
        try:
            if obj.detail_img and hasattr(obj.detail_img, "url"):
                return format_html('<img src="{}" style="height:40px;">', obj.detail_img.url)
        except Exception:
            pass
        return "-"


@admin.register(PopularResponse)
class PopularResponseAdmin(ImportExportModelAdmin):
    resource_class = PopularResponseResource
    fields = ("card_number", "p", "Z")
    list_display = ("id", "card_number", "p", "Z")
    list_filter = ("card_number",)
    search_fields = ("card_number", "p", "Z")
    ordering = ("card_number", "id")


class ResponseCodeInline(admin.TabularInline):
    model = ResponseCode
    extra = 0
    fields = (
        "card",
        "response_num",
        "time",
        "response",
        "inquiry",
        "location",
        "dev_qual",
        "determinants",
        "form_qual",
        "popular",
        "Z",
    )
    show_change_link = True


@admin.register(Client)
class ClientAdmin(admin.ModelAdmin):
    list_display = (
        "name",
        "tester",
        "gender",
        "birthdate",
        "testDate",
        "age",
        "consent",
        "responses_count",
        "frontend_links",
    )
    list_filter = ("gender", "consent", "tester")
    search_fields = ("name", "tester__username")
    ordering = ("-testDate", "name")
    inlines = [ResponseCodeInline]

    @admin.display(description="Responses")
    def responses_count(self, obj):
        return obj.responses.count()

    @admin.display(description="Open (Front)")
    def frontend_links(self, obj):
        links = []
        try:
            url_intermediate = f"{reverse('search')}?client_id={obj.id}"
            links.append(f'<a target="_blank" href="{url_intermediate}">중급</a>')
        except NoReverseMatch:
            pass

        try:
            url_advanced = f"{reverse('advanced_upload')}?client_id={obj.id}"
            links.append(f'<a target="_blank" href="{url_advanced}">고급</a>')
        except NoReverseMatch:
            pass

        return format_html(" | ".join(links)) if links else "-"
