import io
import zipfile
from urllib.parse import quote

from django.contrib import admin
from django.utils import timezone
from django.utils.html import format_html
from django.urls import reverse, NoReverseMatch, path
from django.http import HttpResponse, Http404
from django.core.exceptions import PermissionDenied
from django.utils.text import slugify
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

from .views.advanced import build_client_xlsx_bytes

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
    change_form_template = "admin/scoring/client/change_form.html"
    
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
    actions = ["export_selected_clients"]
    

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
    def get_urls(self):
        urls = super().get_urls()
        my = [
            path(
                "<path:object_id>/export/",
                self.admin_site.admin_view(self.export_one),   # 권한 체크 포함
                name="scoring_client_export",                  # 템플릿에서 사용한 이름
            ),
        ]
        return my + urls

    def export_one(self, request, object_id):
        obj = self.get_object(request, object_id)
        if obj is None:
            raise Http404("수검자를 찾을 수 없습니다.")
        if not self.has_view_or_change_permission(request, obj):
            raise PermissionDenied("권한이 없습니다.")

        safe_name, data = build_client_xlsx_bytes(obj, include_info_sheet=True)

        fallback = slugify(getattr(obj, "name", "") or "client") + ".xlsx"
        response = HttpResponse(
            data,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response["Content-Disposition"] = (
            f'attachment; filename="{fallback}"; filename*=UTF-8\'\'{quote(safe_name)}'
        )
        return response

    @admin.action(description="선택 항목 엑셀/ZIP 다운로드(관리자)")
    def export_selected_clients(self, request, queryset):
        clients = list(queryset)
        if not clients:
            self.message_user(request, "선택된 항목이 없습니다.")
            return

        if len(clients) == 1:
            client = clients[0]
            if not self.has_view_or_change_permission(request, client):
                raise PermissionDenied("권한이 없습니다.")
            safe_name, data = build_client_xlsx_bytes(client, include_info_sheet=True)
            fallback = slugify(getattr(client, "name", "") or "client") + ".xlsx"
            resp = HttpResponse(
                data,
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            resp["Content-Disposition"] = (
                f'attachment; filename="{fallback}"; filename*=UTF-8\'\'{quote(safe_name)}'
            )
            return resp

        mem = io.BytesIO()
        used_names = set()
        with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for c in clients:
                if not self.has_view_or_change_permission(request, c):
                    continue
                fname, data = build_client_xlsx_bytes(c, include_info_sheet=True)

                base = fname
                i = 1
                while fname in used_names:
                    stem, dot, ext = base.rpartition(".")
                    suffix = f" ({i})"
                    if dot:
                        fname = f"{stem}{suffix}.{ext}"
                    else:
                        fname = f"{base}{suffix}"
                    i += 1
                used_names.add(fname)

                zf.writestr(fname, data)

        mem.seek(0)
        ts = timezone.now().strftime("%Y%m%d_%H%M%S")
        zip_name = f"clients_{ts}.zip"
        resp = HttpResponse(mem.getvalue(), content_type="application/zip")
        resp["Content-Disposition"] = (
            f'attachment; filename="{zip_name}"; filename*=UTF-8\'\'{quote(zip_name)}'
        )
        return resp