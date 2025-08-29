from django.contrib.auth.decorators import login_required
from django.http import HttpResponseForbidden
from django.shortcuts import get_object_or_404

from ..models import Client

from .intermediate import export_structural_summary_xlsx as export_intermediate
from .advanced import export_structural_summary_xlsx_advanced as export_advanced


@login_required
def export_structural_summary_xlsx_auto(request, client_id):

    client = get_object_or_404(Client, id=client_id)
    if client.tester != request.user:
        return HttpResponseForbidden("액세스 거부: 해당 정보를 볼 수 있는 권한이 없습니다.")

    user_group = getattr(request.user, "group", "intermediate")
    if user_group == "advanced":
        return export_advanced(request, client_id)
    else:
        return export_intermediate(request, client_id)
