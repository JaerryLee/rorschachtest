from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.http import HttpResponseForbidden
from django.shortcuts import get_object_or_404, redirect

from ..models import Client

@login_required
def export_structural_summary_xlsx_auto(request, client_id):
    client = get_object_or_404(Client, id=client_id)
    if client.tester != request.user:
        return HttpResponseForbidden("액세스 거부: 해당 정보를 볼 수 있는 권한이 없습니다.")

    messages.info(
        request,
        "엑셀 자동 라우터는 더 이상 지원하지 않습니다. "
        "상단의 ‘중급 요약 다운로드’ 또는 ‘고급 요약 다운로드’를 사용해 주세요."
    )
    return redirect('scoring:client_detail', client_id=client.id)