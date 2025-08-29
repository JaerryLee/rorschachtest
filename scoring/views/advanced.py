import logging
from pathlib import Path
from functools import wraps
from collections import Counter
from io import BytesIO
import json
import re

import pandas as pd
import numpy as np

from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.forms import modelformset_factory
from django.http import (
    HttpResponse,
    HttpResponseForbidden,
    HttpResponseNotFound,
    JsonResponse,
)
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter

from ..filters import CardImagesFilter, PResponseFilter, SearchReferenceFilter
from ..forms import BulkResponseUploadForm, ResponseCodeForm
from ..models import (
    CardImages,
    Client,
    PopularResponse,
    ResponseCode,
    SearchReference,
    StructuralSummary,
)

from ._base import (
    group_min_required,
    normalize_card_to_num,
    to_roman,
)


TOTAL_CAP = 100
DEFAULT_EXTRA = 40

def _normalize_text_value(v):
    if v is None:
        return ''
    if not isinstance(v, str):
        return v
    s = v
    s = s.replace('\u00A0', ' ').replace('\u200b', '').replace('\u200c', '')
    s = s.strip()
    trans = {
        '“':'"', '”':'"', '‘':"'", '’':"'", '′':"'", '´':"'", '｀':"'",
        '·':'.', 'ㆍ':'.', '‧':'.', '•':'.',
        '：':':', ' ': ' ', '，':',', '．':'.', '／':'/', '－':'-',
    }
    s = ''.join(trans.get(ch, ch) for ch in s)
    return s

def _compact(s: str) -> str:
    s = (s or '').strip()
    s = re.sub(r'[:：]\s*$', '', s)
    return s.replace(' ', '').lower()

HEADER_MAP_COMPACT = {
    'id': None,
    '카드':'card','card':'card',
    'n':'response_num','응답수':'response_num','response_num':'response_num',
    '시간':'time','time':'time',
    '반응':'response','response':'response',
    '질문':'inquiry','inquiry':'inquiry',
    '회전':'rotation','v':'rotation','rotation':'rotation',
    '반응영역':'location','위치':'location','location':'location',
    '발달질':'dev_qual','dq':'dev_qual','dev_qual':'dev_qual','devqual':'dev_qual',
    '영역번호':'loc_num','loc_num':'loc_num','locnum':'loc_num',
    '결정인':'determinants','determinants':'determinants',
    '형태질':'form_qual','형태 질':'form_qual','form_qual':'form_qual','formquality':'form_qual',
    '(2)':'pair','2':'pair','pair':'pair',
    '내용인':'content','내용':'content','content':'content',
    'p':'popular','popular':'popular',
    'z':'Z','Z':'Z',
    '특수점수':'special','special':'special',
    '코멘트':'comment','메모':'comment','comment':'comment',
}

REQUIRED_FIELDS = [
    'card','response_num','time','response','inquiry','rotation','location',
    'dev_qual','loc_num','determinants','form_qual','pair','content','popular','Z','special','comment'
]

def normalize_header(h):
    if h is None:
        return None
    key = _compact(str(h))
    return HEADER_MAP_COMPACT.get(key, None)

def _pick_input_sheet(wb):
    for name in ('입력', 'responses'):
        if name in wb.sheetnames:
            return wb[name]
    for name in wb.sheetnames:
        ws_try = wb[name]
        raw = [(c.value or '') for c in ws_try[1]]
        mapped = [normalize_header(h) for h in raw]
        idx_map = {f: i for i, f in enumerate(mapped) if f}
        if all(f in idx_map for f in REQUIRED_FIELDS):
            return ws_try
    return wb.active

def _row_is_blank(row_tuple):
    return all((v is None) or (str(v).strip() == '') for v in row_tuple)


@group_min_required('advanced')
def advanced_entry(request):
    cid = request.GET.get('client_id') or request.GET.get('client')
    if cid:
        return redirect('scoring:advanced_upload', client_id=cid)
    return redirect(f"{reverse('scoring:client_list')}?next=advanced")


@group_min_required('advanced')
def advanced_upload(request, client_id):
    client = get_object_or_404(Client, id=client_id)
    if client.tester != request.user:
        return HttpResponseForbidden("본인 수검자에게만 업로드할 수 있습니다.")

    existing_count = ResponseCode.objects.filter(client=client).count()
    has_existing = existing_count > 0

    form = BulkResponseUploadForm(request.user, request.POST or None, request.FILES or None)
    form.fields['client'].queryset = Client.objects.filter(id=client.id, tester=request.user)
    if request.method == 'GET' and not form.is_bound:
        form.initial['client'] = client

    if request.method == 'POST' and form.is_valid():
        posted_client = form.cleaned_data['client']
        if posted_client.id != client.id:
            messages.error(request, "요청 경로의 수검자와 폼의 수검자가 일치하지 않습니다.")
            return render(request, 'advanced_upload.html', {
                'client': client, 'form': form,
                'has_existing': has_existing, 'existing_count': existing_count,
            })

        xfile = form.cleaned_data['file']
        replace = form.cleaned_data['replace_existing']

        try:
            wb = load_workbook(filename=xfile, data_only=True)
            ws = _pick_input_sheet(wb)
        except Exception:
            messages.error(request, "엑셀 파일을 열 수 없습니다. (.xlsx 형식 확인)")
            return render(request, 'advanced_upload.html', {
                'client': client, 'form': form,
                'has_existing': has_existing, 'existing_count': existing_count,
            })

        raw_headers = [(c.value or '') for c in ws[1]]
        mapped = [normalize_header(h) for h in raw_headers]
        index_by_field = {f: idx for idx, f in enumerate(mapped) if f}
        missing = [f for f in REQUIRED_FIELDS if f not in index_by_field]
        if missing:
            messages.error(request, "필수 열이 누락되었습니다: " + ", ".join(missing))
            return render(request, 'advanced_upload.html', {
                'client': client, 'form': form,
                'has_existing': has_existing, 'existing_count': existing_count,
            })

        rows_all = list(ws.iter_rows(min_row=2, values_only=True))
        rows = [r for r in rows_all if not _row_is_blank(r)]
        if not rows:
            messages.warning(request, "업로드 가능한 데이터 행이 없습니다.")
            return render(request, 'advanced_upload.html', {
                'client': client, 'form': form,
                'has_existing': has_existing, 'existing_count': existing_count,
            })

        allow = TOTAL_CAP if replace else max(0, TOTAL_CAP - existing_count)
        if allow <= 0:
            messages.error(request, f"이미 {TOTAL_CAP}행이 저장되어 있어 더 추가할 수 없습니다.")
            return redirect('scoring:client_detail', client_id=client.id)

        trimmed = False
        if len(rows) > allow:
            rows = rows[:allow]
            trimmed = True

        created, errors = 0, []
        with transaction.atomic():
            if replace:
                ResponseCode.objects.filter(client=client).delete()

            for ridx, row in enumerate(rows, start=2):
                data = {}
                for f in REQUIRED_FIELDS:
                    idx = index_by_field[f]
                    v = row[idx] if idx < len(row) else ''
                    v = '' if v is None else v
                    v = _normalize_text_value(v)
                    data[f] = v

                data['card'] = to_roman(data.get('card', ''))
                for int_key in ('response_num', 'loc_num'):
                    txt = str(data.get(int_key, '')).strip()
                    if txt == '':
                        data[int_key] = ''
                    else:
                        try:
                            data[int_key] = int(float(txt))
                        except Exception:
                            errors.append(f"{ridx}행: {int_key} 정수 변환 실패")

                def _std_symbols(s):
                    if not s:
                        return s
                    toks = [t.strip() for t in re.split(r'[;,]', str(s)) if t.strip()]
                    toks = [t.replace(' .', '.').replace('. ', '.') for t in toks]
                    seen, out = set(), []
                    for t in toks:
                        if t not in seen:
                            seen.add(t); out.append(t)
                    return ', '.join(out)

                for sym_key in ('determinants','form_qual','content','special'):
                    data[sym_key] = _std_symbols(data.get(sym_key, ''))

                form_row = ResponseCodeForm(data)
                if not form_row.is_valid():
                    errs = '; '.join([f"{fld}: {', '.join(e)}" for fld, e in form_row.errors.items()])
                    errors.append(f"{ridx}행 유효성 오류 → {errs}")
                    continue

                obj = form_row.save(commit=False)
                obj.client = client
                obj.card = to_roman(obj.card)
                obj.save()
                created += 1

        if trimmed:
            messages.warning(request, f"총 {TOTAL_CAP}행 제한으로 앞 {allow}행만 처리했습니다.")
        if errors:
            messages.warning(request, f"성공 {created}건, 오류 {len(errors)}건")
        else:
            messages.success(request, f"성공 {created}건 업로드 완료")

        return redirect('scoring:client_detail', client_id=client.id)

    return render(request, 'advanced_upload.html', {
        'client': client, 'form': form,
        'has_existing': has_existing, 'existing_count': existing_count,
    })


RESOURCE_DIR = Path(getattr(settings, 'SCORING_RESOURCE_DIR',
                            Path(__file__).resolve().parent.parent / 'resources')).resolve()

DEFAULT_RESOURCE_FILENAMES = {
    'symbol_score':   'symbol_score_mapping.json',
    'response_score': 'response_token_scores_mapping.json',
    'inquiry_score':  'inquiry_token_scores_mapping.json',
    'score_stats':    'scoring_area_card_stats.json',
    'index_stats':    'projection_index_card_stats.json',
}
RESOURCE_FILENAMES = getattr(settings, 'SCORING_RESOURCE_FILENAMES', DEFAULT_RESOURCE_FILENAMES)

def _read_json_df(path: Path, required_cols=None) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"필요한 JSON 데이터 파일이 없습니다: {path}")

    def _ensure_required(df: pd.DataFrame) -> pd.DataFrame:
        if not required_cols:
            return df
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            raise ValueError(
                f"{path.name} 에 필요한 열이 없습니다: {', '.join(missing)} "
                f"(사용 가능한 열: {', '.join(map(str, getattr(df, 'columns', [])))} )"
            )
        return df

    for kwargs in ({}, {'lines': True}):
        try:
            df_try = pd.read_json(path, dtype=False, **kwargs)
            if isinstance(df_try, pd.DataFrame) and not df_try.empty:
                rename = {
                    'card':'카드','Card':'카드','card_id':'카드',
                    'token':'토큰','Token':'토큰',
                    'pos':'품사','POS':'품사',
                    'score':'점수','Score':'점수'
                }
                df_try = df_try.rename(columns=rename)
                return _ensure_required(df_try)
        except Exception:
            pass

    raw = json.loads(path.read_text(encoding='utf-8'))

    if isinstance(raw, dict):
        rows = []

        # 카드 → {mean,std}
        if all(isinstance(v, dict) and set(v.keys()) & {'mean', 'std'} for v in raw.values()):
            for card, obj in raw.items():
                rows.append({'카드': str(card), 'mean': obj.get('mean'), 'std': obj.get('std')})
            return _ensure_required(pd.DataFrame(rows))

        # 카드 → 영역 → {mean,std}
        is_area_stats = any(
            isinstance(areas, dict) and any(isinstance(v, dict) and {'mean','std'} <= set(v.keys())
                                            for v in areas.values())
            for areas in raw.values()
        )
        if is_area_stats:
            for card, areas in raw.items():
                if not isinstance(areas, dict): 
                    continue
                for area_name, stat in areas.items():
                    if isinstance(stat, dict) and {'mean','std'} <= set(stat.keys()):
                        rows.append({'카드': str(card), '채점영역': str(area_name),
                                     'mean': stat.get('mean'), 'std': stat.get('std')})
            return _ensure_required(pd.DataFrame(rows))

        # 카드 → 영역 → {기호: 점수}
        is_symbol_scores = any(isinstance(areas, dict) and any(isinstance(v, dict) for v in areas.values())
                               for areas in raw.values())
        if is_symbol_scores:
            for card, areas in raw.items():
                if not isinstance(areas, dict): 
                    continue
                for area_name, symbols in areas.items():
                    if isinstance(symbols, dict):
                        for sym, sc in symbols.items():
                            rows.append({'카드': str(card), '채점영역': str(area_name),
                                         '기호': str(sym), '점수': sc})
            return _ensure_required(pd.DataFrame(rows))

        for key in ('records','data','rows'):
            if isinstance(raw.get(key), list):
                df = pd.DataFrame(raw[key])
                return _ensure_required(df)

        raise ValueError(f"{path.name} 을(를) DataFrame으로 변환할 수 없습니다.")

    if isinstance(raw, list):
        df = pd.DataFrame(raw)
        rename = {
            'card':'카드','Card':'카드','card_id':'카드',
            'token':'토큰','Token':'토큰',
            'pos':'품사','POS':'품사',
            'score':'점수','Score':'점수'
        }
        df = df.rename(columns=rename)
        return _ensure_required(df)

    raise ValueError(f"{path.name} 을(를) DataFrame으로 변환할 수 없습니다.")


STOPWORDS = set([
    '이','그','저','나','너','그것','이것','저것','들','\n','때','것','그리고','하지만','또는','즉','그렇지','그래서','그러므로',
    '대해','대하여','위해','때문에','그런데','근데','이런','저런','그런','같은','처럼','듯','도','만','또','조차','까지',
    '네','예','수검자','검사자','있다','보이다','Q','A','반응반복','같다','하다','반응','반복','부분','여기','이렇게','거','그렇다','어떻다','얘','보다'
])

try:
    from konlpy.tag import Okt
    _OKT = Okt()

    def _preprocess_text(text: str) -> str:
        text = re.sub(r'[^\w\sㄱ-ㅎㅏ-ㅣ가-힣]', ' ', str(text))
        text = re.sub(r'\s+', ' ', text).strip()
        return text

    def tokenize_with_pos(text: str):
        text = _preprocess_text(text)
        target_pos = {'Noun','Verb','Adjective','Adverb'}
        return [(w, p) for (w, p) in _OKT.pos(text, stem=True) if (p in target_pos and w not in STOPWORDS)]
except Exception:
    def _preprocess_text(text: str) -> str:
        text = re.sub(r'[^\w\sㄱ-ㅎㅏ-ㅣ가-힣]', ' ', str(text))
        text = re.sub(r'\s+', ' ', text).strip()
        return text
    def tokenize_with_pos(text: str):
        text = _preprocess_text(text)
        toks = [t for t in re.split(r'\s+', text) if t and t not in STOPWORDS]
        toks = [t for t in toks if len(t) >= 2]
        return [(t, 'Noun') for t in toks]


def _apply_pair_into_determinants(row):
    det = str(row.get('결정인') or '').strip()
    val2 = row.get('(2)')
    val2 = (str(int(val2)).strip() if (val2 not in (None, '', ' ') and not pd.isna(val2)) else '')
    det_tokens = [x.strip() for x in re.split(r'[,.]', det) if x.strip()]
    if val2 and val2 not in det_tokens:
        det_tokens.append(val2)
    row['결정인'] = ', '.join(det_tokens) if det_tokens else np.nan
    return row

def _calculate_token_score(df, token_freq_df, token_column_name, score_column_name):
    tf = token_freq_df.copy()
    tf['토큰_튜플'] = list(zip(tf['토큰'], tf['품사']))
    token_score_map = tf.set_index(['카드', '토큰_튜플'])['점수'].to_dict()

    def compute_row(row):
        card = str(row['카드'])
        tokens = row.get(token_column_name) or []
        uniq = set(tokens)
        total = 0
        for tok in uniq:
            total += token_score_map.get((card, tok), 2)
        return total

    df[score_column_name] = df.apply(compute_row, axis=1)
    return df

def _apply_symbol_score(df, score_table_df):
    areas = ['결정인', '내용인', '특수점수']
    score_dict = {col: [] for col in areas}
    lookup = score_table_df.set_index(['카드','채점영역','기호'])['점수'].to_dict()

    for _, row in df.iterrows():
        card = str(row['카드'])
        for area in areas:
            score = 0
            cell = row.get(area)
            if pd.isna(cell) or cell in (None, ''):
                score_dict[area].append(0)
                continue
            symbols = [s.strip() for s in str(cell).split(',') if s.strip()]
            for sym in symbols:
                score += lookup.get((card, area, sym), 2)
            score_dict[area].append(score)

    for area in areas:
        df[f'{area}_점수'] = score_dict[area]
    return df


@group_min_required('advanced')
def export_structural_summary_xlsx_advanced(request, client_id):
    try:
        client = Client.objects.get(id=client_id)
        if client.tester != request.user:
            return HttpResponse("액세스 거부: 해당 정보를 볼 수 있는 권한이 없습니다.", status=403)

        response_codes = ResponseCode.objects.filter(client_id=client_id)

        numbers_found = {normalize_card_to_num(rc.card) for rc in response_codes if rc.card}
        required = {str(i) for i in range(1, 11)}
        missing = sorted(required - numbers_found, key=int)
        if missing:
            missing_roman = [to_roman(n) for n in missing]
            return HttpResponse("다음 카드의 반응이 없습니다: " + ", ".join(missing_roman))

        structural_summary, created = StructuralSummary.objects.get_or_create(client_id=client_id)
        if not created:
            try:
                structural_summary.calculate_values()
            except Exception:
                structural_summary.save()

    except Client.DoesNotExist:
        logging.error("해당 ID의 클라이언트를 찾을 수 없음")
        return HttpResponseNotFound("클라이언트 정보를 찾을 수 없습니다.")
    except Exception as e:
        logging.error(f"예기치 못한 오류 발생: {e}")
        return JsonResponse({'error': f"{type(e).__name__}: {str(e)}"}, status=500)

    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    ws = wb.create_sheet(title='상단부')
    wsd = wb.create_sheet(title='하단부')
    wsp = wb.create_sheet(title='특수지표')

    ws.merge_cells('A4:B4'); ws.merge_cells('A14:B14'); ws.merge_cells('A20:D20')
    ws.merge_cells('F4:H4'); ws.merge_cells('G5:H5'); ws.merge_cells('J4:K4')
    ws.merge_cells('M4:N4'); ws.merge_cells('M16:P16')
    wsd.merge_cells('A3:F3'); wsd.merge_cells('H3:I3'); wsd.merge_cells('K3:N3')
    wsd.merge_cells('K5:L5'); wsd.merge_cells('K6:L6'); wsd.merge_cells('K7:L7')
    wsd.merge_cells('K8:L8'); wsd.merge_cells('K9:L9'); wsd.merge_cells('K10:L10')
    wsd.merge_cells('K11:L11'); wsd.merge_cells('K12:L12')
    wsd.merge_cells('A14:D14'); wsd.merge_cells('F14:G14'); wsd.merge_cells('I14:J14'); wsd.merge_cells('L14:M14')

    bckground_cells = ['A4', 'A14', 'A20', 'F4', 'F5', 'G5', 'J4', 'M4', 'M16']
    for cell in bckground_cells:
        ws[cell].fill = PatternFill(start_color="7B68EE", end_color="7B68EE", fill_type='solid')
    bckground_cells2 = ['A3', 'H3', 'K3', 'A14', 'F14', 'I14', 'L14']
    for cell in bckground_cells2:
        wsd[cell].fill = PatternFill(start_color="7B68EE", end_color="7B68EE", fill_type='solid')

    ws.sheet_view.showGridLines = False
    wsd.sheet_view.showGridLines = False

    BORDER_LIST = ['A5:B12', 'A4:B4', 'A14:B14', 'A15:B18', 'A20:D20', 'A21:D21', 'A22:D26', 'F4:H4', 'F5:F5', 'G5:H5',
                   'F6:F29', 'G6:H29', 'J4:K4', 'J5:K31', 'M4:N4', 'M5:N14', 'M16:P16', 'M17:P17', 'M18:P25', 'M26:P30']
    BORDER_LIST2 = ['A3:F3', 'A4:F4', 'A5:F7', 'A8:F9', 'H3:I3', 'H4:I10', 'K3:N3', 'K4:N12', 'A14:D14', 'A15:D19',
                    'F14:G14', 'F15:G21', 'I14:J14', 'I15:J21', 'L14:M14', 'L15:M21']

    def set_border(worksheet, cell_range):
        rows = worksheet[cell_range]
        side = Side(border_style='thin', color="FF000000")
        rows = list(rows)
        max_y = len(rows) - 1
        for pos_y, cells in enumerate(rows):
            max_x = len(cells) - 1
            for pos_x, c in enumerate(cells):
                border = Border(left=c.border.left, right=c.border.right, top=c.border.top, bottom=c.border.bottom)
                if pos_x == 0: border.left = side
                if pos_x == max_x: border.right = side
                if pos_y == 0: border.top = side
                if pos_y == max_y: border.bottom = side
                if pos_x == 0 or pos_x == max_x or pos_y == 0 or pos_y == max_y:
                    c.border = border

    for pos in BORDER_LIST:  set_border(ws, pos)
    for pos in BORDER_LIST2: set_border(wsd, pos)

    ws['A4'] = 'Location Features'
    ws['A5'] = 'Zf'; ws['A6'] = 'Zsum'; ws['A7'] = 'Zest'
    ws['A9'] = 'W'; ws['A10'] = 'D'; ws['A11'] = 'Dd'; ws['A12'] = 'S'
    ws['A14'] = 'Developmental Quality'
    ws['A15'] = '+'; ws['A16'] = 'o'; ws['A17'] = 'v/+'; ws['A18'] = 'v'
    ws['A20'] = 'Form Quality'
    ws['B21'] = 'FQx'; ws['C21'] = 'MQual'; ws['D21'] = 'W+D'
    ws['A22'] = '+'; ws['A23'] = 'o'; ws['A24'] = 'u'; ws['A25'] = '-'; ws['A26'] = 'none'

    ws['F4'] = 'Determinants'
    ws['F5'] = 'Blends'; ws['G5'] = 'Single'
    fields = [
        'M','FM','m',"FC","CF","C","Cn","FC'","C'F","C'",
        'FT','TF','T','FV','VF','V','FY','YF','Y','Fr','rF','FD','F','(2)'
    ]
    real_field = [
        'M','FM','m_l','FC','CF','C','Cn','FCa','CaF','Ca',
        'FT','TF','T','FV','VF','V','FY','YF','Y','Fr','rF','FD','F','pair'
    ]
    s_row = 6
    for field_name in fields:
        ws.cell(row=s_row, column=7, value=field_name); s_row += 1

    ws['J4'] = 'Contents'
    cont_fields = [
        'H','(H)','Hd','(Hd)','Hx','A','(A)','Ad','(Ad)','An',
        'Art','Ay','Bl','Bt','Cg','Cl','Ex','Fd','Fi','Ge','Hh','Ls',
        'Na','Sc','Sx','Xy','Id'
    ]
    cont_real_fields = [
        'H','H_paren','Hd','Hd_paren','Hx','A','A_paren','Ad','Ad_paren','An',
        'Art','Ay','Bl','Bt','Cg','Cl','Ex','Fd_l','Fi','Ge','Hh','Ls',
        'Na','Sc','Sx','Xy','Idio'
    ]
    s_row = 5
    for field_name in cont_fields:
        ws.cell(row=s_row, column=10, value=field_name); s_row += 1

    ws['M4'] = "approach"
    for r, label in enumerate(['I','II','III','IV','V','VI','VII','VIII','IX','X'], start=5):
        ws.cell(row=r, column=13, value=label)

    ws['M16'] = 'Special Scores'
    ws['N17'] = 'Lvl-1'; ws['O17'] = 'Lvl-2'
    sp_real_fields = ['sp_dv','sp_inc','sp_dr','sp_fab','sp_alog','sp_con']
    sp_real_fields2 = ['sp_dv2','sp_inc2','sp_dr2','sp_fab2']
    s_row = 18
    for field_name in ['DV','INC','DR','FAB','ALOG','CON']:
        ws.cell(row=s_row, column=13, value=field_name); s_row += 1
    ws['M24'] = 'Raw Sum6'; ws['M25'] = 'Weighted Sum6'
    ws['M26'] = 'AB'; ws['M27'] = 'AG'; ws['M28'] = 'COP'; ws['M29'] = 'CP'
    ws['O26'] = 'GHR'; ws['O27'] = 'PHR'; ws['O28'] = 'MOR'; ws['O29'] = 'PER'; ws['O30'] = 'PSV'

    ws['B5'] = structural_summary.Zf
    ws['B6'] = structural_summary.Zsum; ws['B6'].number_format = '0.0'
    ws['B7'] = structural_summary.Zest; ws['B7'].number_format = '0.0'
    ws['B9'] = structural_summary.W; ws['B10'] = structural_summary.D
    ws['B11'] = structural_summary.Dd; ws['B12'] = structural_summary.S

    ws['B15'] = structural_summary.dev_plus
    ws['B16'] = structural_summary.dev_o
    ws['B17'] = structural_summary.dev_vplus
    ws['B18'] = structural_summary.dev_v

    ws['B22'] = structural_summary.fqx_plus
    ws['B23'] = structural_summary.fqx_o
    ws['B24'] = structural_summary.fqx_u
    ws['B25'] = structural_summary.fqx_minus
    ws['B26'] = structural_summary.fqx_none
    ws['C22'] = structural_summary.mq_plus
    ws['C23'] = structural_summary.mq_o
    ws['C24'] = structural_summary.mq_u
    ws['C25'] = structural_summary.mq_minus
    ws['C26'] = structural_summary.mq_none
    ws['D22'] = structural_summary.wd_plus
    ws['D23'] = structural_summary.wd_o
    ws['D24'] = structural_summary.wd_u
    ws['D25'] = structural_summary.wd_minus
    ws['D26'] = structural_summary.wd_none

    blends = structural_summary.blends.split(',') if structural_summary.blends else []
    start_row = 6
    for blend in blends:
        ws.cell(row=start_row, column=6, value=blend.strip()); start_row += 1

    row = 6
    for field_name in real_field:
        field_value = getattr(structural_summary, field_name)
        ws.cell(row=row, column=8, value=field_value)
        ws.cell(row=row, column=8).number_format = "#"
        row += 1

    row = 5
    for field_name in cont_real_fields:
        field_value = getattr(structural_summary, field_name)
        ws.cell(row=row, column=11, value=field_value)
        ws.cell(row=row, column=11).number_format = "#"
        row += 1

    ws['N5']  = structural_summary.app_I
    ws['N6']  = structural_summary.app_II
    ws['N7']  = structural_summary.app_III
    ws['N8']  = structural_summary.app_IV
    ws['N9']  = structural_summary.app_V
    ws['N10'] = structural_summary.app_VI
    ws['N11'] = structural_summary.app_VII
    ws['N12'] = structural_summary.app_VIII
    ws['N13'] = structural_summary.app_IX
    ws['N14'] = structural_summary.app_X

    lv1_row = 18
    for field_name in sp_real_fields:
        ws.cell(row=lv1_row, column=14, value=getattr(structural_summary, field_name)); lv1_row += 1
    lv2_row = 18
    for field_name in sp_real_fields2:
        ws.cell(row=lv2_row, column=15, value=getattr(structural_summary, field_name)); lv2_row += 1

    ws['N24'] = structural_summary.sum6
    ws['N25'] = structural_summary.wsum6
    ws['N26'] = structural_summary.sp_ab
    ws['N27'] = structural_summary.sp_ag
    ws['N28'] = structural_summary.sp_cop
    ws['N29'] = structural_summary.sp_cp
    ws['P26'] = structural_summary.sp_ghr
    ws['P27'] = structural_summary.sp_phr
    ws['P28'] = structural_summary.sp_mor
    ws['P29'] = structural_summary.sp_per
    ws['P30'] = structural_summary.sp_psv

    for column_cells in ws.columns:
        length = max(len(str(cell.value)) * 1.1 for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = length

    wsd['A3'] = 'Core'
    wsd['A4'] = 'R'; wsd['C4'] = 'L'
    wsd['A5'] = 'EB'; wsd['A6'] = 'eb'
    wsd['C5'] = 'EA'; wsd['C6'] = 'es'; wsd['C7'] = 'Adj es'
    wsd['E5'] = 'EBper'; wsd['E6'] = 'D'; wsd['E7'] = 'Adj D'
    wsd['A8'] = 'FM'; wsd['A9'] = 'm'
    wsd["C8"] = "SumC'"; wsd['C9'] = 'SumV'
    wsd['E8'] = 'SumT'; wsd['E9'] = 'SumY'

    wsd['H3'] = 'Affect'
    wsd['H4'] = 'FC:CF+C'; wsd['H5'] = 'Pure C'; wsd["H6"] = "SumC':WsumC"
    wsd['H7'] = 'Afr'; wsd['H8'] = 'S'; wsd['H9'] = 'Blends:R'; wsd['H10'] = 'CP'

    wsd['K3'] = 'Interpersonal'
    wsd['K4'] = 'COP'; wsd['M4'] = 'AG'
    wsd['K5'] = 'GHR:PHR'; wsd['K6'] = 'a:p'; wsd['K7'] = 'Food'
    wsd['K8'] = 'SumT'; wsd['K9'] = 'Human Content'; wsd['K10'] = 'Pure H'
    wsd['K11'] = 'PER'; wsd['K12'] = 'Isolation Index'

    wsd['A14'] = 'Ideation'
    wsd['A15'] = 'a:p'; wsd['A16'] = 'Ma:Mp'; wsd['A17'] = 'Intel(2AB+Art+Ay)'; wsd['A18'] = 'MOR'
    wsd['C15'] = 'Sum6'; wsd['C16'] = 'Lvl-2'; wsd['C17'] = 'Wsum6'; wsd['C18'] = 'M-'; wsd['C19'] = 'M none'

    wsd['F14'] = 'Mediation'
    med_fields = ['XA%','WDA%','X-%','S-','P','X+%','Xu%']
    med_real_fields = ['xa_per','wda_per','x_minus_per','s_minus','popular','x_plus_per','xu_per']
    s_row = 15
    for field_name in med_fields:
        wsd.cell(row=s_row, column=6, value=field_name); s_row += 1

    wsd['I14'] = 'Processing'
    pro_fields = ['Zf','W:D:Dd','W:M','Zd','PSV','DQ+','DQv']
    pro_real_fields = ['Zf','W_D_Dd','W_M','Zd','sp_psv','dev_plus','dev_v']
    s_row = 15
    for field_name in pro_fields:
        wsd.cell(row=s_row, column=9, value=field_name); s_row += 1

    wsd['L14'] = 'Self'
    self_fields = ['Ego[3r+(2)/R]','Fr+rF','SumV','FD','An+Xy','MOR','H:(H)+Hd+(Hd)']
    self_real_fields = ['ego','fr_rf','sum_V','fdn','an_xy','sp_mor','h_prop']
    s_row = 15
    for field_name in self_fields:
        wsd.cell(row=s_row, column=12, value=field_name); s_row += 1

    wsd['B4'] = structural_summary.R; wsd['D4'] = structural_summary.L
    wsd['B5'] = structural_summary.ErleBnistypus
    wsd['B6'] = structural_summary.eb
    wsd['D5'] = structural_summary.EA
    wsd['D6'] = structural_summary.es
    wsd['D7'] = structural_summary.adj_es
    wsd['F5'] = 'NA' if structural_summary.EBper == 0 else structural_summary.EBper
    wsd['F6'] = structural_summary.D_score
    wsd['F7'] = structural_summary.adj_D
    wsd['B8'] = structural_summary.sum_FM
    wsd['B9'] = structural_summary.sum_m
    wsd['D8'] = structural_summary.sum_Ca
    wsd['D9'] = structural_summary.sum_V
    wsd['F8'] = structural_summary.sum_T
    wsd['F9'] = structural_summary.sum_Y

    wsd['I4'] = structural_summary.f_c_prop; wsd['I4'].alignment = Alignment(horizontal='right')
    wsd['I5'] = structural_summary.pure_c
    wsd['I6'] = structural_summary.ca_c_prop; wsd['I6'].alignment = Alignment(horizontal='right')
    wsd['I7'] = structural_summary.afr; wsd['I7'].number_format = "0.##;-0.##;0"
    wsd['I8'] = structural_summary.S
    wsd['I9'] = structural_summary.blends_r; wsd['I9'].alignment = Alignment(horizontal='right')
    wsd['I10'] = structural_summary.sp_cp

    wsd['L4'] = structural_summary.sp_cop
    wsd['N4'] = structural_summary.sp_ag
    wsd['M5'] = structural_summary.GHR_PHR; wsd['M5'].alignment = Alignment(horizontal='right')
    wsd['M6'] = structural_summary.a_p; wsd['M6'].alignment = Alignment(horizontal='right')
    wsd['M7'] = structural_summary.Fd_l
    wsd['M8'] = structural_summary.sum_T
    wsd['M9'] = structural_summary.human_cont
    wsd['M10'] = structural_summary.H
    wsd['M11'] = structural_summary.sp_per
    wsd['M12'] = structural_summary.Isol

    wsd['B15'] = structural_summary.a_p; wsd['B15'].alignment = Alignment(horizontal='right')
    wsd['B16'] = structural_summary.Ma_Mp; wsd['B16'].alignment = Alignment(horizontal='right')
    wsd['B17'] = structural_summary.intel
    wsd['B18'] = structural_summary.sp_mor
    wsd['D15'] = structural_summary.sum6
    wsd['D16'] = structural_summary.Lvl_2
    wsd['D17'] = structural_summary.wsum6
    wsd['D18'] = structural_summary.mq_minus
    wsd['D19'] = structural_summary.mq_none

    row = 15
    for field_name in med_real_fields:
        val = getattr(structural_summary, field_name)
        wsd.cell(row=row, column=7, value=val)
        wsd.cell(row=row, column=7).number_format = "0" if isinstance(val, int) else "0.00"
        row += 1

    row = 15
    for field_name in pro_real_fields:
        wsd.cell(row=row, column=10, value=getattr(structural_summary, field_name))
        wsd.cell(row=row, column=10).alignment = Alignment(horizontal='right')
        row += 1

    row = 15
    for field_name in self_real_fields:
        wsd.cell(row=row, column=13, value=getattr(structural_summary, field_name))
        row += 1

    wsd['A23'] = f"PTI={structural_summary.sumPTI}"
    wsd['B23'] = "☑" if structural_summary.sumDEPI >= 5 else "☐"; wsd['B23'].alignment = Alignment(horizontal='right')
    wsd['C23'] = f"DEPI={structural_summary.sumDEPI}"
    wsd['D23'] = "☑" if structural_summary.sumCDI >= 4 else "☐"; wsd['D23'].alignment = Alignment(horizontal='right')
    wsd['E23'] = f"CDI={structural_summary.sumCDI}"
    wsd['F23'] = "☑" if structural_summary.sumSCON >= 8 else "☐"; wsd['F23'].alignment = Alignment(horizontal='right')
    wsd['G23'] = f"S-CON={structural_summary.sumSCON}"
    wsd['H23'] = "☑ HVI" if structural_summary.HVI_premise is True and structural_summary.sumHVI >= 4 else "☐ HVI"
    wsd['H23'].alignment = Alignment(horizontal='right')
    wsd['J23'] = "☑ OBS" if structural_summary.OBS_posi else "☐ OBS"; wsd['J23'].alignment = Alignment(horizontal='right')

    for column_cells in wsd.columns:
        length = max(len(str(cell.value)) * 1.1 for cell in column_cells)
        wsd.column_dimensions[column_cells[0].column_letter].width = length

    wsp['A1'] = "PTI"
    wsp['A2'] = "XA%<.70 AND WDA%<.75"
    wsp['A3'] = "X-%>0.29"
    wsp['A4'] = "LVL2>2 AND FAB2>0"
    wsp['A5'] = "R<17 AND Wsum6>12 OR R>16 AND Wsum6>17*"
    wsp['A6'] = "M- > 1 OR X-% > 0.40"
    wsp['A7'] = "TOTAL"
    row = 2
    for i in range(0, 5):
        value = "✔" if structural_summary.PTI[i] == "o" else ''
        wsp.cell(row=row, column=2, value=value); row += 1
    wsp['B7'] = structural_summary.sumPTI

    wsp['A9'] = "DEPI"
    wsp['A10'] = "SumV>0 OR FD>2"
    wsp['A11'] = "Col-shd blends>0 OR S>2"
    wsp['A12'] = "ego sup AND Fr+rF=0 OR ego inf"
    wsp['A13'] = "Afr<0.46 OR Blends<4"
    wsp['A14'] = "SumShd>FM+m OR SumC'>2"
    wsp['A15'] = "MOR>2 OR INTELL>3"
    wsp['A16'] = "COP<2 OR ISOL>0.24"
    wsp['A17'] = "TOTAL"
    wsp['A18'] = "POSITIVE?"
    row = 10
    for i in range(0, 7):
        value = "✔" if structural_summary.DEPI[i] == "o" else ''
        wsp.cell(row=row, column=2, value=value); row += 1
    wsp['B17'] = structural_summary.sumDEPI
    wsp['B18'] = structural_summary.sumDEPI >= 5

    wsp['A20'] = "CDI"
    wsp['A21'] = "EA<6 OR Daj<0"
    wsp['A22'] = "COP<2 AND AG<2"
    wsp['A23'] = "WSumC<2.5 OR Afr<0.46"
    wsp['A24'] = "p > a+1 OR pure H<2"
    wsp['A25'] = "SumT>1 OR ISOL>0.24 OR Fd>0"
    wsp['A26'] = "TOTAL"
    wsp['A27'] = "POSITIVE?"
    row = 21
    for i in range(0, 5):
        value = "✔" if structural_summary.CDI[i] == "o" else ''
        wsp.cell(row=row, column=2, value=value); row += 1
    wsp['B26'] = structural_summary.sumCDI
    wsp['B27'] = structural_summary.sumCDI >= 4

    wsp['D1'] = "S-CON"
    wsp['D2'] = "SumV+FD>2"
    wsp['D3'] = "col-shd blends>0"
    wsp['D4'] = "ego <0.31 ou >0.44"
    wsp['D5'] = "mor>3"
    wsp['D6'] = "Zd>3.5 ou <-3.5"
    wsp['D7'] = "es>EA"
    wsp['D8'] = "CF+C>FC"
    wsp['D9'] = "X+%<0.70"
    wsp['D10'] = "S>3"
    wsp['D11'] = "P<3 OU P>8"
    wsp['D12'] = "PURE H<2"
    wsp['D13'] = "R<17"
    wsp['D14'] = 'TOTAL'
    wsp['D15'] = 'POSITIVE?'
    row = 2
    for i in range(0, 12):
        value = "✔" if structural_summary.SCON[i] == "o" else ''
        wsp.cell(row=row, column=5, value=value); row += 1
    wsp['E14'] = structural_summary.sumSCON
    wsp['E15'] = structural_summary.sumSCON >= 8

    wsp['D17'] = 'HVI'
    wsp['D18'] = 'SumT = 0'
    wsp['D19'] = "Zf>12"
    wsp['D20'] = "Zd>3.5"
    wsp['D21'] = "S>3"
    wsp['D22'] = "H+(H)+Hd+(Hd)>6"
    wsp['D23'] = "(H)+(A)+(Hd)+(Ad)>3"
    wsp['D24'] = "H+A : 4:1"
    wsp['D25'] = "Cg>3"
    wsp['D26'] = 'TOTAL'
    wsp['D27'] = 'POSITIVE?'
    wsp['E18'] = structural_summary.HVI_premise
    row = 19
    for i in range(0, 7):
        value = "✔" if structural_summary.HVI[i] == "o" else ''
        wsp.cell(row=row, column=5, value=value); row += 1
    wsp['E26'] = structural_summary.sumHVI
    wsp['E27'] = (structural_summary.sumHVI >= 4) and bool(structural_summary.HVI_premise)

    wsp['A29'] = 'OBS'
    wsp['A30'] = 1; wsp['A31'] = 2; wsp['A32'] = 3; wsp['A33'] = 4; wsp['A34'] = 5
    wsp['B30'] = "Dd>3"; wsp['B31'] = "Zf>12"; wsp['B32'] = "Zd>3.0"; wsp['B33'] = "P>7"; wsp['B34'] = "FQ+>1"
    wsp['D30'] = "1-5 are true"
    wsp['D31'] = "FQ+>3 AND 2 items 1-4"
    wsp['D32'] = "X+%>0,89 et 3 items"
    wsp['D33'] = "FQ+>3 et X+%>0,89"
    wsp['D34'] = 'POSITIVE?'
    row = 30
    for i in range(0, 5):
        value = "✔" if structural_summary.OBS[i] == "o" else ''
        wsp.cell(row=row, column=3, value=value); row += 1
    row = 30
    for i in range(5, 9):
        value = "✔" if structural_summary.OBS[i] == "o" else ''
        wsp.cell(row=row, column=5, value=value); row += 1
    wsp['E34'] = structural_summary.OBS_posi

    for column_cells in wsp.columns:
        length = max(len(str(cell.value)) * 1.1 for cell in column_cells)
        wsp.column_dimensions[column_cells[0].column_letter].width = length

    ws2 = wb.create_sheet(title="raw data")

    def _card_num(rc):
        try:
            return int(normalize_card_to_num(rc.card))
        except Exception:
            return 999
    def _n(rc):
        return rc.response_num or 0

    response_codes_sorted = sorted(response_codes, key=lambda rc: (_card_num(rc), _n(rc)))

    response_code_data = []
    for rc in response_codes_sorted:
        response_code_data.append({
            'Card': to_roman(rc.card),
            'N': rc.response_num,
            'time': rc.time,
            'response': rc.response,
            'V': rc.rotation,
            'inquiry': rc.inquiry,
            'Location': rc.location,
            'loc_num': rc.loc_num,
            'Dev Qual': rc.dev_qual,
            'determinants': rc.determinants,
            'Form Quality': rc.form_qual,
            '(2)': rc.pair,
            'Content': rc.content,
            'P': rc.popular,
            'Z': rc.Z,
            'special': rc.special,
            'comment': rc.comment
        })
    df_raw_for_sheet = pd.DataFrame(response_code_data)

    for col_idx, col_name in enumerate(df_raw_for_sheet.columns, start=1):
        cell = ws2.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)

    for row_vals in df_raw_for_sheet.values.tolist():
        ws2.append(row_vals)

    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:{get_column_letter(ws2.max_column)}1"
    for col in range(1, ws2.max_column + 1):
        max_len = 0
        for row in range(1, ws2.max_row + 1):
            v = ws2.cell(row=row, column=col).value
            max_len = max(max_len, len(str(v)) if v is not None else 0)
        ws2.column_dimensions[get_column_letter(col)].width = max(8, min(60, int(max_len * 1.1)))

    try:
        df_raw = pd.DataFrame([{
            'ID': client.id,
            '카드': normalize_card_to_num(rc.card),
            'N': rc.response_num,
            '반응': rc.response,
            '질문': rc.inquiry,
            '결정인': rc.determinants,
            '(2)': rc.pair,
            '내용인': rc.content,
            '특수점수': rc.special,
        } for rc in response_codes])

        columns_to_keep = ['ID','카드','N','반응','질문','결정인','(2)','내용인','특수점수']
        df_selected = df_raw[columns_to_keep].copy()
        df_processed = df_selected.apply(_apply_pair_into_determinants, axis=1).drop(columns=['(2)'])

        df_processed['RESPONSE_토큰'] = df_processed['반응'].apply(lambda x: list(set(tokenize_with_pos(x))))
        df_processed['INQUIRY_토큰']  = df_processed['질문'].apply(lambda x: list(set(tokenize_with_pos(x))))

        rsp_score = _read_json_df(
            RESOURCE_DIR / RESOURCE_FILENAMES['response_score'],
            required_cols=['카드','토큰','품사','점수']
        )
        inq_score = _read_json_df(
            RESOURCE_DIR / RESOURCE_FILENAMES['inquiry_score'],
            required_cols=['카드','토큰','품사','점수']
        )
        sym_score = _read_json_df(
            RESOURCE_DIR / RESOURCE_FILENAMES['symbol_score'],
            required_cols=['카드','채점영역','기호','점수']
        )
        sc_stats = _read_json_df(
            RESOURCE_DIR / RESOURCE_FILENAMES['score_stats'],
            required_cols=['카드','채점영역','mean','std']
        )
        idx_stats = _read_json_df(
            RESOURCE_DIR / RESOURCE_FILENAMES['index_stats'],
            required_cols=['카드','mean','std']
        )

        for _df in (rsp_score, inq_score, sym_score, sc_stats, idx_stats, df_processed):
            _df['카드'] = _df['카드'].astype(str)

        df_scored = df_processed.copy()
        df_scored = _calculate_token_score(df_scored, rsp_score, 'RESPONSE_토큰', 'RESPONSE_점수')
        df_scored = _calculate_token_score(df_scored, inq_score, 'INQUIRY_토큰',  'INQUIRY_점수')
        df_scored = _apply_symbol_score(df_scored, sym_score)

        mean_df = sc_stats.pivot(index='카드', columns='채점영역', values='mean')
        std_df  = sc_stats.pivot(index='카드', columns='채점영역', values='std')

        candidate_areas = ['RESPONSE_점수','INQUIRY_점수','결정인_점수','내용인_점수','특수점수_점수']
        areas = [a for a in candidate_areas if a in df_scored.columns and a in mean_df.columns]

        def _area_z(row, area):
            card = str(row['카드'])
            try:
                m = float(mean_df.loc[card, area])
                s = float(std_df.loc[card, area])
                if s == 0:
                    return np.nan
                return (row[area] - m) / s
            except Exception:
                return np.nan

        for area in areas:
            df_scored[f"{area}_z"] = df_scored.apply(lambda r, a=area: _area_z(r, a), axis=1)

        for col in ['결정인_점수_z','내용인_점수_z','특수점수_점수_z','INQUIRY_점수_z','RESPONSE_점수_z']:
            if col not in df_scored.columns:
                df_scored[col] = 0.0
            else:
                df_scored[col] = df_scored[col].fillna(0.0)

        z_cols = ['결정인_점수_z','내용인_점수_z','특수점수_점수_z','INQUIRY_점수_z','RESPONSE_점수_z']
        df_scored['투사지수_final'] = df_scored[z_cols].sum(axis=1)

        df_scored = df_scored.merge(idx_stats, on='카드', how='left')
        df_scored['std'] = df_scored['std'].replace(0, np.nan)
        df_scored['std'] = df_scored['std'].fillna(df_scored['std'].mean() if not pd.isna(df_scored['std'].mean()) else 1.0)
        df_scored['mean'] = df_scored['mean'].fillna(0.0)
        df_scored['투사지수_T'] = 50 + 10 * (df_scored['투사지수_final'] - df_scored['mean']) / df_scored['std']
        df_scored = df_scored.drop(columns=['mean','std'])

        df_card_avg = (df_scored
                       .groupby(['ID','카드'], as_index=False)['투사지수_T']
                       .mean()
                       .rename(columns={'투사지수_T':'투사지수_카드별평균'}))
        df_card_avg['카드'] = df_card_avg['카드'].astype(int)
        df_card_avg = df_card_avg.sort_values(['ID','카드']).reset_index(drop=True)

        df_id_avg = (df_card_avg
                     .groupby('ID', as_index=False)['투사지수_카드별평균']
                     .mean()
                     .rename(columns={'투사지수_카드별평균':'투사지수_전체평균'}))

        ws_proc = wb.create_sheet(title='processed')
        cols_p = ['ID','카드','N','반응','질문','결정인','내용인','특수점수','RESPONSE_토큰','INQUIRY_토큰']
        _df_p = df_scored.copy()

        def _serialize_tokens(toks):
            if not toks:
                return ''
            out = []
            for t in toks:
                try:
                    w, p = t
                except Exception:
                    continue
                w = str(w).replace(';', '／')
                p = str(p).replace(';', '／')
                out.append(f"{w}/{p}")
            return ';'.join(out)

        _df_p['RESPONSE_토큰'] = _df_p['RESPONSE_토큰'].apply(_serialize_tokens)
        _df_p['INQUIRY_토큰']  = _df_p['INQUIRY_토큰'].apply(_serialize_tokens)
        _df_p = _df_p[cols_p].copy()
        _df_p['카드'] = pd.to_numeric(_df_p['카드'], errors='coerce').fillna(0).astype(int)
        _df_p['N'] = pd.to_numeric(_df_p['N'], errors='coerce')
        _df_p = _df_p.sort_values(['카드','N'], kind='mergesort')
        for i, c in enumerate(_df_p.columns, start=1):
            cell = ws_proc.cell(row=1, column=i, value=c); cell.font = Font(bold=True)
        for rowv in _df_p.values.tolist():
            ws_proc.append(rowv)

        ws_score = wb.create_sheet(title='score')
        cols_s = [
            'ID','카드','N',
            'RESPONSE_점수','INQUIRY_점수','결정인_점수','내용인_점수','특수점수_점수',
            'RESPONSE_점수_z','INQUIRY_점수_z','결정인_점수_z','내용인_점수_z','특수점수_점수_z',
            '투사지수_final','투사지수_T'
        ]
        for need in ['RESPONSE_점수','INQUIRY_점수','결정인_점수','내용인_점수','특수점수_점수',
                     'RESPONSE_점수_z','INQUIRY_점수_z','결정인_점수_z','내용인_점수_z','특수점수_점수_z']:
            if need not in df_scored.columns:
                df_scored[need] = 0.0
        _df_s = df_scored[cols_s].copy()
        _df_s['카드'] = pd.to_numeric(_df_s['카드'], errors='coerce').fillna(0).astype(int)
        _df_s['N'] = pd.to_numeric(_df_s['N'], errors='coerce')
        _df_s = _df_s.sort_values(['카드','N'], kind='mergesort')
        for i, c in enumerate(_df_s.columns, start=1):
            cell = ws_score.cell(row=1, column=i, value=c); cell.font = Font(bold=True)
        for rowv in _df_s.values.tolist():
            ws_score.append(rowv)

        ws_card = wb.create_sheet(title='card_avg')
        for i, c in enumerate(df_card_avg.columns, start=1):
            cell = ws_card.cell(row=1, column=i, value=c); cell.font = Font(bold=True)
        for rowv in df_card_avg.values.tolist():
            ws_card.append(rowv)

        ws_over = wb.create_sheet(title='overall')
        for i, c in enumerate(df_id_avg.columns, start=1):
            cell = ws_over.cell(row=1, column=i, value=c); cell.font = Font(bold=True)
        for rowv in df_id_avg.values.tolist():
            ws_over.append(rowv)

        ws_raw = wb.create_sheet(title='rawdata')

        def _card_num(rc):
            try:
                return int(normalize_card_to_num(rc.card))
            except Exception:
                return 999
        def _n(rc): return rc.response_num or 0

        response_codes_sorted = sorted(response_codes, key=lambda rc: (_card_num(rc), _n(rc)))
        response_code_data = []
        for rc in response_codes_sorted:
            response_code_data.append({
                'Card': to_roman(rc.card),
                'N': rc.response_num,
                'time': rc.time,
                'response': rc.response,
                'V': rc.rotation,
                'inquiry': rc.inquiry,
                'Location': rc.location,
                'loc_num': rc.loc_num,
                'Dev Qual': rc.dev_qual,
                'determinants': rc.determinants,
                'Form Quality': rc.form_qual,
                '(2)': rc.pair,
                'Content': rc.content,
                'P': rc.popular,
                'Z': rc.Z,
                'special': rc.special,
                'comment': rc.comment
            })
        df_raw_for_sheet = pd.DataFrame(response_code_data)

        for col_idx, col_name in enumerate(df_raw_for_sheet.columns, start=1):
            cell = ws_raw.cell(row=1, column=col_idx, value=col_name)
            cell.font = Font(bold=True)
        for row_vals in df_raw_for_sheet.values.tolist():
            ws_raw.append(row_vals)

        ws_raw.freeze_panes = "A2"
        ws_raw.auto_filter.ref = f"A1:{get_column_letter(ws_raw.max_column)}1"
        for col in range(1, ws_raw.max_column + 1):
            max_len = 0
            for row in range(1, ws_raw.max_row + 1):
                v = ws_raw.cell(row=row, column=col).value
                max_len = max(max_len, len(str(v)) if v is not None else 0)
            ws_raw.column_dimensions[get_column_letter(col)].width = max(8, min(60, int(max_len * 1.1)))

    except Exception as e:
        logging.exception("고급 산출 실패")
        ws_err = wb.create_sheet(title='advanced_error')
        ws_err['A1'] = "고급 산출 과정에서 오류가 발생했습니다."
        ws_err['A2'] = f"{type(e).__name__}: {str(e)}"
        ws_err.column_dimensions['A'].width = 120

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    response = HttpResponse(
        output.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=structural_summary_advanced.xlsx'
    return response

@group_min_required('advanced')
def advanced_edit_responses(request, client_id):
    client = get_object_or_404(Client, id=client_id)
    if client.tester != request.user:
        return HttpResponse("액세스 거부: 작성 권한이 없습니다.", status=403)

    qs = ResponseCode.objects.filter(client=client).order_by('id')
    current = qs.count()
    extra = max(0, min(DEFAULT_EXTRA, TOTAL_CAP - current))
    FormSet = modelformset_factory(ResponseCode, form=ResponseCodeForm, extra=extra, max_num=TOTAL_CAP)

    if request.method == 'POST':
        formset = FormSet(request.POST, queryset=qs)
        if formset.is_valid():
            saved = 0
            for form in formset:
                if form.cleaned_data.get('card') and form.cleaned_data.get('response'):
                    inst = form.save(commit=False)
                    inst.client = client
                    inst.card = to_roman(inst.card)
                    inst.save()
                    saved += 1

            structural_summary, _ = StructuralSummary.objects.get_or_create(client=client)
            try:
                structural_summary.calculate_values()
            except Exception:
                structural_summary.save()

            messages.success(request, f"{saved}건 저장되었습니다.")
            return redirect('scoring:client_detail', client_id=client.id)
        else:
            errors = []
            for i, f in enumerate(formset.forms, start=1):
                if f.errors:
                    for k, errs in f.errors.items():
                        errors.append(f"{i}행 {k}: {', '.join(errs)}")
            if errors:
                messages.error(request, "입력 오류: " + " / ".join(errors))
    else:
        formset = FormSet(queryset=qs)

    return render(request, 'update_response_codes.html', {
        'client': client,
        'formset': formset,
    })
