from functools import wraps
import re

from django.contrib.auth.decorators import login_required
from django.http import HttpResponseForbidden

GROUP_LEVEL = {'beginner': 1, 'intermediate': 2, 'advanced': 3}
GROUP_LABEL = {'beginner': '초급', 'intermediate': '중급', 'advanced': '고급'}


def group_min_required(min_group_name: str):
    def decorator(view_func):
        @wraps(view_func)
        def _wrapped_view(request, *args, **kwargs):
            if not request.user.is_authenticated:
                return login_required(view_func)(request, *args, **kwargs)
            user_level = GROUP_LEVEL.get(getattr(request.user, 'group', None), 0)
            required_level = GROUP_LEVEL[min_group_name]
            if user_level >= required_level:
                return view_func(request, *args, **kwargs)
            return HttpResponseForbidden(f"{GROUP_LABEL[min_group_name]} 이상 이수자만 접속 가능한 페이지입니다.")
        return _wrapped_view
    return decorator


UNICODE_ROMAN = {
    'Ⅰ': 'I', 'Ⅱ': 'II', 'Ⅲ': 'III', 'Ⅳ': 'IV', 'Ⅴ': 'V',
    'Ⅵ': 'VI', 'Ⅶ': 'VII', 'Ⅷ': 'VIII', 'Ⅸ': 'IX', 'Ⅹ': 'X'
}

ROMAN_TO_NUM = {'I': '1', 'II': '2', 'III': '3', 'IV': '4', 'V': '5',
                'VI': '6', 'VII': '7', 'VIII': '8', 'IX': '9', 'X': '10'}
NUM_TO_ROMAN = {v: k for k, v in ROMAN_TO_NUM.items()}


def normalize_card_to_num(val: str) -> str:
    if val is None:
        return ''
    s = str(val).strip()

    for u, r in UNICODE_ROMAN.items():
        s = s.replace(u, r)

    su = s.upper()
    if su in ROMAN_TO_NUM:
        return ROMAN_TO_NUM[su]

    su = re.sub(r'[^IVX0-9]', '', su)

    if su in ROMAN_TO_NUM:
        return ROMAN_TO_NUM[su]

    if re.fullmatch(r'([1-9]|10)', su):
        return su

    return s


def to_roman(val: str) -> str:
    n = normalize_card_to_num(val)
    return NUM_TO_ROMAN.get(str(n), str(val).strip())
