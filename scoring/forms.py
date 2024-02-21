from django import forms
from .models import Client, ResponseCode
from django.core.exceptions import ValidationError
from suit.widgets import AutosizedTextarea
import re


def validate_card(value):
    if value not in ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', '1', '2', '3', '4', '5', '6', '7',
                     '8', '9', '10']:
        raise ValidationError(
            "기호 오류"
        )
    return value


def validate_loc(value):
    if value not in ['W', 'WS', 'D', 'DS', 'Dd', 'DdS']:
        raise ValidationError(
            "기호 오류"
        )
    return value


def validate_dev_qual(value):
    if value not in ['+', 'o', 'v/+', 'v']:
        raise ValidationError(
            "기호 오류"
        )
    return value


def validate_determinants(value):
    elements = [element.strip() for element in re.split(r'[,.]+', value.replace(' ', ''))]
    elements_lower = [item.lower() if item not in ['ma', 'mp', 'Ma', 'Mp', 'Ma-p']
                      else item for item in elements]

    for element in elements_lower:
        if element not in ['Ma', 'Mp', 'Ma-p', 'fma', 'fmp', 'fma-p', 'ma', 'mp', 'ma-p', 'fc', 'cf', 'c', 'cn', "fc'",
                           "c'f", "c'", 'ft', 'tf', 't', 'fv', 'vf', 'v', 'fy', 'yf', 'y', 'fr', 'rf', 'fd', 'f']:
            raise ValidationError(
                "기호 오류"
            )

    return value


def validate_special(value):
    elements = [element.strip() for element in re.split(r'[,.]+', value.replace(' ', ''))]

    for element in elements:
        if element not in ['DV', 'DV2', 'DR', 'DR2', 'INC', 'INC2', 'FAB', 'FAB2', 'CON', 'ALOG', 'PSV',
                           'AB', 'AG', 'COP', 'MOR', 'PER', 'CP', 'GHR', 'PHR']:
            raise ValidationError(
                "기호 오류"
            )

    return value


def validate_contents(value):
    elements = [element.strip() for element in re.split(r'[,.]+', value.replace(' ', ''))]
    elements_lower = [item.lower() for item in elements]

    for element in elements_lower:
        if element not in ['h', '(h)', 'hd', '(hd)', 'hx', 'a', '(a)', '(ad)', 'ad', 'an', 'art', 'ay', 'bl', "bt",
                           "cg", "cl", 'ex', 'fi', 'fd', 'ge', 'hh', 'ls', 'na', 'sc', 'sx', 'xy']:
            raise ValidationError(
                "기호 오류"
            )

    return value


def validate_fq(value):
    if value not in ['+', 'o', 'u', '-', 'no']:
        raise ValidationError(
            "기호 오류"
        )
    return value


def validate_P(value):
    if value not in ['P', '']:
        raise ValidationError(
            "P만 가능"
        )
    return value


def validate_Z(value):
    if value not in ['ZA', 'ZW', 'ZD', 'ZS']:
        raise ValidationError(
            "기호 오류"
        )
    return value


def validate_pair(value):
    if value not in ['2']:
        raise ValidationError(
            "2 입력"
        )
    return value


class ClientForm(forms.ModelForm):
    class Meta:
        model = Client
        fields = ['name', 'gender', 'birthdate', 'testDate', 'notes']

    name = forms.CharField(widget=forms.TextInput(attrs={'class': 'form-control'}))
    gender = forms.ChoiceField(choices=[('M', '남성'), ('F', '여성'), ('O', '기타')],
                               widget=forms.Select(attrs={'class': 'form-control'}))
    birthdate = forms.DateField(widget=forms.DateInput(attrs={'class': 'form-control'}))
    testDate = forms.DateField(widget=forms.DateInput(attrs={'class': 'form-control'}))
    notes = forms.CharField(widget=forms.Textarea(attrs={'class': 'form-control'}), required=False)

    def clean(self):
        cleaned_data = super().clean()
        birthdate = cleaned_data.get('birthdate')
        testDate = cleaned_data.get('testDate')

        if birthdate and testDate:
            if testDate < birthdate:
                raise forms.ValidationError("검사일은 생년월일 이후여야 합니다.")

        return cleaned_data


class ResponseCodeForm(forms.ModelForm):
    class Meta:
        model = ResponseCode
        fields = ['card', 'response_num', 'time', 'response', 'inquiry', 'rotation', 'location', 'dev_qual', 'loc_num',
                  'determinants', 'form_qual', 'pair', 'content', 'popular', 'Z', 'special', 'comment']

    card = forms.CharField(required=True, validators=[validate_card], error_messages={'required': '필수'})
    response_num = forms.IntegerField(required=True, error_messages={'required': '필수'})
    time = forms.CharField(required=False)
    response = forms.CharField(required=True, error_messages={'required': '필수'},
                               widget=AutosizedTextarea(attrs={'rows': 1, 'cols': 2}))
    inquiry = forms.CharField(required=True, error_messages={'required': '필수'},
                              widget=AutosizedTextarea(attrs={'rows': 1, 'cols': 3}))
    rotation = forms.CharField(required=False)
    location = forms.CharField(required=True, error_messages={'required': '필수'}, validators=[validate_loc])
    dev_qual = forms.CharField(required=True, error_messages={'required': '필수'}, validators=[validate_dev_qual])
    loc_num = forms.IntegerField(required=False)
    determinants = forms.CharField(required=True, error_messages={'required': '필수'}, validators=[validate_determinants])
    pair = forms.CharField(required=False, validators=[validate_pair])
    form_qual = forms.CharField(required=True, error_messages={'required': '필수'}, validators=[validate_fq])
    content = forms.CharField(required=True, error_messages={'required': '필수'})
    popular = forms.CharField(required=False, validators=[validate_P])
    Z = forms.CharField(required=False, validators=[validate_Z])
    special = forms.CharField(required=False, validators=[validate_special])
    comment = forms.CharField(required=False, widget=AutosizedTextarea(attrs={'rows': 1, 'cols': 3}))

    def clean(self):
        cleaned_data = super().clean()
        cleaned_data = {k: v for k, v in cleaned_data.items() if v}
        loc_value = cleaned_data.get('location', '')
        Z_value = cleaned_data.get('Z', '')
        dq_value = cleaned_data.get('dev_qual', '')
        card_value = cleaned_data.get('card', '')

        # Z 점수 체크
        if 'W' in loc_value and not Z_value and dq_value != 'v':
            self.add_error('Z', ValidationError('Z 점수 필요', code='invalid'))
        elif '+' in dq_value and not Z_value:
            self.add_error('Z', ValidationError('Z 점수 필요', code='invalid'))
        elif card_value in ['1', '4', '5', 'I', 'II',
                            'III'] and 'W' in loc_value and '+' in dq_value and Z_value == 'ZW':
            self.add_error('Z', ValidationError('더 높은 Z', code='invalid'))

        return cleaned_data
