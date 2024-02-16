from .models import SearchReference, CardImages, PopularResponse
from django_filters import FilterSet, CharFilter

class SearchReferenceFilter(FilterSet):
    LOC = CharFilter(field_name='LOC', lookup_expr='icontains', label='LOC')
    Cont = CharFilter(field_name='Cont', lookup_expr='icontains', label='Cont')
    Determinants = CharFilter(field_name='Determinants', lookup_expr='icontains', label='Determinants')

    class Meta:
        model = SearchReference
        fields = ['Card', 'LOC', 'Cont', 'FQ', 'Determinants', 'Item', 'V',]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.form.fields['Card'].required = True
        self.form.fields['LOC'].required = True
        self.form.fields['Cont'].required = True

class CardImagesFilter(FilterSet):
    card_number = CharFilter(
        field_name='card_number',
        lookup_expr='exact',
        method='filter_with_card_number'
    )

    class Meta:
        model = CardImages
        fields = []

    def filter_with_card_number(self, queryset, name, value):
        return queryset.filter(card_number=value)

class PResponseFilter(FilterSet):
    card_number = CharFilter(
        field_name='card_number',
        lookup_expr='exact',
        method='filter_with_card_number'
    )

    class Meta:
        model = PopularResponse
        fields = []

    def filter_with_card_number(self, queryset, name, value):
        return queryset.filter(card_number=value)