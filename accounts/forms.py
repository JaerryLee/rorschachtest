from django.contrib.auth.forms import UserCreationForm
from .models import User


class SignupForm(UserCreationForm):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['email'].required = True
        self.fields['first_name'].required = True
        self.fields['last_name'].required = True
        self.fields['phone'].required = True
        self.fields['username'].help_text = None
        self.fields['password1'].help_text = "-최소 8자 이상, 문자 및 숫자 포함"
        self.fields['password2'].help_text = None
        self.fields['consent'].required = True
        for field in self.fields:
            self.fields[field].label = ''

    class Meta(UserCreationForm.Meta):
        model = User
        fields = ['username', 'email', 'first_name', 'last_name', 'phone', 'consent']
