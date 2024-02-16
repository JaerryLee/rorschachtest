from django.shortcuts import redirect, render

from .forms import SignupForm


def signup(request):
    if request.method == 'POST':
        form = SignupForm(request.POST)
        if form.is_valid():
            user = form.save(commit=False)  # commit=False를 사용하여 임시로 저장
            user.is_active = False
            user.save()
            return redirect("/accounts/wait")
        else:
            for field in form:
                print(field.errors)
    else:
        form = SignupForm(initial={
            'username': request.GET.get('username', ''),
            'email': request.GET.get('email', ''),
            'first_name': request.GET.get('first_name', ''),
            'last_name': request.GET.get('last_name', ''),
            'phone': request.GET.get('phone', ''),
            'password1': request.GET.get('password1', ''),
            'password2': request.GET.get('password2', ''),
            'consent': request.GET.get('consent', '')
        })
    return render(request, 'accounts/signup_form.html', {
        'form': form,
    })


def wait(request):
    return render(request, 'accounts/wait.html')
