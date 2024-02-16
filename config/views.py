from django.shortcuts import render


def main(request):
    return render(request, "main.html")


def greeting(request):
    return render(request, "greeting.html")


def about(request):
    return render(request, "about.html")


def privacy(request):
    return render(request, "privacy_policy.html")


def register(request):
    return render(request, "register.html")


def plan(request):
    return render(request, "plan.html")

def service(request):
    return render(request, "service.html")