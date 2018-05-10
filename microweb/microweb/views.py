from django.contrib.auth import views as auth_views
from django.shortcuts import render, HttpResponseRedirect

def index(request):
    a = request.user
    if a.is_authenticated:
        return HttpResponseRedirect('login')
    else:
        return render(request, "registration/login.html")
