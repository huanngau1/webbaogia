from django.urls import path
from . import views
from .views import download_excel


urlpatterns = [
    path('', views.home, name="home"),
    path('download-excel/', download_excel, name='download_excel'),
]