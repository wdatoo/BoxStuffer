from django.urls import path
from .views import *


urlpatterns = [
    path("", HomeView.as_view(), name="home"),
     path("file-upload", FileView.as_view(), name="file_upload"),
]
