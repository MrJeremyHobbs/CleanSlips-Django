from django.urls import path
#from django.conf.urls import url
from cleanslips import views

urlpatterns = [
    #url(r'^', views.upload, name='uplink'),
    path(r'<campus>', views.upload, name='uplink'),
]