from django.urls import path
from . import views

urlpatterns = [
    path('',views.home, name='home'),
    path('add_company/', views.add_company, name='add_company'),
    path('company_list/', views.company_list, name='company_list'),
    path('add_shipment/', views.add_shipment, name='add_shipment'),
    path('shipment_list/', views.shipment_list, name='shipment_list'),
    path('export/', views.export_shipments_to_excel, name='export_shipments'),
    path('upload_excel/', views.upload_excel, name='upload_excel'),
    path('save_shipment_changes/', views.save_shipment_changes, name='save_shipment_changes'),
    #path('download-report/', views.upload_and_generate_report, name='download_report'),
    path('upload/', views.upload_and_generate_report, name='upload_report'),


    path('transaction_list', views.transaction_list, name='transaction_list'),
    path('export1/', views.export_transactions_excel, name='export_transactions_excel'),

]
