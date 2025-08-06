from django.contrib import admin
from .models import *

# تسجيل الموديل في واجهة الـ Admin
admin.site.register(Company)
admin.site.register(Shipment)
admin.site.register(Transaction)

