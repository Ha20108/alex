from django import forms
from .models import Shipment, Company
from .models import Transaction

class TransactionForm(forms.ModelForm):
    class Meta:
        model = Transaction
        fields = ['date', 'description', 'type', 'currency', 'amount']
        widgets = {
            'date': forms.DateInput(attrs={'class': 'form-control', 'type': 'date'}),
            'description': forms.Textarea(attrs={'class': 'form-control', 'rows': 1, 'placeholder': 'اكتب البيان هنا'}),
            'type': forms.Select(attrs={'class': 'form-select'}),
            'currency': forms.Select(attrs={'class': 'form-select'}),
            'amount': forms.NumberInput(attrs={'class': 'form-control', 'placeholder': '0.00'}),
        }

class ExcelUploadForm(forms.Form):
    excel_file = forms.FileField()

class ShipmentForm(forms.ModelForm):
    class Meta:
        model = Shipment
        fields = [  
            'company',
            'agency',
            'ducmentsno',
            'Acidno',
            'documents_received_date',
            'expected_arrival_date',
            'storge_data',
            'Delivery_data',
            'NoCE_data',
            'End_customs_data',
            'exchange_data',
            'vessel_name',
            'bill_of_lading',
            'weight',
            'packages_count',
            'invoice_number',
            'comment'
        ]
        labels = {
            'company': 'اسم الشركة',
            'agency': 'اسم التوكيل',
            'ducmentsno': 'رقم الاقرار',
            'Acidno': 'ACID',
            'documents_received_date': 'تاريخ استلام المستندات',
            'expected_arrival_date': 'تاريخ الوصول المتوقع',
            'storge_data': 'تاريخ التخزين',
            'Delivery_data': 'اذن التسليم',
            'NoCE_data': 'فتح الشهاده',
            'End_customs_data': 'تاريخ الانتهاء',
            'exchange_data': 'تاريخ الصرف',
            'vessel_name': 'اسم المركب',
            'bill_of_lading': 'رقم البوليصه',
            'weight': 'الوزن',
            'packages_count': 'عدد الطرود',
            'invoice_number': 'رقم الفاتورة',
            'comment': 'ملاحظات'
        }

class CompanyForm(forms.ModelForm):
    class Meta:
        model = Company
        fields = ['name', ]


class UploadAttendanceForm(forms.Form):
    file = forms.FileField(label="ارفع ملف الحضور (Excel)")


