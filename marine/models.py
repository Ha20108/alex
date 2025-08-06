from django.db import models

# نموذج الشركة
class Company(models.Model):
    name = models.CharField(max_length=255, unique=True)
    #tax_number = models.CharField(max_length=255 ,null=True , blank=True)

    def __str__(self):
        return self.name

CURRENCY_CHOICES = [
    ('EGP', 'جنيه'),
    ('USD', 'دولار'),
    ('EUR', 'يورو'),
]

TYPE_CHOICES = [
    ('in', 'وارد'),
    ('out', 'صادر'),
]

class Transaction(models.Model):
    date = models.DateField()
    description = models.TextField()
    type = models.CharField(max_length=3, choices=TYPE_CHOICES)
    currency = models.CharField(max_length=3, choices=CURRENCY_CHOICES)
    amount = models.DecimalField(max_digits=12, decimal_places=2)

    def __str__(self):
        return f"{self.date} - {self.description} - {self.amount} {self.currency}"

# نموذج الشحنة
class Shipment(models.Model):
    company = models.ForeignKey(Company, on_delete=models.CASCADE)
    agency = models.CharField(max_length=255 ,null=True , blank=True)
    ducmentsno = models.CharField(max_length=255 ,null=True , blank=True)
    Acidno = models.IntegerField(null=True , blank=True)
    NoCR = models.IntegerField(null=True , blank=True)
    documents_received_date = models.CharField(max_length=255 ,null=True , blank=True)
    expected_arrival_date = models.CharField(max_length=255 ,null=True , blank=True)
    storge_data = models.CharField(max_length=255 ,null=True , blank=True)
    Delivery_data = models.CharField(max_length=255 ,null=True , blank=True)
    NoCE_data = models.CharField(max_length=255 ,null=True , blank=True)
    End_customs_data = models.CharField(max_length=255 ,null=True , blank=True)
    exchange_data = models.CharField(max_length=255 ,null=True , blank=True)
    vessel_name = models.CharField(max_length=255 ,null=True , blank=True)
    bill_of_lading = models.CharField(max_length=255 ,null=True , blank=True)
    weight = models.CharField(max_length=255 ,null=True , blank=True)
    packages_count = models.CharField(max_length=255 ,null=True , blank=True)
    invoice_number = models.CharField(max_length=255 ,null=True , blank=True)
    comment=models.TextField(null=True , blank=True)

    def __str__(self):
        return f"Shipment {self.company}  {self.ducmentsno}"

