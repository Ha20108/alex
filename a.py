
import tkinter as tk
from tkinter import messagebox
import sqlite3

# إنشاء قاعدة بيانات SQLite
def connect_db():
    return sqlite3.connect("shipments.db")

# إضافة شحنة إلى قاعدة البيانات
def add_shipment():
    company_name = company_name_entry.get()
    vessel_name = vessel_name_entry.get()
    bill_of_lading = bill_of_lading_entry.get()
    documents_received_date = documents_received_date_entry.get()
    expected_arrival_date = expected_arrival_date_entry.get()
    weight = weight_entry.get()
    packages_count = packages_count_entry.get()
    invoice_number = invoice_number_entry.get()

    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO shipments (company_name, vessel_name, bill_of_lading, 
                               documents_received_date, expected_arrival_date, 
                               weight, packages_count, invoice_number) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)""", 
        (company_name, vessel_name, bill_of_lading, documents_received_date, 
         expected_arrival_date, weight, packages_count, invoice_number))
    conn.commit()
    conn.close()
    messagebox.showinfo("Success", "تم إضافة الشحنة بنجاح")

# واجهة Tkinter
root = tk.Tk()
root.title("إضافة شحنة")

tk.Label(root, text="اسم الشركة:").grid(row=0, column=0)
company_name_entry = tk.Entry(root)
company_name_entry.grid(row=0, column=1)

tk.Label(root, text="اسم المركب:").grid(row=1, column=0)
vessel_name_entry = tk.Entry(root)
vessel_name_entry.grid(row=1, column=1)

tk.Label(root, text="رقم بوليصة الشحن:").grid(row=2, column=0)
bill_of_lading_entry = tk.Entry(root)
bill_of_lading_entry.grid(row=2, column=1)

tk.Label(root, text="تاريخ استلام الأوراق:").grid(row=3, column=0)
documents_received_date_entry = tk.Entry(root)
documents_received_date_entry.grid(row=3, column=1)

tk.Label(root, text="تاريخ الوصول المتوقع:").grid(row=4, column=0)
expected_arrival_date_entry = tk.Entry(root)
expected_arrival_date_entry.grid(row=4, column=1)

tk.Label(root, text="الوزن (كجم):").grid(row=5, column=0)
weight_entry = tk.Entry(root)
weight_entry.grid(row=5, column=1)

tk.Label(root, text="عدد الطرود:").grid(row=6, column=0)
packages_count_entry = tk.Entry(root)
packages_count_entry.grid(row=6, column=1)

tk.Label(root, text="رقم الفاتورة:").grid(row=7, column=0)
invoice_number_entry = tk.Entry(root)
invoice_number_entry.grid(row=7, column=1)

tk.Button(root, text="إضافة الشحنة", command=add_shipment).grid(row=8, column=1)

root.mainloop()
