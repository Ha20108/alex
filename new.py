import pandas as pd
from datetime import datetime, timedelta

# تحميل ملف الحضور
file_path = '55.xlsx'  # ← ضع اسم الملف الذي حفظته هنا
df = pd.read_excel(file_path)

# تحويل عمود التاريخ إلى نوع تاريخ
df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)

# تحديد الشهر المطلوب (أبريل 2025)
start_date = datetime(2025, 4, 1)
end_date = datetime(2025, 4, 30)

# إنشاء قائمة بكل الأيام في أبريل 2025 باستثناء الجمعة
all_days = pd.date_range(start=start_date, end=end_date, freq='D')
working_days = [d for d in all_days if d.weekday() != 4]  # 4 = الجمعة

# استخراج أسماء الموظفين من العمود
employee_names = df['Name'].dropna().unique()

# إنشاء قائمة للنتائج
report_data = []

# إنشاء تقرير لكل موظف ولكل يوم عمل
for date in working_days:
    for name in employee_names:
        record = df[(df['Name'] == name) & (df['Date'] == date)]
        if not record.empty:
            clock_in = record.iloc[0].get('Clock In', '')
            clock_out = record.iloc[0].get('Clock Out', '')
            Late = record.iloc[0].get('Late', '')
            status = 'حاضر'
        else:
            clock_in = ''
            clock_out = ''
            Late = ''
            status = 'غائب'
        
        report_data.append({
            'التاريخ': date.strftime('%Y-%m-%d'),
            'الموظف': name,
            'الدخول': clock_in,
            'الخروج': clock_out,
            'التاخير': Late,
            'الحالة': status
        })

# تحويل النتيجة إلى DataFrame
report_df = pd.DataFrame(report_data)


# 🔀 ترتيب حسب الموظف ثم التاريخ
report_df = report_df.sort_values(by=['الموظف', 'التاريخ'])


# حفظ التقرير إلى ملف Excel
report_df.to_excel('تقرير_الحضور_أبريل2025.xlsx', index=False)

print("✅ تم إنشاء التقرير بنجاح: تقرير_الحضور_أبريل2025.xlsx")
