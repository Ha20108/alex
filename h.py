import pandas as pd
from datetime import datetime, timedelta
import re

def convert_to_minutes(value):
    if pd.isna(value):
        return 0
    try:
        # إذا كانت القيمة على شكل 00:44
        match = re.match(r'^(\d{1,2}):(\d{2})$', str(value).strip())
        if match:
            hours, minutes = map(int, match.groups())
            return hours * 60 + minutes
        else:
            # ممكن تكون قيمة مجرد عدد دقائق فقط
            return int(value)
    except:
        return 0

def format_minutes_arabic(total_minutes):
    days = total_minutes // (24 * 60)
    hours = (total_minutes % (24 * 60)) // 60
    minutes = total_minutes % 60
    return f"{days} يوم  {hours} ساعة  {minutes} دقيقة"


# تحميل ملف الحضور
file_path = '55.xlsx'  # ← ضع اسم الملف الذي حفظته هنا
df = pd.read_excel(file_path)

# تحويل عمود التاريخ إلى نوع تاريخ
df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)

# تحديد الشهر المطلوب (أبريل 2025)
start_date = datetime(2025, 4, 13)
end_date = datetime(2025, 4, 16)

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

# إنشاء الورقة الملخص التي تحتوي على إجمالي التأخير وعدد أيام الغياب لكل موظف
summary_data = []

for name, group in report_df.groupby('الموظف'):
    total_minutes = group['التاخير'].apply(convert_to_minutes).sum()
    total_absent = (group['الحالة'] == 'غائب').sum()
    summary_data.append({
        'الموظف': name,
        'إجمالي التأخير (بالدقائق)': format_minutes_arabic(total_minutes),
        'عدد أيام الغياب': total_absent
    })

summary_df = pd.DataFrame(summary_data)

# حفظ البيانات في ملف Excel مع ورقتين:
# 1. ورقة تحتوي على بيانات الحضور اليومية.
# 2. ورقة تحتوي على الملخص.
with pd.ExcelWriter('تقرير_الحضور_مع_الملخص.xlsx') as writer:
    # حفظ بيانات الحضور التفصيلية بدون تغيير
    report_df.to_excel(writer, sheet_name='بيانات الحضور', index=False)
    
    # حفظ الملخص في ورقة منفصلة
    summary_df.to_excel(writer, sheet_name='الملخص', index=False)

print("✅ تم حفظ التقرير بنجاح مع ملخص في ورقة منفصلة: تقرير_الحضور_مع_الملخص.xlsx")
