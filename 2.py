import pandas as pd
from datetime import datetime
import re

def convert_to_minutes(value):
    if pd.isna(value):
        return 0
    try:
        match = re.match(r'^(\d{1,2}):(\d{2})$', str(value).strip())
        if match:
            hours, minutes = map(int, match.groups())
            return hours * 60 + minutes
        else:
            return int(value)
    except:
        return 0

def format_minutes_arabic(total_minutes):
    days = total_minutes // (24 * 60)
    hours = (total_minutes % (24 * 60)) // 60
    minutes = total_minutes % 60
    return f"{days} يوم  {hours} ساعة  {minutes} دقيقة"

# تحميل ملف الحضور
file_path = '55.xlsx'
df = pd.read_excel(file_path)

# تحويل التاريخ
df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)

# تحديد الشهر المطلوب
start_date = datetime(2025, 4, 1)
end_date = datetime(2025, 4, 30)
all_days = pd.date_range(start=start_date, end=end_date, freq='D')
working_days = all_days[all_days.weekday != 4]  # الجمعة

# كل التركيبات الممكنة بين الموظفين وأيام العمل
employees = df['Name'].dropna().unique()
multi_index = pd.MultiIndex.from_product([employees, working_days], names=['الموظف', 'التاريخ'])
full_df = pd.DataFrame(index=multi_index).reset_index()

# تجهيز df الأصلي
df['التاريخ'] = df['Date'].dt.normalize()
df_reduced = df[['Name', 'Date', 'Clock In', 'Clock Out', 'Late']].rename(
    columns={'Name': 'الموظف', 'Date': 'التاريخ', 'Clock In': 'الدخول', 'Clock Out': 'الخروج', 'Late': 'التاخير'}
)

# دمج بين الجدولين (full_df وdf_reduced)
merged = pd.merge(full_df, df_reduced, how='left', on=['الموظف', 'التاريخ'])
merged['الحالة'] = merged['الدخول'].apply(lambda x: 'حاضر' if pd.notna(x) else 'غائب')

# ترتيب البيانات
merged = merged.sort_values(by=['الموظف', 'التاريخ'])

# حساب الملخص
summary = merged.groupby('الموظف').agg({
    'التاخير': lambda x: format_minutes_arabic(sum(convert_to_minutes(v) for v in x)),
    'الحالة': lambda x: (x == 'غائب').sum()
}).rename(columns={'التاخير': 'إجمالي التأخير', 'الحالة': 'عدد أيام الغياب'}).reset_index()

# حفظ الملف
with pd.ExcelWriter('تقرير_الحضور_مع_الملخص.xlsx') as writer:
    merged.to_excel(writer, sheet_name='بيانات الحضور', index=False)
    summary.to_excel(writer, sheet_name='الملخص', index=False)

print("✅ تم حفظ التقرير بنجاح وبسرعة 🚀")
