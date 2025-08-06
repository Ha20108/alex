def convert_arabic_numbers_to_english(input_string):
    arabic_to_english = {'٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4', '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9'}
    return ''.join(arabic_to_english.get(i, i) for i in input_string)

def convert_to_arabic_numbers(input_string):
    arabic_numbers = {'0': '٠', '1': '١', '2': '٢', '3': '٣', '4': '٤', '5': '٥', '6': '٦', '7': '٧', '8': '٨', '9': '٩'}
    return ''.join(arabic_numbers.get(i, i) for i in input_string)

# تحويل الأرقام إلى العربية
print([convert_to_arabic_numbers("ملاحظات 1")])
print(convert_to_arabic_numbers("ملاحظات 12"))
print(convert_to_arabic_numbers("ملاحظات 2"))
