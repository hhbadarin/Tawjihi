import pandas as pd
import os
from datetime import datetime
from tabulate import tabulate
import arabic_reshaper
from bidi.algorithm import get_display
import unicodedata
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

# Get user inputs
folder_prefix = input("Please enter the directorate's name in english (e.g., hebron): ").strip()
location_name = input("Please enter the directorate's name in Arabic (e.g., الخليل): ").strip()

# Get today's date
today_str = datetime.today().strftime('%d-%m-%Y')

# Paths and filenames
desktop_path = os.path.expanduser('~/Desktop')
folder_name = f'{folder_prefix}-{today_str}'
folder_path = os.path.join(desktop_path, folder_name)

# Create output folder
os.makedirs(folder_path, exist_ok=True)

# Move CSV from Desktop to folder
input_filename = f'{folder_prefix}-{today_str}.csv'
original_input_path = os.path.join(desktop_path, input_filename)
new_input_path = os.path.join(folder_path, input_filename)

if os.path.exists(original_input_path):
    os.rename(original_input_path, new_input_path)
else:
    raise FileNotFoundError(f"File not found on Desktop: {original_input_path}")

# Set input/output file paths
input_path = new_input_path
output_filename = f'{today_str} توزيع غياب {location_name}.xlsx'
output_path = os.path.join(folder_path, output_filename)

# Load CSV
df = pd.read_csv(input_path, encoding='utf-8-sig', sep=';')
df.columns = [unicodedata.normalize("NFKC", col).strip() for col in df.columns]

# Debug column names
print("Actual columns in the CSV file:")
for col in df.columns:
    print(repr(col))

# Extract required columns
columns_to_extract = ['اسم الطالب', 'رقم الجلوس', 'الفرع', 'الجنس', 'القاعة']
missing_columns = [col for col in columns_to_extract if col not in df.columns]
if missing_columns:
    raise KeyError(f"Missing columns: {missing_columns}")

df_extracted = df[columns_to_extract]

# Arabic fix for printing only
def fix_arabic_column(df, columns):
    for col in columns:
        df[col] = df[col].apply(lambda x: get_display(arabic_reshaper.reshape(str(x))) if pd.notnull(x) else x)
    return df

# Duplicates
duplicate_ids = df_extracted[df_extracted.duplicated(subset='رقم الجلوس', keep=False)]
has_duplicates = not duplicate_ids.empty

if has_duplicates:
    print("\n" + get_display(arabic_reshaper.reshape("تحذير: توجد أرقام جلوس مكررة!")))
    duplicate_ids_rtl = fix_arabic_column(duplicate_ids.copy(), ['اسم الطالب', 'رقم الجلوس', 'الفرع', 'القاعة', 'الجنس'])
    duplicate_ids_rtl = duplicate_ids_rtl[['رقم الجلوس', 'اسم الطالب', 'الجنس', 'الفرع', 'القاعة']]
    print(tabulate(
        duplicate_ids_rtl,
        headers=[get_display(arabic_reshaper.reshape(h)) for h in ['رقم الجلوس', 'اسم الطالب', 'الجنس', 'الفرع', 'القاعة']],
        tablefmt='fancy_grid', showindex=False, colalign=("center",)*5
    ))
else:
    print("\n" + get_display(arabic_reshaper.reshape("لا توجد أرقام جلوس مكررة.")))

# Count summaries
count_by_branch = df_extracted.groupby('الفرع')['اسم الطالب'].count().reset_index()
count_by_branch.columns = ['الفرع', 'عدد الطلاب']

count_by_hall = df_extracted.groupby('القاعة')['اسم الطالب'].count().reset_index()
count_by_hall.columns = ['القاعة', 'عدد الطلاب']

# Sort
df_extracted = df_extracted.sort_values(by='الفرع')

# Excel export
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    df_extracted.to_excel(writer, sheet_name='الطلاب', index=False)
    count_by_branch.to_excel(writer, sheet_name='الأعداد حسب الفرع', index=False)
    count_by_hall.to_excel(writer, sheet_name='الأعداد حسب القاعة', index=False)
    if has_duplicates:
        duplicate_ids.to_excel(writer, sheet_name='تكرار رقم الجلوس', index=False)

print("\n✅ Excel file saved:", output_path)

# Print summaries to console
print("\n" + get_display(arabic_reshaper.reshape("عدد الطلاب لكل فرع:")))
count_by_branch_rtl = fix_arabic_column(count_by_branch.copy(), ['الفرع'])[['عدد الطلاب', 'الفرع']]
print(tabulate(count_by_branch_rtl,
               headers=[get_display(arabic_reshaper.reshape('عدد الطلاب')), get_display(arabic_reshaper.reshape('الفرع'))],
               tablefmt='fancy_grid', showindex=False, colalign=("center", "center")))

print("\n" + get_display(arabic_reshaper.reshape("عدد الطلاب لكل قاعة:")))
count_by_hall_rtl = fix_arabic_column(count_by_hall.copy(), ['القاعة'])[['عدد الطلاب', 'القاعة']]
print(tabulate(count_by_hall_rtl,
               headers=[get_display(arabic_reshaper.reshape('عدد الطلاب')), get_display(arabic_reshaper.reshape('القاعة'))],
               tablefmt='fancy_grid', showindex=False, colalign=("center", "center")))
