import pandas as pd
import os
from datetime import datetime
from tabulate import tabulate
import arabic_reshaper
from bidi.algorithm import get_display
import unicodedata

# Get today's date in the desired format
today_str = datetime.today().strftime('%d-%m-%Y')

# Define folder and file paths
desktop_path = os.path.expanduser('~/Desktop')
folder_name = f'jerusalem-{today_str}'
folder_path = os.path.join(desktop_path, folder_name)

# Create the folder if it doesn't exist
os.makedirs(folder_path, exist_ok=True)

# Set input and output file paths
input_filename = f'jerusalem-{today_str}.csv'
output_filename = f'{today_str} توزيع غياب القدس.xlsx'
input_path = os.path.join(folder_path, input_filename)
output_path = os.path.join(folder_path, output_filename)

# Load CSV file with semicolon as delimiter
df = pd.read_csv(input_path, encoding='utf-8-sig', sep=';')

# Normalize and clean column names
df.columns = [unicodedata.normalize("NFKC", col).strip() for col in df.columns]

# Debug: show actual column names
print("Actual columns in the CSV file:")
for col in df.columns:
    print(repr(col))

# Define required columns
columns_to_extract = ['اسم الطالب', 'رقم الجلوس', 'الفرع','الجنس', 'القاعة']

# Check for missing columns
missing_columns = [col for col in columns_to_extract if col not in df.columns]
if missing_columns:
    raise KeyError(f"The following required columns are missing from the CSV: {missing_columns}")

# Extract the desired columns
df_extracted = df[columns_to_extract]

# Function to fix Arabic text display for console printing only
def fix_arabic_column(df, columns):
    for col in columns:
        df[col] = df[col].apply(
            lambda x: get_display(arabic_reshaper.reshape(str(x))) if pd.notnull(x) else x
        )
    return df

# Check for duplicates in 'رقم الجلوس'
duplicate_ids = df_extracted[df_extracted.duplicated(subset='رقم الجلوس', keep=False)]

has_duplicates = not duplicate_ids.empty
if has_duplicates:
    print("\n" + get_display(arabic_reshaper.reshape("تحذير: توجد أرقام جلوس مكررة!")))
    
    duplicate_ids_rtl = duplicate_ids.copy()
    duplicate_ids_rtl = fix_arabic_column(duplicate_ids_rtl, ['اسم الطالب', 'رقم الجلوس', 'الفرع', 'القاعة', 'الجنس'])
    duplicate_ids_rtl = duplicate_ids_rtl[['رقم الجلوس', 'اسم الطالب', 'الجنس', 'الفرع', 'القاعة']]

    print(tabulate(
        duplicate_ids_rtl,
        headers=[get_display(arabic_reshaper.reshape(h)) for h in ['رقم الجلوس', 'اسم الطالب', 'الجنس', 'الفرع', 'القاعة']],
        tablefmt='fancy_grid',
        showindex=False,
        colalign=("center", "center", "center", "center", "center")
    ))
else:
    print("\n" + get_display(arabic_reshaper.reshape("لا توجد أرقام جلوس مكررة.")))

# Count of الطلاب by branch and hall
count_by_branch = df_extracted.groupby('الفرع')['اسم الطالب'].count().reset_index()
count_by_branch.columns = ['الفرع', 'عدد الطلاب']

count_by_hall = df_extracted.groupby('القاعة')['اسم الطالب'].count().reset_index()
count_by_hall.columns = ['القاعة', 'عدد الطلاب']

# Sort by الفرع
df_extracted = df_extracted.sort_values(by='الفرع')

# Save to Excel with multiple sheets
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    df_extracted.to_excel(writer, sheet_name='الطلاب', index=False)
    count_by_branch.to_excel(writer, sheet_name='الأعداد حسب الفرع', index=False)
    count_by_hall.to_excel(writer, sheet_name='الأعداد حسب القاعة', index=False)
    if has_duplicates:
        duplicate_ids.to_excel(writer, sheet_name='تكرار رقم الجلوس', index=False)

print("\nFile saved successfully with multiple sheets:", output_path)

# Console display for counts
print("\n" + get_display(arabic_reshaper.reshape("عدد الطلاب لكل فرع:")))
count_by_branch_rtl = count_by_branch.copy()
count_by_branch_rtl = fix_arabic_column(count_by_branch_rtl, ['الفرع'])
count_by_branch_rtl = count_by_branch_rtl[['عدد الطلاب', 'الفرع']]
print(tabulate(
    count_by_branch_rtl,
    headers=[get_display(arabic_reshaper.reshape('عدد الطلاب')), get_display(arabic_reshaper.reshape('الفرع'))],
    tablefmt='fancy_grid',
    showindex=False,
    colalign=("center", "center")
))

print("\n" + get_display(arabic_reshaper.reshape("عدد الطلاب لكل قاعة:")))
count_by_hall_rtl = count_by_hall.copy()
count_by_hall_rtl = fix_arabic_column(count_by_hall_rtl, ['القاعة'])
count_by_hall_rtl = count_by_hall_rtl[['عدد الطلاب', 'القاعة']]
print(tabulate(
    count_by_hall_rtl,
    headers=[get_display(arabic_reshaper.reshape('عدد الطلاب')), get_display(arabic_reshaper.reshape('القاعة'))],
    tablefmt='fancy_grid',
    showindex=False,
    colalign=("center", "center")
))
