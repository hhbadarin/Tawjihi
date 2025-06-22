import pandas as pd
import os
from tabulate import tabulate
import arabic_reshaper
from bidi.algorithm import get_display

# Define file paths
input_path = os.path.expanduser('~/Desktop/nhebron.csv')
output_path = os.path.expanduser('~/Desktop/nhebron-Abs.xlsx')

# Load CSV file
df = pd.read_csv(input_path, encoding='utf-8-sig')

# Extract the desired columns
columns_to_extract = ['اسم الطالب','رقم الجلوس', 'الجنس', 'القاعة', 'الفرع']
df_extracted = df[columns_to_extract]

# Count of الطلاب by branch (الفرع) — do NOT reshape here
count_by_branch = df_extracted.groupby('الفرع')['اسم الطالب'].count().reset_index()
count_by_branch.columns = ['الفرع', 'عدد الطلاب']

# Count of الطلاب by hall (القاعة) — do NOT reshape here
count_by_hall = df_extracted.groupby('القاعة')['اسم الطالب'].count().reset_index()
count_by_hall.columns = ['القاعة', 'عدد الطلاب']

# Save to Excel — normal Arabic text, no reshaping applied here
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    df_extracted.to_excel(writer, sheet_name='الطلاب', index=False)
    count_by_branch.to_excel(writer, sheet_name='الأعداد حسب الفرع', index=False)
    count_by_hall.to_excel(writer, sheet_name='الأعداد حسب القاعة', index=False)

print("File saved successfully with multiple sheets:", output_path)

# Function to fix Arabic text display for console printing only
def fix_arabic_column(df, columns):
    for col in columns:
        df[col] = df[col].apply(
            lambda x: get_display(arabic_reshaper.reshape(str(x))) if pd.notnull(x) else x
        )
    return df

# Prepare and print branch count table (reshaped for console)
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

# Prepare and print hall count table (reshaped for console)
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
