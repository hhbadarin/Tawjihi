import pandas as pd
import os
from tabulate import tabulate
import arabic_reshaper
from bidi.algorithm import get_display

# Define file paths
input_path = os.path.expanduser('~/Desktop/yatta.csv')
output_path = os.path.expanduser('~/Desktop/yatta-Abs.xlsx')

# Load CSV file
df = pd.read_csv(input_path, encoding='utf-8-sig')

# Extract the desired columns
columns_to_extract = ['رقم الجلوس', 'اسم الطالب', 'الجنس', 'القاعة', 'الفرع']
df_extracted = df[columns_to_extract]

# Save to Excel
df_extracted.to_excel(output_path, index=False)

print("File saved successfully:", output_path)

# Function to fix Arabic text display in specified columns
def fix_arabic_column(df, columns):
    for col in columns:
        df[col] = df[col].apply(
            lambda x: get_display(arabic_reshaper.reshape(str(x))) if pd.notnull(x) else x
        )
    return df

# Count of students by branch (الفرع)
count_by_branch = df_extracted.groupby('الفرع')['اسم الطالب'].count().reset_index()
count_by_branch.columns = ['الفرع', 'عدد الطلاب']
count_by_branch = fix_arabic_column(count_by_branch, ['الفرع'])

# Count of students by hall (القاعة)
count_by_hall = df_extracted.groupby('القاعة')['اسم الطالب'].count().reset_index()
count_by_hall.columns = ['القاعة', 'عدد الطلاب']
count_by_hall = fix_arabic_column(count_by_hall, ['القاعة'])

# Print branch count table
print("\n" + get_display(arabic_reshaper.reshape("عدد الطلاب لكل فرع:")))
count_by_branch_rtl = count_by_branch[['عدد الطلاب', 'الفرع']]
print(tabulate(
    count_by_branch_rtl,
    headers=[get_display(arabic_reshaper.reshape('عدد الطلاب')), get_display(arabic_reshaper.reshape('الفرع'))],
    tablefmt='fancy_grid',
    showindex=False,
    colalign=("center", "center")
))

# Print hall count table
print("\n" + get_display(arabic_reshaper.reshape("عدد الطلاب لكل قاعة:")))
count_by_hall_rtl = count_by_hall[['عدد الطلاب', 'القاعة']]
print(tabulate(
    count_by_hall_rtl,
    headers=[get_display(arabic_reshaper.reshape('عدد الطلاب')), get_display(arabic_reshaper.reshape('القاعة'))],
    tablefmt='fancy_grid',
    showindex=False,
    colalign=("center", "center")
))
