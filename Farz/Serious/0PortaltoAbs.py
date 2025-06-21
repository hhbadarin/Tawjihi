import pandas as pd
import os
from tabulate import tabulate
import arabic_reshaper
from bidi.algorithm import get_display

# Define file paths
input_path = os.path.expanduser('~/Desktop/hebron.csv')
output_path = os.path.expanduser('~/Desktop/Hebron-Abs.xlsx')

# Load CSV file
df = pd.read_csv(input_path, encoding='utf-8-sig')

# Extract the desired columns
columns_to_extract = ['اسم الطالب','رقم الجلوس', 'الجنس', 'القاعة', 'الفرع']
df_extracted = df[columns_to_extract]

# Save to Excel
df_extracted.to_excel(output_path, index=False)

print("File saved successfully:", output_path)

# Count of students by branch (الفرع)
count_by_branch = df_extracted.groupby('الفرع')['اسم الطالب'].count().reset_index()
count_by_branch.columns = ['الفرع', 'عدد الطلاب']

# Function to fix Arabic text display in specified columns
def fix_arabic_column(df, columns):
    for col in columns:
        df[col] = df[col].apply(
            lambda x: get_display(arabic_reshaper.reshape(str(x))) if pd.notnull(x) else x
        )
    return df

# Apply Arabic text fixing
count_by_branch = fix_arabic_column(count_by_branch, ['الفرع'])

# Print heading properly
print("\n" + get_display(arabic_reshaper.reshape("عدد الطلاب لكل فرع:")))

# Reverse columns order for RTL terminal display
count_by_branch_rtl = count_by_branch[['عدد الطلاب', 'الفرع']]

# Print the table with swapped headers to match reversed columns, center text in columns
print(tabulate(
    count_by_branch_rtl,
    headers=[get_display(arabic_reshaper.reshape('عدد الطلاب')), get_display(arabic_reshaper.reshape('الفرع'))],
    tablefmt='fancy_grid',
    showindex=False,
    colalign=("center", "center")
))
