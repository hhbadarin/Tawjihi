import pandas as pd
from datetime import datetime
import os

# Get today's date
today = datetime.today()
today_str_full = today.strftime('%d-%m-%Y')  # e.g., 27-06-2025
today_str_short = today.strftime('%d-%m')    # e.g., 27-06

# Get user input
district_en = input("Enter the English name of the directorate (e.g., hebron): ").strip()
district_ar = input("Enter the Arabic name of the directorate (e.g., الخليل): ").strip()

# Get current username to build Desktop path
username = os.getlogin()

# Define the folder path on Desktop
folder_path = f"C:/Users/{username}/Desktop/{district_en}-{today_str_full}"

# Define input and output file paths
input_path = os.path.join(folder_path, f"{district_en}-{today_str_full}.csv")
output_path = os.path.join(folder_path, f"غياب {district_ar} حسب الفرع {today_str_short}.xlsx")

# Read the CSV file
df = pd.read_csv(input_path, encoding='utf-8', sep=';')  # Adjust sep if needed

# Extract required columns and sort by الفرع
columns_needed = ['اسم الطالب', 'رقم الجلوس', 'الفرع', 'الجنس', 'القاعة']
df_filtered = df[columns_needed].sort_values(by='الفرع')

# Group by الفرع and write each group to a separate sheet
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    for branch_name, group in df_filtered.groupby('الفرع'):
        # Sheet names must be <= 31 characters and can't contain these chars: \ / * ? [ ]
        sheet_name = str(branch_name)[:31].replace('/', '-').replace('\\', '-')
        group.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"✔️ تم حفظ الملف بنجاح، وكل فرع في ورقة منفصلة:\n{output_path}")
