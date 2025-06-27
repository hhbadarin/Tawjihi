import pandas as pd

# File paths
input_path = r'C:\Users\hp\Desktop\yatta-24-06-2025.csv'
output_path = r'C:\Users\hp\Desktop\غياب يطا جاهز 24-06.xlsx'

# Read the CSV file
df = pd.read_csv(input_path, encoding='utf-8', sep=';')  # Change sep if needed

# Extract required columns
columns_needed = ['اسم الطالب', 'رقم الجلوس', 'الفرع','الجنس','القاعة']
df_filtered = df[columns_needed]

# Save to XLSX
df_filtered.to_excel(output_path, index=False, engine='openpyxl')
