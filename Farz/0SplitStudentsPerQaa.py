import pandas as pd

# Load the Excel file
file_path = '~/Desktop/source.xlsm'
df = pd.read_excel(file_path)

# Clean up column names
df.columns = df.columns.str.strip()

# Keep only the 5 required columns
columns_to_keep = ['الإسم', 'رقم الجلوس', 'الفرع', 'الجنس', 'القاعة']
df = df[columns_to_keep]

# Write each hall to its own sheet
with pd.ExcelWriter('~/Desktop/توزيع الغياب حسب القاعة.xlsx', engine='openpyxl') as writer:
    for hall, group in df.groupby('القاعة'):
        # Drop rows with missing names (optional)
        hall_df = group.dropna(subset=['الإسم']).reset_index(drop=True)

        # Ensure Excel sheet name is valid
        safe_sheet_name = str(hall).strip()[:31]
        hall_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
