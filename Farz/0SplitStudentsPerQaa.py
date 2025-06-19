import pandas as pd

# Load the Excel file
file_path = '~/Desktop/source.xlsm'
df = pd.read_excel(file_path)

# Clean up column names
df.columns = df.columns.str.strip()

# Create Excel writer
with pd.ExcelWriter('~/Desktop/توزيع الغياب حسب القاعة.xlsx', engine='openpyxl') as writer:
    for hall, group in df.groupby('القاعة'):
        # Drop any rows where الإسم is missing (optional)
        hall_df = group.dropna(subset=['الإسم']).reset_index(drop=True)
        
        # Limit sheet name to 31 characters (Excel limit)
        safe_sheet_name = str(hall)[:31]

        # Write to sheet
        hall_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
