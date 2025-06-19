import pandas as pd
import os

# Path to your Excel file
file_path = os.path.expanduser("~/Desktop/SPlit.xlsx")

# Load the Excel file
excel_file = pd.ExcelFile(file_path)

# List to collect dataframes
dfs = []

# Loop through each sheet
for sheet_name in excel_file.sheet_names:
    df = excel_file.parse(sheet_name)
    df['Source Sheet'] = sheet_name  # Optional: to track origin
    dfs.append(df)

# Combine all sheets into one DataFrame (columns aligned by name)
combined_df = pd.concat(dfs, ignore_index=True)

# Save to CSV
output_path = os.path.expanduser("~/Desktop/Combined_Sheets.csv")
combined_df.to_csv(output_path, index=False, encoding='utf-8-sig')

print(f"✅ All sheets combined into CSV: {output_path}")
