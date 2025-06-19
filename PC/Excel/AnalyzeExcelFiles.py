import pandas as pd
import os

# Path to your Excel file
file_path = os.path.expanduser("~/Desktop/Split.xlsx")

# Load the Excel file (only metadata at this point)
excel_file = pd.ExcelFile(file_path)

# Dictionary to hold sheet name -> set of columns
sheet_columns = {}

# Read column names from each sheet
for sheet in excel_file.sheet_names:
    df = excel_file.parse(sheet, nrows=0)  # Only read headers
    sheet_columns[sheet] = set(df.columns)

# Find union of all columns
all_columns = set().union(*sheet_columns.values())

# Report differences
print("=== Column Comparison Across Sheets ===\n")
for sheet, columns in sheet_columns.items():
    missing = all_columns - columns
    extra = columns - all_columns
    print(f"Sheet: {sheet}")
    print(f" - Columns: {sorted(columns)}")
    if missing:
        print(f"   ⚠️ Missing: {sorted(missing)}")
    if extra:
        print(f"   ➕ Extra: {sorted(extra)}")
    print()

# Optional: Export column summary to Excel
summary = pd.DataFrame.from_dict(
    {sheet: list(cols) for sheet, cols in sheet_columns.items()},
    orient='index'
).transpose()

summary_path = os.path.expanduser("~/Desktop/Sheet_Column_Comparison.xlsx")
summary.to_excel(summary_path, index=False)
print(f"Column summary saved to: {summary_path}")
