import pandas as pd
import os

# Path to your Excel files folder
folder_path = os.path.expanduser("~/Desktop/ExcelFiles")

# Get all Excel files in the folder
excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]

# Create an empty DataFrame
merged_df = pd.DataFrame()

# Loop through the Excel files and append to merged_df
for file in excel_files:
    file_path = os.path.join(folder_path, file)
    df = pd.read_excel(file_path)
    df['Source File'] = file  # Optional: track source file
    merged_df = pd.concat([merged_df, df], ignore_index=True)

# Save the merged data
output_path = os.path.expanduser("~/Desktop/Merged.xlsx")
merged_df.to_excel(output_path, index=False)

print(f"Merged {len(excel_files)} files into: {output_path}")
