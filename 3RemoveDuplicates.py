import pandas as pd

# Load only the required columns from the updated Excel file
df = pd.read_excel(
    "~/Desktop/Source_Updated.xlsx",
    usecols=["الرقم الوطني", "المدرسة"],
    dtype=str  # Preserve leading zeros
)

# Drop duplicate rows based on both columns
df_unique = df.drop_duplicates()

# Rename columns
df_unique.rename(columns={
    "الرقم الوطني": "ID",
    "المدرسة": "SchoolName"
}, inplace=True)

# Save the result to CSV
df_unique.to_csv("~\Documents\GitHub\Microsoft-Graph\Mailbox\Exams\Temp.csv", index=False, encoding='utf-8-sig')

print("✅ Temp.csv has been saved with columns renamed to ID and SchoolName.")
