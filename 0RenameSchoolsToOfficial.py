import pandas as pd

# Explicitly define data types for both 'الرقم الوطني' and 'الهوية' as string
dtype_mapping = {
    "الرقم الوطني": str,
    "الهوية": str
}

# Read the source Excel file with specified dtypes
source_df = pd.read_excel(
    "~/Desktop/Source.xlsx",
    dtype=dtype_mapping
)

# Read the SchoolsList CSV with ID as string
schools_df = pd.read_csv(
    "~/Documents/GitHub/Microsoft-Graph/Mailbox/Exams/SchoolsList.csv",
    dtype={"ID": str}
)

# Merge the school names using الرقم الوطني ↔ ID
merged_df = source_df.merge(
    schools_df,
    left_on='الرقم الوطني',
    right_on='ID',
    how='left'
)

# Replace اسم المدرسة with SchoolName from the reference
merged_df['المدرسة'] = merged_df['SchoolName']

# Drop ID and SchoolName columns
merged_df.drop(['ID', 'SchoolName'], axis=1, inplace=True)

# Save to Excel using ExcelWriter, preserving date format and avoiding float conversion
with pd.ExcelWriter("~/Desktop/Source_Updated.xlsx", engine='xlsxwriter', datetime_format='yyyy-mm-dd', date_format='yyyy-mm-dd') as writer:
    merged_df.to_excel(writer, index=False)

print("✅ المدرسة column has been updated and الهوية preserved with leading zeros.")
