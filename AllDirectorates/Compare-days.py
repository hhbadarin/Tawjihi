import pandas as pd
import os
from datetime import datetime

# === User Input: Directorate name ===
directorate = input("Enter the directorate english name (e.g., hebron): ").strip()
directorate_ar = input("Enter the directorate arabic name (e.g., الخليل): ").strip()
# === Today's Date ===
today = datetime.now()
today_str = today.strftime("%d-%m-%Y")

# === Folder path for today ===
folder_name = f"{directorate}-{today_str}"
desktop = os.path.expanduser("~/Desktop")
folder_path = os.path.join(desktop, folder_name)

# === File paths ===
old_file = os.path.join(folder_path, f"{directorate}-03-07-2025.csv")  # Fixed old file date
new_file = os.path.join(folder_path, f"{directorate}-{today_str}.csv")  # Today's file

# === Read CSVs ===
df_old = pd.read_csv(old_file, sep=';')
df_new = pd.read_csv(new_file, sep=';')

# === Normalize 'رقم الجلوس' ===
df_old["رقم الجلوس"] = df_old["رقم الجلوس"].astype(str).str.strip()
df_new["رقم الجلوس"] = df_new["رقم الجلوس"].astype(str).str.strip()

# === Find new entries ===
new_entries = df_new[~df_new["رقم الجلوس"].isin(df_old["رقم الجلوس"])]

# === Keep only specified columns ===
columns_to_keep = ["اسم الطالب", "رقم الجلوس", "الفرع", "الجنس", "القاعة"]
new_entries = new_entries[columns_to_keep]

# === Save result ===
output_path = os.path.join(folder_path,f"{directorate_ar}-الغيابات الجديدة-{today_str}.csv")
new_entries.to_csv(output_path, index=False, sep=';', encoding='utf-8-sig')

print(f"✅ New entries saved to: {output_path}")
