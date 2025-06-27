import xlwings as xw
import pandas as pd
import os

# === Configurable Parameters ===
exam_date = "26/6"  # 👈 Change this to the correct exam date string

# === File Paths ===
excel_path = "C:/Users/hp/Desktop/shebron-26-06-2025/جنوب الخليل غياب تراكمي 24-06.xlsx"
csv_path = "C:/Users/hp/Desktop/shebron-26-06-2025/shebron-26-06-2025.csv"
output_path = "C:/Users/hp/Desktop/غياب تراكمي جنوب الخليل 26-06 حسب برنامج هيثم.xlsx"

# === Load CSV absentee data ===
df = pd.read_csv(csv_path, delimiter=";")
df.columns = df.columns.str.strip().str.replace('\ufeff', '')
df["رقم الجلوس"] = df["رقم الجلوس"].dropna().apply(lambda x: str(int(float(x)))).str.strip()
df["اسم الطالب"] = df["اسم الطالب"].fillna("").astype(str).str.strip()
df["الجنس"] = df["الجنس"].fillna("").astype(str).str.strip()
df["القاعة"] = df["القاعة"].fillna("").astype(str).str.strip()

# === Build mapping of absentees by رقم الجلوس
absentees = {
    sid: {
        "name": name,
        "gender": gender,
        "room": room
    }
    for sid, name, gender, room in zip(
        df["رقم الجلوس"], df["اسم الطالب"], df["الجنس"], df["القاعة"]
    )
}

# === Group absentees by room name (or branch)
grouped_absentees = {}
for sid, info in absentees.items():
    branch = info.get("room", "").strip()
    if branch not in grouped_absentees:
        grouped_absentees[branch] = []
    grouped_absentees[branch].append((sid, info))

# === Open Excel workbook
app = xw.App(visible=False)
wb = app.books.open(excel_path)

# === Build sheet name → sheet object map
sheet_map = {sheet.name.strip(): sheet for sheet in wb.sheets}

# === Column positions
start_row = 14
serial_col = 2   # B
name_col = 3     # C
id_col = 4       # D
gender_col = 16  # P
room_col = 17    # Q

# === Counter
total_added = 0
total_marked = 0

# === Process each branch
for branch, students in grouped_absentees.items():
    # Match branch to sheet
    matched_sheet = None
    for sheet_name in sheet_map:
        if branch in sheet_name or sheet_name in branch:
            matched_sheet = sheet_map[sheet_name]
            break

    if not matched_sheet:
        print(f"⚠️ Skipping branch '{branch}': No matching sheet found.")
        continue

    ws = matched_sheet

    # === Find exam column
    x_col = None
    for col in range(1, ws.used_range.last_cell.column + 1):
        val = ws.cells(13, col).value
        if val and str(val).strip() == exam_date:
            x_col = col
            break
    if not x_col:
        print(f"❌ Date column '{exam_date}' not found in sheet '{ws.name}'")
        continue

    # === Collect existing students
    existing_ids = {}
    row = start_row
    serial = 1
    while True:
        val = ws.cells(row, id_col).value
        if not val:
            break
        try:
            sid = str(int(float(val))).strip()
            existing_ids[sid] = row
            ws.cells(row, serial_col).value = serial
            serial += 1
        except:
            pass
        row += 1

    last_student_row = row - 1
    added = 0
    marked = 0

    # === Process each student
    for sid, info in students:
        name = info["name"]
        gender = info["gender"]
        room = info["room"]

        if sid in existing_ids:
            row = existing_ids[sid]
            ws.cells(row, x_col).value = "X"
            marked += 1
        else:
            insert_row = last_student_row + 1
            ws.api.Rows(insert_row).Insert()

            ws.cells(insert_row, serial_col).value = serial
            ws.cells(insert_row, name_col).value = name
            ws.cells(insert_row, id_col).value = sid
            ws.cells(insert_row, x_col).value = "X"
            ws.cells(insert_row, gender_col).value = gender
            ws.cells(insert_row, room_col).value = room

            serial += 1
            last_student_row += 1
            added += 1

    total_marked += marked
    total_added += added
    print(f"✅ Sheet '{ws.name}': marked {marked}, added {added}")

# === Save & close
wb.save(output_path)
wb.close()
app.quit()

print(f"🎉 Done for {exam_date}: Total marked = {total_marked}, Total added = {total_added}")
