import xlwings as xw
import pandas as pd

# === Configurable Parameters ===
exam_date = "26/6"  # 👈 Change this to any date like "23/6", "26/6", etc.

# === File Paths ===
excel_path = "C:/Users/hp/Desktop/hebron-26-06-2025/الخليل غياب تراكمي 26-06.xlsx"
csv_path = "C:/Users/hp/Desktop/hebron-26-06-2025/hebron-26-06-2025.csv"
output_path = "C:/Users/hp/Desktop/غياب تراكمي الخليل 26-06 حسب برنامج هيثم.xlsx"

# === Load Excel via xlwings (preserves text box and footer) ===
app = xw.App(visible=False)
wb = app.books.open(excel_path)
ws = wb.sheets[0]

# === Load and clean absentee data ===
df = pd.read_csv(csv_path, delimiter=";")
df.columns = df.columns.str.strip().str.replace('\ufeff', '')
df["رقم الجلوس"] = df["رقم الجلوس"].dropna().apply(lambda x: str(int(float(x)))).str.strip()
df["اسم الطالب"] = df["اسم الطالب"].fillna("").astype(str).str.strip()
df["الجنس"] = df["الجنس"].fillna("").astype(str).str.strip()
df["القاعة"] = df["القاعة"].fillna("").astype(str).str.strip()

# Build mapping of absentees by ID
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
absentee_ids = set(absentees.keys())

# === Column Positions (fixed)
start_row = 14
serial_col = 2   # B
name_col = 3     # C
id_col = 4       # D
gender_col = 16  # P
room_col = 17    # Q

# === Find exam date column dynamically from row 13
x_col = None
for col in range(1, ws.used_range.last_cell.column + 1):
    val = ws.cells(13, col).value
    if val and str(val).strip() == exam_date:
        x_col = col
        break

if not x_col:
    app.quit()
    raise ValueError(f"❌ Column for exam date '{exam_date}' not found in row 13.")

# === Collect existing student IDs
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
        # Also update serial numbers
        ws.cells(row, serial_col).value = serial
        serial += 1
    except:
        pass
    row += 1

last_student_row = row - 1
added = 0
marked = 0

# === Process absentees
for sid in absentee_ids:
    student = absentees[sid]
    name = student["name"]
    gender = student["gender"]
    room = student["room"]

    if sid in existing_ids:
        row = existing_ids[sid]
        ws.cells(row, x_col).value = "X"
        marked += 1
    else:
        insert_row = last_student_row + 1
        ws.api.Rows(insert_row).Insert()  # Insert new row (footer stays)

        ws.cells(insert_row, serial_col).value = serial
        ws.cells(insert_row, name_col).value = name
        ws.cells(insert_row, id_col).value = sid
        ws.cells(insert_row, x_col).value = "X"
        ws.cells(insert_row, gender_col).value = gender
        ws.cells(insert_row, room_col).value = room

        serial += 1
        last_student_row += 1
        added += 1

# === Save workbook (preserves formatting, text boxes, footer)
wb.save(output_path)
wb.close()
app.quit()

print(f"✅ Done for {exam_date}: Marked {marked} existing, added {added} new.")
