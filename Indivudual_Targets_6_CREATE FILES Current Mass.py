import os
import calendar
from datetime import datetime
import shutil
import win32com.client as win32

# ---------------------------------------------------
# 1. Build dynamic file path
# ---------------------------------------------------
base_path = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets\Regions"

current_month = calendar.month_name[datetime.now().month]
file_name = f"Individual_targets_{current_month}_Total.xlsx"
file_path = os.path.join(base_path, file_name)

print("Main file:", file_path)

# ---------------------------------------------------
# 2. Create a copy with "_" at the end
# ---------------------------------------------------
base, ext = os.path.splitext(file_path)
copy_path = f"{base}_{ext}"

shutil.copy2(file_path, copy_path)
print("Created working copy:", copy_path)

# ---------------------------------------------------
# 3. Start Excel
# ---------------------------------------------------
excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(copy_path)
ws = wb.Worksheets("individual")

# ---------------------------------------------------
# 4. Unhide ALL columns in sheet "individual"
# ---------------------------------------------------
ws.Columns.Hidden = False

# ---------------------------------------------------
# 5. Copy all and paste as VALUES
# ---------------------------------------------------
used_range = ws.UsedRange
used_range.Copy()
ws.Range("A1").PasteSpecial(Paste=-4163)  # xlPasteValues
excel.CutCopyMode = False

# ---------------------------------------------------
# 6. Delete columns DY → EK
# ---------------------------------------------------
ws.Range("DY:EK").Delete()

# ---------------------------------------------------
# 7. Delete all other sheets except "individual"
# ---------------------------------------------------
for sh in list(wb.Worksheets):
    if sh.Name != "individual":
        sh.Delete()

wb.Save()

# ---------------------------------------------------
# 8. Filter by Region and create region files
# ---------------------------------------------------
regions = [
    "Северозапад", "София 1", "София 2", "София 3", "София 4",
    "Бургас", "Варна 1", "Варна 2", "Велико Търново",
    "Пловдив 1", "Пловдив 2", "Пловдив 3", "Стара Загора"
]

xlUp = -4162
xlToLeft = -4159
xlCellTypeVisible = 12

last_row = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
last_col = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

data_range = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col))

for region in regions:

    if ws.AutoFilterMode:
        ws.AutoFilterMode = False

    data_range.AutoFilter(Field=5, Criteria1=region)

    try:
        visible = data_range.SpecialCells(xlCellTypeVisible)
    except Exception:
        print(f"No rows for region: {region}")
        continue

    new_wb = excel.Workbooks.Add()
    new_ws = new_wb.Worksheets(1)
    new_ws.Name = "individual"

    visible.Copy(new_ws.Range("A1"))

    region_file = f"{base}_{region}{ext}"
    new_wb.SaveAs(region_file)
    new_wb.Close(SaveChanges=False)

    print("Created:", region_file)

if ws.AutoFilterMode:
    ws.AutoFilterMode = False

wb.Save()
wb.Close(SaveChanges=True)
excel.Quit()

print("All done.")
