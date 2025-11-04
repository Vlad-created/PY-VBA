import win32com.client as win32
from datetime import datetime
from dateutil.relativedelta import relativedelta

# === Setup ===
file_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\Py Check\data.xlsx"
excel = win32.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open(file_path)
ws = wb.Sheets("Detail1")

# === Get last 4 months as strings ===
today = datetime.today()
last_4_months = [str((today - relativedelta(months=i)).month) for i in range(4)]

# === Apply AutoFilter on column H (column 8) ===
ws.Range("H1").AutoFilter(Field=8, Criteria1=last_4_months, Operator=7)

# === Delete visible rows using SpecialCells ===
try:
    visible_range = ws.Range("A2:A" + str(ws.UsedRange.Rows.Count)).SpecialCells(12)  # 12 = xlCellTypeVisible
    visible_range.EntireRow.Delete()
    print(f"🧹 Deleted filtered rows for months: {last_4_months}")
except Exception as e:
    print(f"⚠️ No visible rows to delete or error occurred: {e}")

# === Reset filters safely ===
try:
    # Check if there's still data below header
    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # -4162 = xlUp
    if last_row > 1:
        ws.Range(f"A1:H{last_row}").AutoFilter(Field=8)  # Reapply filter to valid range
        if ws.FilterMode and ws.ShowAllData:
            ws.ShowAllData()
    if ws.AutoFilterMode:
        ws.AutoFilterMode = False
    print("🚫 Filters fully cleared.")
except Exception as e:
    print(f"⚠️ Error resetting filters: {e}")
