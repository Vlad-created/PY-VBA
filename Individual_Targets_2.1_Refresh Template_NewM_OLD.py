import time
import win32com.client as win32

file_path = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets\Templete_newmonth.xlsx"

# Start Excel
excel = win32.Dispatch("Excel.Application")
excel.Visible = True  # Show Excel window

# Open workbook
wb = excel.Workbooks.Open(file_path)

# Collect all Power Query names BEFORE refresh
query_names = [q.Name for q in wb.Queries]

print("Queries found in workbook:")
for q in query_names:
    print(" -", q)

# Refresh all connections
wb.RefreshAll()

# Wait for refresh to complete
time.sleep(45)

# Save only — do NOT close
wb.Save()

print("\nThe following queries were refreshed and saved:")
for q in query_names:
    print(" -", q)

# Excel remains open with workbook visible
