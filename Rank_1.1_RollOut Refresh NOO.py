import time
import win32com.client as win32

file_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\RollOut.xlsx"

# Start Excel
excel = win32.Dispatch("Excel.Application")
excel.Visible = False

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
time.sleep(30)

# Save and close
wb.Save()
wb.Close()
excel.Quit()

print("\nThe following queries were refreshed and saved:")
for q in query_names:
    print(" -", q)
