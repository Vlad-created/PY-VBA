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

# Optional: force synchronous refresh for OLEDB connections
for conn in wb.Connections:
    try:
        conn.OLEDBConnection.BackgroundQuery = False
    except:
        pass

# Refresh all connections
wb.RefreshAll()

# Wait until all background refreshes are done
print("\nWaiting for refresh to complete...")
while any(conn.Refreshing for conn in wb.Connections):
    time.sleep(1)

# Save only — do NOT close
wb.Save()

print("\nThe following queries were refreshed and saved:")
for q in query_names:
    print(" -", q)

print("\nExcel remains open. Close it manually when you're done.")
