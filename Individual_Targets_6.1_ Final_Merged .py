import os
import time
import datetime
import win32com.client as win32

folder = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets"

current_month = datetime.datetime.now().strftime("%B")

src_path = None
for filename in os.listdir(folder):
    if (
        filename.startswith("Individual ")
        and filename.endswith("Final_Merged.xlsx")
        and current_month in filename
    ):
        src_path = os.path.join(folder, filename)
        break

if not src_path:
    raise FileNotFoundError("Source file not found.")

print(f"Opening file: {src_path}")

# Start Excel
excel = win32.Dispatch("Excel.Application")
excel.Visible = True  # Show Excel window

# Open workbook
wb = excel.Workbooks.Open(src_path)

# Collect all Power Query names BEFORE refresh
query_names = [q.Name for q in wb.Queries]

print("Queries found in workbook:")
for q in query_names:
    print(" -", q)

# Optional: force synchronous refresh for OLEDB connections
for conn in wb.Connections:
    try:
        conn.OLEDBConnection.BackgroundQuery = False
    except Exception:
        pass

# Refresh all connections
wb.RefreshAll()

print("\nWaiting for refresh to complete...")
time.sleep(2)

def is_refreshing(wb):
    for conn in wb.Connections:
        # Check OLEDB connections
        try:
            if conn.OLEDBConnection.Refreshing:
                return True
        except Exception:
            pass

        # Check other connection types
        try:
            if conn.Refreshing:
                return True
        except Exception:
            pass

    return False

while is_refreshing(wb):
    time.sleep(1)


# Save only — do NOT close
wb.Save()

print("\nThe following queries were refreshed and saved:")
for q in query_names:
    print(" -", q)

print("\nExcel remains open. Close it manually when you're done.")
