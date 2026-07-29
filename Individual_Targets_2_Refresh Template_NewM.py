import time
import win32com.client as win32

def wait_for_refresh(wb, check_interval=1):
    """Wait until all connections finish refreshing."""
    while True:
        refreshing = False

        for conn in wb.Connections:
            # OLEDB connections
            try:
                if conn.OLEDBConnection.Refreshing:
                    refreshing = True
                    break
            except Exception:
                pass

            # ODBC connections
            try:
                if conn.ODBCConnection.Refreshing:
                    refreshing = True
                    break
            except Exception:
                pass

        if not refreshing:
            break

        time.sleep(check_interval)


file_path = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets\Templete_newmonth.xlsx"

# Start Excel
excel = win32.Dispatch("Excel.Application")
excel.Visible = True  # Keep Excel visible

# Open workbook
wb = excel.Workbooks.Open(file_path)

# Collect all Power Query names BEFORE refresh
query_names = [q.Name for q in wb.Queries]

print("Queries found in workbook:")
for q in query_names:
    print(" -", q)

# Force synchronous refresh for OLEDB connections (optional but safer)
for conn in wb.Connections:
    try:
        conn.OLEDBConnection.BackgroundQuery = False
    except Exception:
        pass

# Refresh all
wb.RefreshAll()

# Wait until Excel finishes refreshing
print("\nWaiting for refresh to complete...")
wait_for_refresh(wb)

# Save only — keep Excel open
wb.Save()

print("\nThe following queries were refreshed and saved:")
for q in query_names:
    print(" -", q)

print("\nExcel remains open. Close it manually when you're done.")
