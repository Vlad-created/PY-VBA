import time
import win32com.client as win32

def refresh_workbook(path):
    print(f"\nOpening workbook: {path}")

    wb = excel.Workbooks.Open(path)

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

    # Trigger refresh
    wb.RefreshAll()

    print("Waiting for refresh to complete...")

    # Excel method that blocks until all async PQ refreshes finish
    excel.CalculateUntilAsyncQueriesDone()

    # Small delay to ensure model commits
    time.sleep(45)

    wb.Save()
    wb.Close()

    print("\nThe following queries were refreshed and saved:")
    for q in query_names:
        print(" -", q)


# Start Excel once
excel = win32.Dispatch("Excel.Application")
excel.Visible = False  # Run in background

# File paths
file_Append_Targets = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Append_Targets.xlsx"
file_Targets = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Targets shops march 2018 za nalivane.xlsx"
file_Targets_total_NBS = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Targets_shops_TOTAL_NBS.xlsx"

# Refresh all workbooks
refresh_workbook(file_Append_Targets)
refresh_workbook(file_Targets)
refresh_workbook(file_Targets_total_NBS)

# Quit Excel
excel.Quit()

print("\nAll workbooks refreshed, saved, and closed.")
