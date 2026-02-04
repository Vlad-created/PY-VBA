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

    # Refresh all connections
    wb.RefreshAll()

    # Wait for refresh to complete
    time.sleep(59)

    # Save and close
    wb.Save()
    wb.Close()

    print("\nThe following queries were refreshed and saved:")
    for q in query_names:
        print(" -", q)


# Start Excel once
excel = win32.Dispatch("Excel.Application")
excel.Visible = False

# File paths
file_Append_Targets = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Append_Targets.xlsx"
file_Targets   = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Targets shops march 2018 za nalivane.xlsx"
file_Targets_total_NBS = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Targets_shops_TOTAL_NBS.xlsx"

# Refresh both workbooks
refresh_workbook(file_Append_Targets)
refresh_workbook(file_Targets)
refresh_workbook(file_Targets_total_NBS)

# Quit Excel
excel.Quit()

print("\nAll workbooks refreshed, saved, and closed.")
