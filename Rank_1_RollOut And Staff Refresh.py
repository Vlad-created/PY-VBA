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
    time.sleep(30)

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
file_rollout = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\RollOut.xlsx"
file_staff   = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\STAFF.xlsx"

# Refresh both workbooks
refresh_workbook(file_rollout)
refresh_workbook(file_staff)

# Quit Excel
excel.Quit()

print("\nAll workbooks refreshed, saved, and closed.")
