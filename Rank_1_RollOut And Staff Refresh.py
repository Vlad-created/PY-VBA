import time
import win32com.client as win32

# ---------------------------------------------------------
# Helper: Wait until Excel is responsive (avoids COM errors)
# ---------------------------------------------------------
def wait_until_excel_ready(excel, timeout=60):
    start = time.time()
    while True:
        try:
            _ = excel.Ready
            return
        except:
            time.sleep(0.5)
        if time.time() - start > timeout:
            return


# ---------------------------------------------------------
# Helper: Wait until all connections finish refreshing
# ---------------------------------------------------------
def wait_for_refresh(wb, check_interval=1):
    excel = wb.Application

    while True:
        wait_until_excel_ready(excel)

        refreshing = False

        try:
            for conn in wb.Connections:
                try:
                    if conn.OLEDBConnection.Refreshing:
                        refreshing = True
                        break
                except:
                    pass

                try:
                    if conn.ODBCConnection.Refreshing:
                        refreshing = True
                        break
                except:
                    pass

        except Exception:
            # Excel still busy → wait and retry
            time.sleep(check_interval)
            continue

        if not refreshing:
            break

        time.sleep(check_interval)


# ---------------------------------------------------------
# Main refresh function
# ---------------------------------------------------------
def refresh_workbook(path):
    print(f"\nOpening workbook: {path}")

    wb = excel.Workbooks.Open(path)

    # List queries
    query_names = [q.Name for q in wb.Queries]
    print("Queries found:")
    for q in query_names:
        print(" -", q)

    # Disable background refresh for stability
    for conn in wb.Connections:
        try:
            conn.OLEDBConnection.BackgroundQuery = False
        except:
            pass

    # Start refresh
    print("Starting refresh...")
    wb.RefreshAll()

    # Excel built‑in async wait
    excel.CalculateUntilAsyncQueriesDone()

    # Additional safety wait
    print("Waiting for refresh to complete...")
    wait_for_refresh(wb)

    # Save and close
    wb.Save()
    wb.Close()

    print("\nRefreshed and saved:")
    for q in query_names:
        print(" -", q)


# ---------------------------------------------------------
# Start Excel once
# ---------------------------------------------------------
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
