import time
import win32com.client as win32

def wait_for_refresh(wb, check_interval=1):
    while True:
        refreshing = False

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

        if not refreshing:
            break

        time.sleep(check_interval)


def check_refresh_errors(wb):
    results = {}

    for conn in wb.Connections:
        name = conn.Name
        status = None

        try:
            status = conn.OLEDBConnection.LastRefreshStatus
        except:
            pass

        try:
            status = conn.ODBCConnection.LastRefreshStatus
        except:
            pass

        if status is None:
            results[name] = "Unknown (no status available)"
        elif status == 0:
            results[name] = "OK"
        else:
            results[name] = f"ERROR (status code {status})"

    return results


def refresh_file(path, excel):
    print(f"\nOpening workbook: {path}")
    wb = excel.Workbooks.Open(path)

    query_names = [q.Name for q in wb.Queries]
    print("Queries found:")
    for q in query_names:
        print(" -", q)

    for conn in wb.Connections:
        try:
            conn.OLEDBConnection.BackgroundQuery = False
        except:
            pass

    wb.RefreshAll()
    print("Waiting for refresh to complete...")
    wait_for_refresh(wb)

    print("Checking refresh results...")
    results = check_refresh_errors(wb)

    wb.Save()
    wb.Close()

    print("\nREFRESH SUMMARY:")
    for conn_name, status in results.items():
        print(f" - {conn_name}: {status}")

    print(f"Workbook '{path}' refreshed, saved, and closed.\n")


# -------------------------
# ADD ALL FILES HERE
# -------------------------

files_to_refresh = [
    r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets\Templete_newmonth.xlsx",
    r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\RollOut.xlsx",
    r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\STAFF.xlsx",
    r"\\virgo\Work\OpControl\2024\SomeOtherFolder\AnotherFile.xlsx"
]

# Start Excel once
excel = win32.Dispatch("Excel.Application")
excel.Visible = False

# Refresh all files
for file_path in files_to_refresh:
    refresh_file(file_path, excel)

excel.Quit()

print("\nAll workbooks refreshed, saved, and closed.")
