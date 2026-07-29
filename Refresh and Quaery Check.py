import time
import win32com.client as win32

def wait_for_refresh(wb, check_interval=1):
    """Wait until all connections finish refreshing."""
    while True:
        refreshing = False

        for conn in wb.Connections:
            # OLEDB
            try:
                if conn.OLEDBConnection.Refreshing:
                    refreshing = True
                    break
            except Exception:
                pass

            # ODBC
            try:
                if conn.ODBCConnection.Refreshing:
                    refreshing = True
                    break
            except Exception:
                pass

        if not refreshing:
            break

        time.sleep(check_interval)


def check_refresh_errors(wb):
    """Return dict with success/error status for each connection."""
    results = {}

    for conn in wb.Connections:
        name = conn.Name
        status = None

        # OLEDB
        try:
            status = conn.OLEDBConnection.LastRefreshStatus
        except Exception:
            pass

        # ODBC
        try:
            status = conn.ODBCConnection.LastRefreshStatus
        except Exception:
            pass

        if status is None:
            results[name] = "Unknown (no status available)"
        elif status == 0:
            results[name] = "OK"
        else:
            results[name] = f"ERROR (status code {status})"

    return results


file_path = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets\Templete_newmonth.xlsx"

# Start Excel
excel = win32.Dispatch("Excel.Application")
excel.Visible = False  # Excel runs hidden; change to True if you want to see it

# Open workbook
wb = excel.Workbooks.Open(file_path)

# Collect Power Query names
query_names = [q.Name for q in wb.Queries]

print("Queries found in workbook:")
for q in query_names:
    print(" -", q)

# Force synchronous refresh for OLEDB
for conn in wb.Connections:
    try:
        conn.OLEDBConnection.BackgroundQuery = False
    except Exception:
        pass

# Refresh all
wb.RefreshAll()

print("\nWaiting for refresh to complete...")
wait_for_refresh(wb)

# Check for errors
print("\nChecking refresh results...")
results = check_refresh_errors(wb)

# Save and close
wb.Save()
wb.Close()

print("\nREFRESH SUMMARY:")
for conn_name, status in results.items():
    print(f" - {conn_name}: {status}")

print("\nWorkbook refreshed, saved, and closed.")
