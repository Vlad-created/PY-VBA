import os
import time
import subprocess
import traceback
from datetime import datetime
import shutil

# COM imports
import win32com.client as win32
import win32com.client as win32com  # fixed missing import used for WMI
# optional GUI bring-to-front
try:
    import win32gui, win32con
    HAS_PYWIN32 = True
except Exception:
    HAS_PYWIN32 = False

# Configuration
EXCEL_FILE = r"\\hope\Shop support\smart_app_ott_watched.xlsx"
DEBUG_VISIBLE = True
REFRESH_TIMEOUT_SECONDS = 1200
POLL_INTERVAL = 1.0
KILL_EXISTING_EXCEL = True

# Helpers
def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log(msg):
    print(f"[{now()}] {msg}")

def list_excel_processes_wmi():
    try:
        locator = win32com.Dispatch("WbemScripting.SWbemLocator")
        svc = locator.ConnectServer(".", "root\\cimv2")
        procs = svc.ExecQuery("Select ProcessId, Name from Win32_Process where Name='EXCEL.EXE'")
        rows = []
        for p in procs:
            try:
                owner = p.GetOwner()
                owner_user = f"{owner.Domain}\\{owner.User}"
            except Exception:
                owner_user = "<unknown>"
            rows.append((int(p.ProcessId), owner_user))
        if not rows:
            log("No running EXCEL.EXE processes found.")
        else:
            for pid, owner in rows:
                log(f"EXCEL process Id={pid} Owner={owner}")
        return rows
    except Exception as e:
        log(f"Error listing Excel processes via WMI: {e}")
        traceback.print_exc()
        return []

def kill_excel_processes_wmi():
    try:
        locator = win32com.Dispatch("WbemScripting.SWbemLocator")
        svc = locator.ConnectServer(".", "root\\cimv2")
        procs = svc.ExecQuery("Select ProcessId from Win32_Process where Name='EXCEL.EXE'")
        for p in procs:
            try:
                pid = int(p.ProcessId)
                log(f"Terminating EXCEL process Id={pid}")
                p.Terminate()
            except Exception:
                pass
        time.sleep(1)
    except Exception as e:
        log(f"Error terminating Excel processes: {e}")
        traceback.print_exc()

def try_bring_signin_to_front():
    if not HAS_PYWIN32:
        log("pywin32 not installed; cannot attempt to bring sign-in window to front.")
        return False
    SIGNIN_TITLES = [
        "Sign in to your account",
        "Microsoft - Pick an account",
        "Pick an account",
        "Sign in"
    ]
    found = []
    def enum_cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if title:
            for t in SIGNIN_TITLES:
                if t.lower() in title.lower():
                    found.append((hwnd, title))
                    break
    win32gui.EnumWindows(enum_cb, None)
    if found:
        for hwnd, title in found:
            log(f"Found sign-in window: {title}")
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
        return True
    else:
        log("No sign-in window found on desktop (it may be on another session).")
        return False

# Main routine
def refresh_excel_workbook():
    log("Starting Excel refresh routine")
    # session info
    sess = os.environ.get("SESSIONNAME")
    user = os.environ.get("USERNAME") or None
    log(f"ENV SESSIONNAME: {sess}")
    log(f"ENV USERNAME: {user}")
    try:
        whoami = subprocess.check_output(["whoami"], shell=True, text=True).strip()
        log(f"whoami: {whoami}")
    except Exception:
        pass

    # list Excel processes
    rows = list_excel_processes_wmi()
    if rows and KILL_EXISTING_EXCEL:
        log("KILL_EXISTING_EXCEL is True — terminating existing Excel processes to ensure interactive session")
        kill_excel_processes_wmi()
        time.sleep(1)
        list_excel_processes_wmi()

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = DEBUG_VISIBLE

    wb = None
    prev_calc = None
    try:
        wb = excel.Workbooks.Open(EXCEL_FILE)
        log(f"Workbook opened: {EXCEL_FILE}")

        # detect pbiazure connections
        has_pbiazure = False
        try:
            for conn in wb.Connections:
                try:
                    if hasattr(conn, "OLEDBConnection"):
                        cs = str(conn.OLEDBConnection.Connection or "").lower()
                        if "pbiazure" in cs or "powerbi" in cs or "claimstoken" in cs:
                            has_pbiazure = True
                            break
                except Exception:
                    pass
        except Exception:
            pass

        if has_pbiazure:
            log("Detected Power BI / pbiazure connection(s). Attempting to surface sign-in dialog.")
            surfaced = try_bring_signin_to_front()
            if not surfaced:
                log("Sign-in window not found on this desktop. Please ensure you are running the script interactively.")
            try:
                input("If you see the Excel sign-in dialog, sign in now. Press Enter here after sign-in to continue...")
            except Exception:
                log("Interactive input not available; pausing 30 seconds to allow manual sign-in.")
                time.sleep(30)

        # disable background refresh where possible
        try:
            for conn in wb.Connections:
                try:
                    if hasattr(conn, "OLEDBConnection"):
                        conn.OLEDBConnection.BackgroundQuery = False
                    if hasattr(conn, "ODBCConnection"):
                        conn.ODBCConnection.BackgroundQuery = False
                except Exception:
                    pass
        except Exception:
            pass

        try:
            prev_calc = excel.Calculation
            excel.Calculation = -4105  # xlCalculationAutomatic
        except Exception:
            prev_calc = None

        # trigger refresh
        try:
            wb.RefreshAll()
        except Exception as e:
            log(f"wb.RefreshAll() raised: {e}")

        # best-effort wait
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        def any_refreshing(workbook, app):
            xlDone = 0
            try:
                for conn in workbook.Connections:
                    try:
                        if hasattr(conn, "Refreshing") and conn.Refreshing:
                            return True
                    except Exception:
                        pass
                    try:
                        if hasattr(conn, "OLEDBConnection") and getattr(conn.OLEDBConnection, "Refreshing", False):
                            return True
                    except Exception:
                        pass
                for ws in workbook.Worksheets:
                    try:
                        for qt in ws.QueryTables:
                            try:
                                if getattr(qt, "Refreshing", False):
                                    return True
                            except Exception:
                                pass
                    except Exception:
                        pass
                try:
                    if app.CalculationState != xlDone:
                        return True
                except Exception:
                    pass
            except Exception:
                pass
            return False

        start = time.time()
        last_log = start
        while any_refreshing(wb, excel):
            elapsed = time.time() - start
            if elapsed > REFRESH_TIMEOUT_SECONDS:
                log("Warning: refresh timeout reached.")
                break
            if time.time() - last_log > 10:
                log(f"Waiting for refresh to complete... elapsed {int(elapsed)}s")
                last_log = time.time()
            time.sleep(POLL_INTERVAL)

        log("Refresh finished or timed out; saving workbook.")
        wb.Save()
        log("Workbook saved.")
    except Exception as e:
        log(f"Error during Excel refresh: {e}")
        traceback.print_exc()
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if prev_calc is not None:
                excel.Calculation = prev_calc
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
    log("Excel refresh routine completed")

if __name__ == "__main__":
    refresh_excel_workbook()
