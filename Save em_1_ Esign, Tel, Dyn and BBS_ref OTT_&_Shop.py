import win32com.client as win32
import os
import shutil
import time
import traceback
from datetime import datetime

# -------------------------
# User's Edge browser tabs metadata (context only)
edge_all_open_tabs = [
    {"pageTitle":"<WebsiteContent_GZQRUr5sCbNdd66FRVn2W></WebsiteContent_GZQRUr5sCbNdd66FRVn2W>",
     "pageUrl":"<WebsiteContent_GZQRUr5sCbNdd66FRVn2W></WebsiteContent_GZQRUr5sCbNdd66FRVn2W>",
     "tabId":-1,"isCurrent":True}
]

# -------------------------
# Configuration
# -------------------------
TASKS = [
    {"subject": "ESignConsent", "destination": r"\\hope\Shop support\ESignConsent_New.txt"},
    {"subject": "Telesales", "destination": r"\\hope\Shop support\telesales.txt"},
    {"subject": "DynLeasing", "destination": r"\\hope\Shop support\DynLeasing_SQLReplica.txt"},
    {"subject": "Client_Addressess", "destination": r"\\hope\Shop support\bbs_address_short.txt"}
]

SENDER_NAME = "SQL Replica"
OUTLOOK_FOLDER = "SQL Replica"
EXCEL_FILE = r"\\hope\Shop support\smart_app_ott_watched.xlsx"
EXCEL_FILE2 = r"\\hope\Shop support\shops_structure.xlsx"

DEBUG_VISIBLE = True
REFRESH_TIMEOUT_SECONDS = 1200
POLL_INTERVAL = 1.0

# -------------------------
# Helpers
# -------------------------
def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log(msg):
    print(f"[{now()}] {msg}")

def file_mod_time(path):
    try:
        ts = os.path.getmtime(path)
        return datetime.fromtimestamp(ts)
    except Exception:
        return None

def save_attachment_and_overwrite(attachment, destination_path):
    temp = os.path.join(os.environ.get("TEMP", "."), attachment.FileName)
    attachment.SaveAsFile(temp)
    dest_dir = os.path.dirname(destination_path)
    if dest_dir and not os.path.exists(dest_dir):
        os.makedirs(dest_dir, exist_ok=True)
    shutil.copyfile(temp, destination_path)
    try:
        os.remove(temp)
    except Exception:
        pass

# -------------------------
# Outlook processing
# -------------------------
def process_emails():
    log("Starting email processing")
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    try:
        inbox = outlook.GetDefaultFolder(6)
        target_folder = inbox.Folders[OUTLOOK_FOLDER]
    except Exception:
        log(f"Subfolder '{OUTLOOK_FOLDER}' not found inside your Inbox.")
        return

    messages = target_folder.Items
    messages.Sort("ReceivedTime", True)

    def process_task(subject_match, destination_path):
        log(f"Processing: {subject_match}")
        target_email = None
        for msg in messages:
            try:
                if msg.Class == 43:
                    sender = (msg.SenderName or "").strip().lower()
                    subject = (msg.Subject or "").strip().lower()
                    if SENDER_NAME.lower() in sender and subject_match.lower() in subject:
                        target_email = msg
                        break
            except Exception:
                continue

        if not target_email:
            log(f"No email found for subject: {subject_match}")
            return

        log(f"  Found email from: {target_email.SenderName}")
        log(f"  Subject: {target_email.Subject}")
        log(f"  Received: {target_email.ReceivedTime}")

        try:
            if target_email.Attachments.Count == 0:
                log("  Email has no attachments.")
                return

            attachment = target_email.Attachments.Item(1)
            save_attachment_and_overwrite(attachment, destination_path)
            log(f"  Saved RAW file to: {destination_path}")
            try:
                mt = file_mod_time(destination_path)
                log(f"  Destination file modified time: {mt}")
            except Exception:
                pass
        except Exception as e:
            log(f"  Error saving attachment: {e}")
            traceback.print_exc()

    for task in TASKS:
        process_task(task["subject"], task["destination"])

    log("All email tasks completed.")

# -------------------------
# Excel refresh 1
# -------------------------
def refresh_excel_workbook():
    log("Opening Excel to refresh Power Query connections...")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = DEBUG_VISIBLE

    xlCalculationAutomatic = -4105
    xlDone = 0

    wb = None
    prev_calc = None
    try:
        wb = excel.Workbooks.Open(EXCEL_FILE)
        log(f"Workbook opened: {EXCEL_FILE}")

        # Diagnostics
        try:
            qnames = [q.Name for q in wb.Queries]
        except Exception:
            qnames = []
        log("Queries in workbook:")
        for q in qnames:
            log(f" - {q}")

        for idx, conn in enumerate(wb.Connections, start=1):
            try:
                name = getattr(conn, "Name", "<no name>")
                log(f"Connection #{idx}: {name}")
                if hasattr(conn, "OLEDBConnection"):
                    cs = conn.OLEDBConnection.Connection
                    log(f"  OLEDBConnection snippet: {str(cs)[:300]}")
            except Exception:
                pass

        # Detect Power BI
        has_pbiazure = False
        for conn in wb.Connections:
            try:
                if hasattr(conn, "OLEDBConnection"):
                    cs = str(conn.OLEDBConnection.Connection or "").lower()
                    if "pbiazure" in cs or "powerbi" in cs or "claimstoken" in cs:
                        has_pbiazure = True
                        break
            except Exception:
                pass

        if has_pbiazure and DEBUG_VISIBLE:
            log("Detected Power BI / pbiazure connection(s). An interactive sign-in may be required.")
            log("Please sign in using the Excel sign-in dialog that appears, then press Enter here to continue.")
            try:
                input("Press Enter after you have completed sign-in in Excel...")
            except Exception:
                log("Interactive input not available; pausing for 30 seconds.")
                time.sleep(30)

        # Disable background refresh
        for conn in wb.Connections:
            try:
                if hasattr(conn, "OLEDBConnection"):
                    conn.OLEDBConnection.BackgroundQuery = False
                if hasattr(conn, "ODBCConnection"):
                    conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass

        # Calculation mode
        try:
            prev_calc = excel.Calculation
            excel.Calculation = xlCalculationAutomatic
        except Exception:
            prev_calc = None

        # Per-connection refresh
        log("Attempting per-connection Refresh() calls (best-effort)...")
        for conn in wb.Connections:
            try:
                name = getattr(conn, "Name", "<unknown>")
                log(f"  Calling Refresh() on connection: {name}")
                try:
                    conn.Refresh()
                except Exception:
                    pass
            except Exception:
                traceback.print_exc()

        # RefreshAll
        log("Calling wb.RefreshAll()")
        try:
            wb.RefreshAll()
        except Exception as e:
            log(f"wb.RefreshAll() raised: {e}")

        # Wait for refresh
        def any_refreshing(workbook, app):
            for conn in workbook.Connections:
                try:
                    if getattr(conn, "Refreshing", False):
                        return True
                except Exception:
                    pass
            try:
                if app.CalculationState != xlDone:
                    return True
            except Exception:
                pass
            return False

        start = time.time()
        last_heartbeat = start
        while any_refreshing(wb, excel):
            elapsed = time.time() - start
            if elapsed > REFRESH_TIMEOUT_SECONDS:
                log("Warning: refresh timeout reached.")
                break
            if time.time() - last_heartbeat > 10:
                log(f"Waiting for refresh to complete... elapsed {int(elapsed)}s")
                last_heartbeat = time.time()
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

    log("Excel refresh process completed (open → refresh → save → close).")

# -------------------------
# Excel refresh 2 (HARDENED)
# -------------------------
def refresh_excel_workbook2():
    log("Waiting 5 seconds before opening second workbook...")
    time.sleep(5)

    log("Opening Excel to refresh Power Query connections...")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = DEBUG_VISIBLE
    excel.AutomationSecurity = 3  # Disable external data loading

    xlCalculationAutomatic = -4105
    xlDone = 0

    wb = None
    prev_calc = None

    try:
        # Retry logic for opening workbook
        for attempt in range(3):
            try:
                wb = excel.Workbooks.Open(EXCEL_FILE2)
                break
            except Exception as e:
                log(f"Open attempt {attempt+1} failed: {e}")
                time.sleep(2)

        if wb is None:
            log("Failed to open workbook after 3 attempts.")
            return

        log(f"Workbook opened: {EXCEL_FILE2}")

        # Protect against COM disconnect
        try:
            connections = wb.Connections
        except Exception:
            log("Workbook COM object disconnected while accessing Connections.")
            wb.Close(False)
            excel.Quit()
            return

        # Diagnostics
        try:
            qnames = [q.Name for q in wb.Queries]
        except Exception:
            qnames = []
        log("Queries in workbook:")
        for q in qnames:
            log(f" - {q}")

        # List connections safely
        try:
            for idx, conn in enumerate(connections, start=1):
                try:
                    name = getattr(conn, "Name", "<no name>")
                    log(f"Connection #{idx}: {name}")
                except Exception:
                    pass
        except Exception:
            log("Error reading connections (COM disconnect).")

        # Disable background refresh
        for conn in connections:
            try:
                if hasattr(conn, "OLEDBConnection"):
                    conn.OLEDBConnection.BackgroundQuery = False
                if hasattr(conn, "ODBCConnection"):
                    conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass

        # Calculation mode
        try:
            prev_calc = excel.Calculation
            excel.Calculation = xlCalculationAutomatic
        except Exception:
            prev_calc = None

        # Per-connection refresh
        log("Attempting per-connection Refresh() calls (best-effort)...")
        for conn in connections:
            try:
                name = getattr(conn, "Name", "<unknown>")
                log(f"  Calling Refresh() on connection: {name}")
                try:
                    conn.Refresh()
                except Exception:
                    pass
            except Exception:
                traceback.print_exc()

        # RefreshAll
        log("Calling wb.RefreshAll()")
        try:
            wb.RefreshAll()
        except Exception as e:
            log(f"wb.RefreshAll() raised: {e}")

        # Wait for refresh
        def any_refreshing2(workbook, app):
            for conn in workbook.Connections:
                try:
                    if getattr(conn, "Refreshing", False):
                        return True
                except Exception:
                    pass
            try:
                if app.CalculationState != xlDone:
                    return True
            except Exception:
                pass
            return False

        start = time.time()
        last_heartbeat = start
        while any_refreshing2(wb, excel):
            elapsed = time.time() - start
            if elapsed > REFRESH_TIMEOUT_SECONDS:
                log("Warning: refresh timeout reached.")
                break
            if time.time() - last_heartbeat > 10:
                log(f"Waiting for refresh to complete... elapsed {int(elapsed)}s")
                last_heartbeat = time.time()
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

    log("Excel refresh process completed (open → refresh → save → close).")

# -------------------------
# Main
# -------------------------
if __name__ == "__main__":
    start_time = time.time()
    log("Script started")
    try:
        process_emails()
        refresh_excel_workbook()
        refresh_excel_workbook2()
    except Exception as e:
        log(f"Unhandled exception: {e}")
        traceback.print_exc()
    finally:
        elapsed = time.time() - start_time
        log(f"Script finished in {elapsed:.1f} seconds")
