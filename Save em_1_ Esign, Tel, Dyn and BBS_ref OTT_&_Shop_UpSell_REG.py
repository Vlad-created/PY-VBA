import win32com.client as win32
import os
import shutil
import time
import traceback
from datetime import datetime
import psutil

# -------------------------
# Configuration
# -------------------------
SENDER_NAME = "SQL Replica"          # default sender for TASKS
OUTLOOK_FOLDER = "SQL Replica"       # default subfolder inside Inbox for TASKS

TASKS = [
    {"subject": "ESignConsent",
     "destinations": [r"\\hope\Shop support\ESignConsent_New.txt"]},
    {"subject": "Telesales",
     "destinations": [r"\\hope\Shop support\telesales.txt"]},
    {"subject": "DynLeasing",
     "destinations": [r"\\hope\Shop support\DynLeasing_SQLReplica.txt"]},
    {"subject": "Client_Addressess",
     "destinations": [r"\\hope\Shop support\bbs_address_short.txt"]},

    # NEW
    {"subject": "Regular report reg_user_shop_info is done. Details in mail.",
     "sender": "Revenue BZ",
     "folder": None,                 # None = Inbox itself; set a subfolder name if the mail is filed elsewhere
     "extension": ".xlsx",           # pick the attachment with this extension
     "destinations": [
         r"\\virgo.mobiltel.bg\Work\OpControl\2016\deviceGames\reg_user_shop_info.xlsx",
         r"\\virgo.mobiltel.bg\Work\OpControl\2016\SALES\Reff_Tables\reg_user_shop_info\reg_user_shop_info.xlsx",
     ]},
]

SENDER_NAME2 = "l2automation@amdocs.com"
OUTLOOK_FOLDER2 = "UpSell_AI"

EXCEL_FILE = r"\\hope\Shop support\smart_app_ott_watched.xlsx"
EXCEL_FILE2 = r"\\hope\Shop support\shops_structure.xlsx"
EXCEL_FILE3 = r"\\hope\Shop support\AIupsellAdvisorData.xlsx"

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

def wait_for_excel_shutdown():
    for _ in range(50):  # up to 5 seconds
        if not any(p.name().lower() == "excel.exe" for p in psutil.process_iter()):
            return
        time.sleep(0.1)

def file_mod_time(path):
    try:
        ts = os.path.getmtime(path)
        return datetime.fromtimestamp(ts)
    except Exception:
        return None

def save_attachment_and_overwrite(attachment, destination_paths):
    """Save attachment once to TEMP, then copy it to every destination."""
    if isinstance(destination_paths, str):
        destination_paths = [destination_paths]

    temp = os.path.join(os.environ.get("TEMP", "."), attachment.FileName)
    attachment.SaveAsFile(temp)
    try:
        for dest in destination_paths:
            try:
                dest_dir = os.path.dirname(dest)
                if dest_dir and not os.path.exists(dest_dir):
                    os.makedirs(dest_dir, exist_ok=True)
                shutil.copyfile(temp, dest)
                log(f"  Saved RAW file to: {dest}")
                log(f"  Destination file modified time: {file_mod_time(dest)}")
            except Exception as e:
                log(f"  Error saving to {dest}: {e}")
                traceback.print_exc()
    finally:
        try:
            os.remove(temp)
        except Exception:
            pass

def ensure_dir(path):
    try:
        os.makedirs(path, exist_ok=True)
        return True
    except Exception as e:
        log(f"  Could not create directory {path}: {e}")
        return False

def save_attachment_to_folder_if_not_exists(attachment, dest_folder):
    if not os.path.exists(dest_folder):
        if not ensure_dir(dest_folder):
            return False, None

    filename = attachment.FileName
    dest_path = os.path.join(dest_folder, filename)

    if os.path.exists(dest_path):
        log(f"  Destination file already exists, skipping save: {dest_path}")
        return False, dest_path

    temp = os.path.join(os.environ.get("TEMP", "."), filename)
    attachment.SaveAsFile(temp)
    shutil.copyfile(temp, dest_path)
    try:
        os.remove(temp)
    except Exception:
        pass

    return True, dest_path

# -------------------------
# Outlook processing (TASKS)
# -------------------------
def process_emails():
    log("Starting email processing")
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)

    folder_cache = {}  # folder name (or None) -> sorted Items

    def get_messages(folder_name):
        if folder_name in folder_cache:
            return folder_cache[folder_name]
        try:
            folder = inbox if folder_name is None else inbox.Folders[folder_name]
        except Exception:
            log(f"Folder '{folder_name}' not found inside your Inbox.")
            folder_cache[folder_name] = None
            return None
        items = folder.Items
        items.Sort("ReceivedTime", True)
        folder_cache[folder_name] = items
        return items

    def process_task(task):
        subject_match = task["subject"]
        sender_match = task.get("sender", SENDER_NAME)
        folder_name = task["folder"] if "folder" in task else OUTLOOK_FOLDER
        extension = (task.get("extension") or "").lower()
        destinations = task["destinations"]

        log(f"Processing: {subject_match}")
        messages = get_messages(folder_name)
        if messages is None:
            return

        target_email = None
        for msg in messages:
            try:
                if msg.Class == 43:
                    sender = (msg.SenderName or "").strip().lower()
                    subject = (msg.Subject or "").strip().lower()
                    if sender_match.lower() in sender and subject_match.lower() in subject:
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
            count = target_email.Attachments.Count
            if count == 0:
                log("  Email has no attachments.")
                return

            attachment = None
            if extension:
                for i in range(1, count + 1):
                    att = target_email.Attachments.Item(i)
                    if att.FileName.lower().endswith(extension):
                        attachment = att
                        break
                if attachment is None:
                    log(f"  No '{extension}' attachment found.")
                    return
            else:
                attachment = target_email.Attachments.Item(1)

            log(f"  Attachment: {attachment.FileName}")
            save_attachment_and_overwrite(attachment, destinations)
        except Exception as e:
            log(f"  Error saving attachment: {e}")
            traceback.print_exc()

    for task in TASKS:
        process_task(task)

    log("All email tasks completed.")

# -------------------------
# UpSell_AI specific processing
# -------------------------
def process_upsell_ai_email():
    log("Processing UpSell_AI folder for latest A1 Sales Advisor Agent | KPI email")
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    try:
        inbox = outlook.GetDefaultFolder(6)
        target_folder = inbox.Folders[OUTLOOK_FOLDER2]
    except Exception:
        log(f"Subfolder '{OUTLOOK_FOLDER2}' not found inside your Inbox.")
        return

    messages = target_folder.Items
    messages.Sort("ReceivedTime", True)

    subject_match = "A1 Sales Advisor Agent | KPI"
    target_email = None

    for msg in messages:
        try:
            if msg.Class != 43:
                continue
            sender = (msg.SenderName or "").strip().lower()
            sender_full = (getattr(msg, "SenderEmailAddress", "") or "").strip().lower()
            subject = (msg.Subject or "").strip().lower()

            if subject_match.lower() in subject and (
                SENDER_NAME2.lower() in sender or SENDER_NAME2.lower() in sender_full
            ):
                target_email = msg
                break
        except Exception:
            continue

    if not target_email:
        log("No matching email found in UpSell_AI.")
        return

    log(f"  Found email from: {target_email.SenderName}")
    log(f"  Subject: {target_email.Subject}")
    log(f"  Received: {target_email.ReceivedTime}")

    try:
        if target_email.Attachments.Count == 0:
            log("  Email has no attachments.")
            return

        # Allowed extensions
        allowed_ext = {".xlsm", ".xlsx", ".xls", ".txt"}

        attachment = target_email.Attachments.Item(1)
        filename = attachment.FileName.lower()

        # Check extension
        if not any(filename.endswith(ext) for ext in allowed_ext):
            log(f"  Attachment skipped (unsupported file type): {filename}")
            return

        dest_folder = r"\\hope\Shop support\AIupsellAdvisor"

        saved, path = save_attachment_to_folder_if_not_exists(attachment, dest_folder)
        if saved:
            log(f"  Attachment saved to: {path}")
        else:
            log(f"  Attachment skipped or not saved: {path}")

    except Exception as e:
        log(f"  Error processing attachments: {e}")
        traceback.print_exc()


# -------------------------
# Excel refresh 1 (EXCEL_FILE)
# -------------------------
def refresh_excel_workbook():
    wait_for_excel_shutdown()
    log("Opening Excel to refresh Power Query connections (smart_app_ott_watched)...")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = DEBUG_VISIBLE

    xlCalculationAutomatic = -4105
    xlDone = 0

    wb = None
    prev_calc = None
    try:
        wb = excel.Workbooks.Open(EXCEL_FILE)
        log(f"Workbook opened: {EXCEL_FILE}")

        for conn in wb.Connections:
            try:
                if hasattr(conn, "OLEDBConnection"):
                    conn.OLEDBConnection.BackgroundQuery = False
                if hasattr(conn, "ODBCConnection"):
                    conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass

        prev_calc = excel.Calculation
        excel.Calculation = xlCalculationAutomatic

        wb.RefreshAll()

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
        while any_refreshing(wb, excel):
            if time.time() - start > REFRESH_TIMEOUT_SECONDS:
                log("Warning: refresh timeout reached.")
                break
            time.sleep(POLL_INTERVAL)

        wb.Save()
        log("Workbook saved.")
        log("Waiting 5 seconds before continuing...")
        time.sleep(5)


    except Exception as e:
        log(f"Error during Excel refresh: {e}")
        traceback.print_exc()
    finally:
        try:
            if wb:
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

    log("Excel refresh process completed for smart_app_ott_watched.")

# -------------------------
# Excel refresh 2 (EXCEL_FILE2)
# -------------------------
def refresh_excel_workbook2():
    wait_for_excel_shutdown()
    log("Opening Excel to refresh Power Query connections (shops_structure)...")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = DEBUG_VISIBLE

    xlCalculationAutomatic = -4105
    xlDone = 0

    wb = None
    prev_calc = None

    try:
        wb = excel.Workbooks.Open(EXCEL_FILE2)
        log(f"Workbook opened: {EXCEL_FILE2}")

        for conn in wb.Connections:
            try:
                if hasattr(conn, "OLEDBConnection"):
                    conn.OLEDBConnection.BackgroundQuery = False
                if hasattr(conn, "ODBCConnection"):
                    conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass

        prev_calc = excel.Calculation
        excel.Calculation = xlCalculationAutomatic

        wb.RefreshAll()

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
        while any_refreshing2(wb, excel):
            if time.time() - start > REFRESH_TIMEOUT_SECONDS:
                log("Warning: refresh timeout reached.")
                break
            time.sleep(POLL_INTERVAL)

        wb.Save()
        log("Workbook saved.")
        log("Waiting 5 seconds before continuing...")
        time.sleep(5)


    except Exception as e:
        log(f"Error during Excel refresh: {e}")
        traceback.print_exc()
    finally:
        try:
            if wb:
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

    log("Excel refresh process completed for shops_structure.")

# -------------------------
# Excel refresh 3 (EXCEL_FILE3)
# -------------------------
def refresh_excel_workbook3():
    wait_for_excel_shutdown()
    log("Opening Excel to refresh ALL queries (AIupsellAdvisorData)...")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = DEBUG_VISIBLE

    xlDone = 0
    wb = None

    try:
        wb = excel.Workbooks.Open(EXCEL_FILE3)
        log(f"Workbook opened: {EXCEL_FILE3}")

        for conn in wb.Connections:
            try:
                if hasattr(conn, "OLEDBConnection"):
                    conn.OLEDBConnection.BackgroundQuery = False
                if hasattr(conn, "ODBCConnection"):
                    conn.ODBCConnection.BackgroundQuery = False
            except Exception:
                pass

        wb.RefreshAll()

        def any_refreshing3(workbook, app):
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
        while any_refreshing3(wb, excel):
            if time.time() - start > REFRESH_TIMEOUT_SECONDS:
                log("Warning: refresh timeout reached.")
                break
            time.sleep(POLL_INTERVAL)

        wb.Save()
        log("Workbook saved.")
        log("Waiting 5 seconds before continuing...")
        time.sleep(5)


    except Exception as e:
        log(f"Error during Excel refresh: {e}")
        traceback.print_exc()
    finally:
        try:
            if wb:
                wb.Close(False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass

    log("Excel refresh process completed for AIupsellAdvisorData.")

# -------------------------
# Main
# -------------------------
if __name__ == "__main__":
    start_time = time.time()
    log("Script started")
    try:
        process_emails()
        process_upsell_ai_email()
        refresh_excel_workbook()
        refresh_excel_workbook2()
        refresh_excel_workbook3()
    except Exception as e:
        log(f"Unhandled exception: {e}")
        traceback.print_exc()
    finally:
        elapsed = time.time() - start_time
        log(f"Script finished in {elapsed:.1f} seconds")