import win32com.client as win32
import os
import sys
import shutil
import time

# -----------------------------
# CONFIGURATION
# -----------------------------

SENDER_NAME = "Rumen I. Dimitrov"
SUBJECT_MATCH = "Дигитално разписани бланки"

destination_path = r"\\virgo\Work\OpControl\2016\Bogdan\Prima Sales Registration\Дигитално разписани - получени файлове\TTl registrations - to upload.xlsx"
refresh_file_path = r"\\virgo\Work\OpControl\2016\Bogdan\Prima Sales Registration\TTl registrations (002).xlsx"

print("STEP 1: Checking Outlook for email...")

# -----------------------------
# PART 1 — DOWNLOAD ATTACHMENT
# -----------------------------

# Start Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Correct Inbox access
try:
    target_folder = outlook.GetDefaultFolder(6)  # 6 = Inbox
except Exception as e:
    print("ERROR: Cannot access Outlook Inbox:", e)
    sys.exit(1)

# Restrict to normal mail items only
messages = target_folder.Items
messages = messages.Restrict("[MessageClass]='IPM.Note'")
messages.Sort("ReceivedTime", True)

# Find matching email
target_email = None
for msg in messages:
    try:
        if SENDER_NAME.lower() in msg.SenderName.lower() and SUBJECT_MATCH.lower() in msg.Subject.lower():
            target_email = msg
            break
    except:
        continue

if not target_email:
    print("No matching email found.")
    sys.exit(0)

print(f"Found email from: {target_email.SenderName}")
print(f"Subject: {target_email.Subject}")
print(f"Received at: {target_email.ReceivedTime}")

# Find Excel attachment
excel_attachment = None
for att in target_email.Attachments:
    if att.FileName.lower().endswith((".xlsx", ".xlsm", ".xls")):
        excel_attachment = att
        break

if not excel_attachment:
    print("No Excel attachment found.")
    sys.exit(0)

# Save to temp
temp_path = os.path.join(os.environ["TEMP"], "temp_attachment.xlsx")
excel_attachment.SaveAsFile(temp_path)

if not os.path.exists(temp_path):
    print("ERROR: Attachment failed to save.")
    sys.exit(1)

print("Attachment downloaded successfully.")

# Check if destination file is locked
try:
    with open(destination_path, "a"):
        pass
except PermissionError:
    print("Destination file is OPEN or LOCKED. Close it and try again.")
    sys.exit(1)

# Overwrite destination file
try:
    shutil.copyfile(temp_path, destination_path)
    os.remove(temp_path)
except Exception as e:
    print("ERROR copying file:", e)
    sys.exit(1)

print(f"SUCCESS: Attachment saved to:\n{destination_path}")

# -----------------------------
# PART 2 — REFRESH EXCEL FILE
# -----------------------------

print("\nSTEP 2: Opening Excel and refreshing queries...")

excel = win32.Dispatch("Excel.Application")
excel.Visible = True

try:
    wb = excel.Workbooks.Open(refresh_file_path)
except Exception as e:
    print("ERROR opening Excel file:", e)
    sys.exit(1)

# List queries
query_names = [q.Name for q in wb.Queries]

print("Queries found in workbook:")
for q in query_names:
    print(" -", q)

# Disable background refresh
for conn in wb.Connections:
    try:
        conn.OLEDBConnection.BackgroundQuery = False
    except:
        pass

# Refresh all queries
wb.RefreshAll()

print("\nWaiting for refresh to complete...")

def is_refreshing(wb):
    for conn in wb.Connections:
        try:
            if conn.Refreshing:
                return True
        except AttributeError:
            pass
    return False

while is_refreshing(wb):
    time.sleep(1)

# Save workbook
wb.Save()

print("\nThe following queries were refreshed and saved:")
for q in query_names:
    print(" -", q)

print("\nDONE! Excel remains open — close it manually when finished.")
