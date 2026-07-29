import win32com.client as win32
import os
import sys
import shutil
import time

# -----------------------------
# PART 1 — DOWNLOAD ATTACHMENT
# -----------------------------

SENDER_NAME = "Tsvetelina Kostadinova"
SUBJECT_MATCH = "Template бонуси А1 shops"

destination_path = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets\Тemplate_A1_shops.xlsx"
refresh_file_path = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets\Templete_newmonth.xlsx"

print("STEP 1: Checking Outlook for email...")

# Start Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access folder
try:
    inbox = outlook.GetDefaultFolder(6)
    target_folder = inbox.Folders["Individual Targets"]
except:
    print("ERROR: Folder 'Individual Targets' not found.")
    sys.exit(1)

messages = target_folder.Items
messages.Sort("ReceivedTime", True)

# Find email
target_email = None
for msg in messages:
    if msg.Class == 43:
        if SENDER_NAME.lower() in msg.SenderName.lower() and SUBJECT_MATCH.lower() in msg.Subject.lower():
            target_email = msg
            break

if not target_email:
    print("No matching email found.")
    sys.exit(0)

# Show email info
email_time = target_email.ReceivedTime
email_time_str = email_time.strftime("%d.%m.%Y %H:%M:%S")

print(f"Found email from {target_email.SenderName}")
print(f"Email received on: {email_time_str}")

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
    print("Attachment failed to save.")
    sys.exit(1)

# Check if destination is locked
try:
    os.rename(destination_path, destination_path)
except PermissionError:
    print("Destination file is OPEN or LOCKED. Close it and try again.")
    sys.exit(1)

# Overwrite safely
shutil.copyfile(temp_path, destination_path)
os.remove(temp_path)

print(f"SUCCESS: Attachment saved to:\n{destination_path}")

# -----------------------------
# PART 2 — REFRESH EXCEL FILE
# -----------------------------

print("\nSTEP 2: Opening Excel and refreshing queries...")

excel = win32.Dispatch("Excel.Application")
excel.Visible = True  # Show Excel window

# Open workbook
wb = excel.Workbooks.Open(refresh_file_path)

# Collect Power Query names BEFORE refresh
query_names = [q.Name for q in wb.Queries]

print("Queries found in workbook:")
for q in query_names:
    print(" -", q)

# Force synchronous refresh for OLEDB connections
for conn in wb.Connections:
    try:
        conn.OLEDBConnection.BackgroundQuery = False
    except:
        pass

# -----------------------------
# FUNCTION: Check if refreshing
# -----------------------------
def is_refreshing(workbook):
    for conn in workbook.Connections:
        try:
            if conn.Refreshing:
                return True
        except:
            pass
    return False

# Refresh all
wb.RefreshAll()

print("\nWaiting for refresh to complete...")

# Wait loop
while is_refreshing(wb):
    time.sleep(1)

# Save workbook
wb.Save()

print("\nThe following queries were refreshed and saved:")
for q in query_names:
    print(" -", q)

print("\nDONE! Excel remains open — close it manually when finished.")
