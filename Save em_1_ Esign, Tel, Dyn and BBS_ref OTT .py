import win32com.client as win32
import os
import shutil
import time

TASKS = [
    {
        "subject": "ESignConsent",
        "destination": r"\\hope\Shop support\ESignConsent_New.txt"
    },
    {
        "subject": "Telesales",
        "destination": r"\\hope\Shop support\telesales.txt"
    },
    {
        "subject": "DynLeasing",
        "destination": r"\\hope\Shop support\DynLeasing_SQLReplica.txt"
    },
    {
        "subject": "Client_Addressess",
        "destination": r"\\hope\Shop support\bbs_address_short.txt"
    }
]

SENDER_NAME = "SQL Replica"
OUTLOOK_FOLDER = "SQL Replica"

EXCEL_FILE = r"\\hope\Shop support\smart_app_ott_watched.xlsx"


outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

try:
    inbox = outlook.GetDefaultFolder(6)  # Inbox
    target_folder = inbox.Folders[OUTLOOK_FOLDER]
except:
    print(f"Subfolder '{OUTLOOK_FOLDER}' not found inside your Inbox.")
    exit()

messages = target_folder.Items
messages.Sort("ReceivedTime", True)

def process_task(subject_match, destination_path):
    print("\n---------------------------------------------")
    print(f"Processing: {subject_match}")

    target_email = None

    for msg in messages:
        if msg.Class == 43:  # Mail item
            sender = msg.SenderName.strip().lower()
            subject = msg.Subject.strip().lower()

            if SENDER_NAME.lower() in sender and subject_match.lower() in subject:
                target_email = msg
                break

    if not target_email:
        print(f"No email found for subject: {subject_match}")
        return

    print(f"  Found email from: {target_email.SenderName}")
    print(f"  Subject: {target_email.Subject}")
    print(f"  Received: {target_email.ReceivedTime}")

    if target_email.Attachments.Count == 0:
        print("  Email has no attachments.")
        return

    attachment = target_email.Attachments.Item(1)

    # Save RAW file
    temp_raw = os.path.join(os.environ["TEMP"], attachment.FileName)
    attachment.SaveAsFile(temp_raw)

    print(f"  Downloaded attachment: {attachment.FileName}")

    # Copy raw file to destination
    shutil.copyfile(temp_raw, destination_path)
    os.remove(temp_raw)

    print(f"  Saved RAW file to: {destination_path}")



for task in TASKS:
    process_task(task["subject"], task["destination"])

print("\nAll email tasks completed.")



print("\nOpening Excel to refresh Power Query connections...")

excel = win32.Dispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Open(EXCEL_FILE)

# Collect query names before refresh
query_names = [q.Name for q in wb.Queries]

print("\nQueries found in workbook:")
for q in query_names:
    print(" -", q)

# Trigger refresh
wb.RefreshAll()

# Wait for refresh to complete
time.sleep(30)

# Save and close
wb.Save()
wb.Close()
excel.Quit()

print("\nThe following queries were refreshed and saved:")
for q in query_names:
    print(" -", q)

print("\nProcess completed successfully.")
