import win32com.client as win32
import os
import shutil

# Match parameters
SENDER_NAME = "Revenue BZ"
SUBJECT_MATCH = "Regular report reg_user_shop_info is done"

# Destination files to overwrite
destination_paths = [
    r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\User Names.xlsx",
    r"\\virgo\work\OpControl\2016\deviceGames\reg_user_shop_info.xlsx"
]

def is_file_locked(path):
    if not os.path.exists(path):
        return False
    try:
        with open(path, "a"):
            return False
    except PermissionError:
        return True

# Start Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # Inbox

# Get newest messages first
messages = inbox.Items
messages.Sort("ReceivedTime", True)

# Find the newest matching email
target_email = None
for msg in messages:
    if msg.Class == 43:  # MailItem
        if msg.SenderName == SENDER_NAME and SUBJECT_MATCH in msg.Subject:
            target_email = msg
            break

if not target_email:
    print("No matching email found.")
    exit()

print("Found email:")
print(f"  From: {target_email.SenderName}")
print(f"  Subject: {target_email.Subject}")
print(f"  Received: {target_email.ReceivedTime}")

# Find Excel attachment
attachment = None
for att in target_email.Attachments:
    if att.FileName.lower().endswith((".xlsx", ".xlsm", ".xls")):
        attachment = att
        break

if not attachment:
    print("Email has no Excel attachment.")
    exit()

# Save to temp
temp_path = os.path.join(os.environ["TEMP"], attachment.FileName)
attachment.SaveAsFile(temp_path)

print(f"Downloaded attachment: {attachment.FileName}")

# Overwrite each destination file
for dest in destination_paths:
    print(f"\nProcessing: {dest}")

    if is_file_locked(dest):
        print("  ERROR: File is OPEN or LOCKED. Skipping.")
        continue

    try:
        shutil.copyfile(temp_path, dest)
        print("  SUCCESS: File overwritten.")
    except Exception as e:
        print(f"  FAILED: {e}")

# Remove temp file
os.remove(temp_path)
print("\nAll done.")
