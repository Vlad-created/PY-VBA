import win32com.client as win32
import os
import shutil

# Match parameters
SENDER_NAME = "SQL Replica"
SUBJECT_MATCH = "Telesales"

# Destination file to overwrite
destination_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\Adhoc\Check\telesales.txt"

# Start Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Access Inbox → SQL Replica subfolder
try:
    inbox = outlook.GetDefaultFolder(6)  # Inbox
    target_folder = inbox.Folders["SQL Replica"]
except:
    print("Subfolder 'SQL Replica' not found inside your Inbox.")
    exit()

# Get newest messages first
messages = target_folder.Items
messages.Sort("ReceivedTime", True)

# Find the newest matching email
target_email = None
for msg in messages:
    if msg.Class == 43:  # MailItem
        sender = msg.SenderName.strip().lower()
        subject = msg.Subject.strip().lower()
        if SENDER_NAME.lower() in sender and SUBJECT_MATCH.lower() in subject:
            target_email = msg
            break

if not target_email:
    print("No matching email found in folder 'SQL Replica'.")
    exit()

print(f"Found email:")
print(f"  From: {target_email.SenderName}")
print(f"  Subject: {target_email.Subject}")
print(f"  Received: {target_email.ReceivedTime}")

# Check attachments
if target_email.Attachments.Count == 0:
    print("Email has no attachments.")
    exit()

attachment = target_email.Attachments.Item(1)

# Save attachment EXACTLY as-is (raw bytes)
temp_raw = os.path.join(os.environ["TEMP"], attachment.FileName)
attachment.SaveAsFile(temp_raw)

print(f"Downloaded attachment: {attachment.FileName}")

# Copy raw file to destination but rename to .txt
shutil.copyfile(temp_raw, destination_path)

# Cleanup
os.remove(temp_raw)

print(f"Attachment saved RAW (no Excel conversion) to:\n{destination_path}")
