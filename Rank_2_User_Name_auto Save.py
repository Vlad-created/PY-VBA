import win32com.client as win32
import os
import shutil

# Match parameters
SENDER_NAME = "Revenue BZ"
SUBJECT_MATCH = "Regular report reg_user_shop_info is done"

# Destination file to overwrite
destination_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\User Names.xlsx"

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

print(f"Found email:")
print(f"  From: {target_email.SenderName}")
print(f"  Subject: {target_email.Subject}")
print(f"  Received: {target_email.ReceivedTime}")

# Check attachments
if target_email.Attachments.Count == 0:
    print("Email has no attachments.")
    exit()

attachment = target_email.Attachments.Item(1)

# Save to temp
temp_path = os.path.join(os.environ["TEMP"], attachment.FileName)
attachment.SaveAsFile(temp_path)

print(f"Downloaded attachment: {attachment.FileName}")

# Copy over existing Excel file (works across drives)
shutil.copyfile(temp_path, destination_path)

# Remove temp file
os.remove(temp_path)

print(f"Attachment saved over existing file:\n{destination_path}")
