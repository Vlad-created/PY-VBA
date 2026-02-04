import win32com.client as win32
import os
import shutil

# Match parameters
SENDER_NAME = "SQL Replica"
SUBJECT_MATCH = "ESignConsent"

# Destination file to overwrite
destination_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\Adhoc\Check\ESignConsent_New.txt"

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

# Save XLS to temp
temp_xls = os.path.join(os.environ["TEMP"], attachment.FileName)
attachment.SaveAsFile(temp_xls)

print(f"Downloaded attachment: {attachment.FileName}")

# Convert XLS → TXT using Excel COM
excel = win32.Dispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Open(temp_xls)

# Save as TXT (tab-delimited)
temp_txt = os.path.join(os.environ["TEMP"], "ESignConsent_temp.txt")
wb.SaveAs(temp_txt, FileFormat=win32.constants.xlTextWindows)

wb.Close(False)
excel.Quit()

# Copy TXT over destination
shutil.copyfile(temp_txt, destination_path)

# Cleanup
os.remove(temp_xls)
os.remove(temp_txt)

print(f"Attachment converted and saved over:\n{destination_path}")
