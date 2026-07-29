import win32com.client as win32
import os

# Outlook search parameters
SENDER = "revenue.bz@a1.bg"
SUBJECT = "Regular report reg_user_shop_info is done. Details in mail."

# Destination file (existing Excel file to overwrite)
destination_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\User Names.xlsx"

# Start Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # Inbox

# Use Restrict to filter emails instantly
filter_query = f"[SenderEmailAddress] = '{SENDER}' AND [Subject] = '{SUBJECT}'"
messages = inbox.Items.Restrict(filter_query)

# Sort newest first
messages.Sort("ReceivedTime", True)

if len(messages) == 0:
    print("No matching email found.")
    exit()

email = messages[0]

print(f"Found email:")
print(f"  From: {email.SenderEmailAddress}")
print(f"  Subject: {email.Subject}")
print(f"  Received: {email.ReceivedTime}")

# Check attachments
if email.Attachments.Count == 0:
    print("Email has no attachments.")
    exit()

attachment = email.Attachments.Item(1)

# Save attachment to temp
temp_path = os.path.join(os.environ["TEMP"], attachment.FileName)
attachment.SaveAsFile(temp_path)

print(f"Downloaded attachment: {attachment.FileName}")

# Overwrite existing Excel file
os.replace(temp_path, destination_path)

print(f"Attachment saved over existing file:\n{destination_path}")
