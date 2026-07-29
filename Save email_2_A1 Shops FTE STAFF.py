import win32com.client as win32
import os
import shutil
from datetime import datetime, timedelta

# Match parameters
SENDER_NAME = "Teodora Karayazova"
SUBJECT_MATCH = "FTE monthly report"

# Destination folder
DEST_FOLDER = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Staff"

# ---------------------------------------------------------
# 1. Calculate last day of previous month
# ---------------------------------------------------------
today = datetime.today()
first_day_this_month = today.replace(day=1)
last_day_prev_month = first_day_this_month - timedelta(days=1)
date_str = last_day_prev_month.strftime("%Y_%m_%d")

new_filename = f"A1 Shops FTE {date_str}.xlsx"
destination_path = os.path.join(DEST_FOLDER, new_filename)

# ---------------------------------------------------------
# 2. Start Outlook and search email
# ---------------------------------------------------------
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # Inbox

messages = inbox.Items
messages.Sort("ReceivedTime", True)

target_email = None
for msg in messages:
    if msg.Class == 43:  # MailItem
        if msg.SenderName == SENDER_NAME and SUBJECT_MATCH in msg.Subject:
            target_email = msg
            break

if not target_email:
    print("No matching email found.")
    exit()

print(f"Found email from {target_email.SenderName} with subject: {target_email.Subject}")

# ---------------------------------------------------------
# 3. Save attachment
# ---------------------------------------------------------
if target_email.Attachments.Count == 0:
    print("Email has no attachments.")
    exit()

attachment = target_email.Attachments.Item(1)

temp_path = os.path.join(os.environ["TEMP"], attachment.FileName)
attachment.SaveAsFile(temp_path)

print(f"Downloaded attachment: {attachment.FileName}")

# Copy to destination with new name
shutil.copyfile(temp_path, destination_path)
os.remove(temp_path)

print(f"Saved as: {destination_path}")

# ---------------------------------------------------------
# 4. Open Excel and modify sheet REPORT
# ---------------------------------------------------------
excel = win32.Dispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Open(destination_path)
ws = wb.Worksheets("REPORT")

# ---------------------------------------------------------
# 5. Fill column AA
# ---------------------------------------------------------
ws.Range("AA1").Value = "SAP"

# Build formula with correct sheet name
sheet_name = f"New FTEs {last_day_prev_month.strftime('%d.%m.%Y')}"
formula = f'=IFNA(XLOOKUP(D2;\'{sheet_name}\'!F:F;\'{sheet_name}\'!E:E);"")'

last_row = ws.Cells(ws.Rows.Count, "D").End(-4162).Row  # xlUp = -4162

ws.Range(f"AA2:AA{last_row}").FormulaLocal = formula

# ---------------------------------------------------------
# 6. Check for empty cells
# ---------------------------------------------------------
empty_cells = ws.Range(f"AA2:AA{last_row}").SpecialCells(4) if True else None  # 4 = xlCellTypeBlanks

if empty_cells:
    print(f"⚠ Empty cells found in AA column: {empty_cells.Count}")
else:
    print("✔ No empty cells in column AA")

# ---------------------------------------------------------
# 7. Save and close
# ---------------------------------------------------------
wb.Save()
wb.Close()
excel.Quit()

print("Processing completed successfully.")
