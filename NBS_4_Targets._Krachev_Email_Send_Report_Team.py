import os
import win32com.client as win32

folder_path = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Krachev_ Crak"

required_files = [
    "Пловдив 1",
    "Пловдив 2",
    "Пловдив 3",
    "Стара Загора"
]

recipient = "vladislav.kolesnik@a1.bg; mario.l.dimitrov@a1.bg"

email_body = """
Здравей,<br><br>
Изпращам таргетите по региони.<br><br>
Поздрави
"""

SENDER_EMAIL = "ReportsTeam@a1.bg"

def send_files():
    outlook = win32.Dispatch("Outlook.Application")
    session = outlook.Session

    # Find the correct sending account
    send_account = None
    for account in session.Accounts:
        if account.SmtpAddress.lower() == SENDER_EMAIL.lower():
            send_account = account
            break

    if not send_account:
        print(f"⚠ Cannot find Outlook account: {SENDER_EMAIL}")
        print("Make sure the mailbox is added to your Outlook profile with send-as rights.")
        return

    mail = outlook.CreateItem(0)
    mail.SendUsingAccount = send_account

    mail.To = recipient
    mail.Subject = "Прикачени регионални файлове NBS_Targets"
    mail.HTMLBody = email_body

    missing_files = []

    for base_name in required_files:
        file_path = os.path.join(folder_path, base_name + ".xlsx")
        if os.path.exists(file_path):
            mail.Attachments.Add(file_path)
            print(f"✔ Attached: {file_path}")
        else:
            missing_files.append(file_path)
            print(f"⚠ Missing file: {file_path}")

    if missing_files:
        print("🚫 Email not sent due to missing files.")
    else:
        mail.Send()
        print(f"📨 Email sent successfully from {SENDER_EMAIL}")

if __name__ == "__main__":
    send_files()
