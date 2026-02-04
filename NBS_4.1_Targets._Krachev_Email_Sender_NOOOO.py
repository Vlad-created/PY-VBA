import os
import win32com.client as win32


# Folder containing the files
folder_path = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Krachev_ Crak"

# Files you want to send (exact names without extension)
required_files = [
    "Пловдив 1",
    "Пловдив 2",
    "Пловдив 3",
    "Стара Загора"
]

# Email recipients
recipient = "ivan.krachev@a1.bg; ReportsTeam@a1.bg"

# Email body (HTML)
email_body = """
Здравей,<br><br>
Изпращам таргетите по региони.<br><br>
Поздрави
"""



def send_files():
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = recipient
    mail.Subject = "Прикачени регионални файлове NBS_Targets"
    mail.HTMLBody = email_body

    # Attach files
    for base_name in required_files:
        file_path = os.path.join(folder_path, base_name + ".xlsx")

        if os.path.exists(file_path):
            mail.Attachments.Add(file_path)
            print(f"✔ Attached: {file_path}")
        else:
            print(f"⚠ Missing file: {file_path}")

    mail.Send()
    print("📨 Email sent successfully.")

if __name__ == "__main__":
    send_files()
