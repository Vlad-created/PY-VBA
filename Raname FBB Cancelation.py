import os
from datetime import datetime, timedelta
import xlwings as xw

# Define folder path
folder_path = (r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation")

# Get yesterday's date in dd.mm format
yesterday_date = (datetime.today() - timedelta(days=1)).strftime("%d.%m")

# Open the workbook that contains region names
wb = xw.Book(os.path.join(folder_path, "REPORT SENDER.new.xlsm"))
sheet = wb.sheets["Sheet1"]

# Loop through rows 2 to 15 in column B
for i in range(2, 16):
    region_name = str(sheet.range(f"B{i}").value).strip()
    if region_name:
        # Loop through all files in the folder
        for file_name in os.listdir(folder_path):
            if (
                "FBB_Cancellation Rate_" in file_name and
                f"_{region_name}.xlsx" in file_name and
                yesterday_date not in file_name
            ):
                # Build new file name
                new_file_name = f"FBB_Cancellation Rate_{yesterday_date}_{region_name}.xlsx"
                old_path = os.path.join(folder_path, file_name)
                new_path = os.path.join(folder_path, new_file_name)

                try:
                    os.rename(old_path, new_path)
                    print(f"✅ Renamed: {file_name} → {new_file_name}")
                except Exception as e:
                    print(f"⚠️ Error renaming {file_name}: {e}")

print("Files renamed successfully.")
