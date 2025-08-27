import xlwings as xw
import os
from datetime import datetime, timedelta

# Define the workbook path directly
folder_path = (r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation")
workbook_path = os.path.join(folder_path, "REPORT SENDER.new.xlsm")

# Open the workbook and read region names from Sheet1
wb_master = xw.Book(workbook_path)
sheet = wb_master.sheets["Sheet1"]

# Loop through rows 2 to 15 in column B
for i in range(2, 16):
    region_name = str(sheet.range(f"B{i}").value).strip()
    if region_name:
        file_name = f"FBB_Cancellation Rate_{(datetime.today() - timedelta(days=1)).strftime('%d.%m')}_{region_name}.xlsx"
        full_path = os.path.join(folder_path, file_name)

        if os.path.exists(full_path):
            wb = xw.Book(full_path)
            wb.app.screen_updating = False
            wb.app.display_alerts = False
            wb.app.calculation = 'manual'

            wb.api.RefreshAll()
            wb.save()
            wb.close()
        else:
            print(f"⚠️ File not found: {file_name}")

print("Optimized refresh complete for all available region files.")
