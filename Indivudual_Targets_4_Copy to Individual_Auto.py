import openpyxl
import xlwings as xw
import os
import datetime

# Folder containing the destination file
folder = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets"

# Source template
src_path = os.path.join(folder, "Templete_newmonth.xlsx")

# Detect current month (e.g., "February")
current_month = datetime.datetime.now().strftime("%B")

# Find the destination file dynamically
dst_path = None
for filename in os.listdir(folder):
    if filename.startswith("Individual ") and filename.endswith(".xlsm") and current_month in filename:
        dst_path = os.path.join(folder, filename)
        break

if dst_path is None:
    raise FileNotFoundError(
        f"No 'Individual <Month>.xlsm' file found for {current_month} in folder:\n{folder}"
    )

print(f"Using destination file:\n{dst_path}")

# Load source workbook (openpyxl)
src_wb = openpyxl.load_workbook(src_path, data_only=True)

# Explicitly select sheet "Templete"
if "Templete" not in src_wb.sheetnames:
    raise ValueError("Sheet 'Templete' not found in source workbook.")

src_ws = src_wb["Templete"]

# Determine last row in source (column B)
last_row = src_ws.max_row

# Read data from columns B–F, starting at row 2
data = []
for row in range(2, last_row + 1):
    row_values = [
        src_ws[f"B{row}"].value,
        src_ws[f"C{row}"].value,
        src_ws[f"D{row}"].value,
        src_ws[f"E{row}"].value,
        src_ws[f"F{row}"].value,
    ]
    data.append(row_values)

print(f"Loaded {len(data)} rows from source file.")

# Open destination workbook (xlwings)
app = xw.App(visible=True, add_book=False)
dst_wb = app.books.open(dst_path)

# Select the correct sheet
dst_ws = dst_wb.sheets["All_FTEs"]

# Write data into columns E–I starting at row 2
dst_ws.range("E2").value = data

# Save but keep workbook open
dst_wb.save()

print("Data successfully copied into sheet 'All_FTEs'.")
print("Excel remains open for review.")
