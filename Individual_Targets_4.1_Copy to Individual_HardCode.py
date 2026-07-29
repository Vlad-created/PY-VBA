import openpyxl
import xlwings as xw

# Source and destination paths
src_path = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets\Templete_newmonth.xlsx"
dst_path = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets\Individual January'26.xlsm"

# Load source workbook (openpyxl)
src_wb = openpyxl.load_workbook(src_path, data_only=True)
src_ws = src_wb.active

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

# Open destination workbook (xlwings) WITHOUT creating Book1
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
