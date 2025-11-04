import os
import re
from datetime import datetime, timedelta
import win32com.client as win32
import xlwings as xw
from dateutil.relativedelta import relativedelta

# === PART 0: Locate and open yesterday's file ===

folder_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\Base FBB Cancellation_Rate"
yesterday = datetime.today() - timedelta(days=1)
target_date_str = yesterday.strftime("%d.%m.%Y")
filename_pattern = f"Shops Report - {target_date_str}"

matched_file = None
for file in os.listdir(folder_path):
    if filename_pattern in file and file.lower().endswith((".xlsx", ".xlsm", ".xls")):
        matched_file = os.path.join(folder_path, file)
        break

if not matched_file:
    print(f"❌ No file found for yesterday: {filename_pattern}")
    exit()

excel = win32.Dispatch("Excel.Application")
excel.Visible = True
wb_com = excel.Workbooks.Open(matched_file)
print(f"📂 Opened file: {matched_file}")

# === PART 1: Filter PivotTable and extract Grand Total ===

today = datetime.today()
months = [str((today - relativedelta(months=i)).month) for i in range(4)]
telco_filter = ["Cable Internet", "Unlimited Internet"]
channel_filter = ["A1 Shops", "Franchise"]

def apply_filter(pivot, field_name, allowed_items):
    try:
        pf = pivot.PivotFields(field_name)
        pf.ClearAllFilters()
        pf.EnableMultiplePageItems = True
        for item in pf.PivotItems():
            try: item.Visible = True
            except: pass
        for item in pf.PivotItems():
            try: item.Visible = item.Name in allowed_items
            except: pass
        print(f"✅ Filter applied on '{field_name}' → {allowed_items}")
    except Exception as e:
        print(f"❌ Error applying filter on '{field_name}': {e}")

try:
    sheet = wb_com.Sheets("Main_Report")
except Exception as e:
    print(f"❌ Sheet 'Main_Report' not found: {e}")
    exit()

for pivot in sheet.PivotTables():
    print(f"\n📊 PivotTable: {pivot.Name} on sheet '{sheet.Name}'")
    try:
        field_names = [pf.Name for pf in pivot.PivotFields()]
    except Exception as e:
        print(f"❌ Error retrieving field names: {e}")
        continue

    if "ORDER_CREATE_MONTH" in field_names:
        apply_filter(pivot, "ORDER_CREATE_MONTH", months)
    else:
        print("⚠️ Field 'ORDER_CREATE_MONTH' not found.")

    if "Common_Telco" in field_names:
        apply_filter(pivot, "Common_Telco", telco_filter)
    else:
        print("⚠️ Field 'Common_Telco' not found.")

    if "CLS_CHANNEL" in field_names:
        apply_filter(pivot, "CLS_CHANNEL", channel_filter)
    else:
        print("⚠️ Field 'CLS_CHANNEL' not found.")

    try:
        last_cell = pivot.TableRange2.Cells(pivot.TableRange2.Rows.Count, pivot.TableRange2.Columns.Count)
        last_cell.ShowDetail = True
        new_sheet = wb_com.Sheets(wb_com.Sheets.Count)
        new_sheet.Name = "Sheet1"
        print("📄 Grand Total detail extracted and renamed to 'Sheet1'.")
    except Exception as e:
        print(f"❌ Error extracting Grand Total or renaming sheet: {e}")
    break

# === PART 2: Format columns in Sheet1 using xlwings ===

app = xw.apps.active
wb_xw = app.books[wb_com.Name]

try:
    ws = wb_xw.sheets["Sheet1"]
    for col in ["A", "F", "Y", "AS"]:
        ws.range(f"{col}:{col}").number_format = "@"
        last_row = ws.range(f"{col}{ws.cells.last_cell.row}").end("up").row
        rng = ws.range(f"{col}1:{col}{last_row}")
        values = rng.value
        if not isinstance(values, list): values = [[values]]
        converted = []
        for v in values:
            val = v[0] if isinstance(v, list) else v
            text = "" if val is None else str(val)
            if isinstance(val, (int, float)):
                text = f"{val:.0f}" if float(val).is_integer() else f"{val}"
            converted.append([text])
        rng.value = converted
    print(f"✅ Converted to text without scientific notation: {wb_xw.name}")
except Exception as e:
    print(f"⚠️ Error formatting Sheet1: {e}")
