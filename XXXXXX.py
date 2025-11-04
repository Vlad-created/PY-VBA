import os
from datetime import datetime, timedelta
import win32com.client as win32
import xlwings as xw
from dateutil.relativedelta import relativedelta

def col_letter_to_index(letter):
    letter = letter.upper()
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx

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

# === PART 1: Filter PivotTable and extract Grand Total detail safely ===
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

extraction_sheet = None
# capture sheet names before extraction
before_sheets = [s.Name for s in wb_com.Sheets]

for pivot in sheet.PivotTables():
    try:
        field_names = [pf.Name for pf in pivot.PivotFields()]
    except Exception:
        continue

    if "ORDER_CREATE_MONTH" in field_names:
        apply_filter(pivot, "ORDER_CREATE_MONTH", months)
    if "Common_Telco" in field_names:
        apply_filter(pivot, "Common_Telco", telco_filter)
    if "CLS_CHANNEL" in field_names:
        apply_filter(pivot, "CLS_CHANNEL", channel_filter)

    try:
        # ShowDetail on the Grand Total cell (bottom-right of pivot)
        last_cell = pivot.TableRange2.Cells(pivot.TableRange2.Rows.Count, pivot.TableRange2.Columns.Count)
        last_cell.ShowDetail = True
        # wait a moment for Excel to create the sheet (no sleep since synchronous COM should create it)
        after_sheets = [s.Name for s in wb_com.Sheets]
        new_sheets = [n for n in after_sheets if n not in before_sheets]
        if not new_sheets:
            # fallback: sometimes Excel creates a sheet but reuses a name; find the sheet with large usedrange
            candidates = []
            for s in wb_com.Sheets:
                if s.Name not in before_sheets:
                    candidates.append(s)
            if candidates:
                extraction_sheet = candidates[-1]
            else:
                # try last sheet
                extraction_sheet = wb_com.Sheets(wb_com.Sheets.Count)
        else:
            extraction_sheet = wb_com.Sheets(new_sheets[-1])
        # ensure name Sheet1 exists and is unique: delete existing Sheet1 first
        try:
            if "Sheet1" in [s.Name for s in wb_com.Sheets]:
                # don't delete the newly created sheet if it is already Sheet1
                if extraction_sheet.Name != "Sheet1":
                    excel.DisplayAlerts = False
                    wb_com.Sheets("Sheet1").Delete()
                    excel.DisplayAlerts = True
        except Exception:
            excel.DisplayAlerts = True

        # rename the extracted sheet to Sheet1
        try:
            extraction_sheet.Name = "Sheet1"
        except Exception as e:
            # if rename fails because name exists, ensure we point to the correct sheet
            extraction_sheet = wb_com.Sheets("Sheet1")
        print("📄 Grand Total detail extracted and renamed to 'Sheet1'.")
    except Exception as e:
        print(f"❌ Error extracting Grand Total or renaming sheet: {e}")
    break

# verify extraction sheet has data rows
try:
    s = wb_com.Sheets("Sheet1")
    used_rows = s.UsedRange.Rows.Count
    used_cols = s.UsedRange.Columns.Count
    print(f"ℹ️ Sheet1 used range rows={used_rows}, cols={used_cols}")
    if used_rows < 2:
        print("⚠️ Extracted sheet has no data rows. Aborting copy.")
        # leave workbooks open for inspection
        exit()
except Exception as e:
    print(f"❌ Unable to inspect Sheet1: {e}")
    exit()

# === PART 2: Attach to workbook via xlwings and prepare source sheet ===
# ensure xlwings has the same workbook opened
app = xw.apps.active
# locate xlwings book by full path if possible (safer than by name)
wb_xw = None
for b in app.books:
    try:
        if os.path.abspath(b.fullname) == os.path.abspath(wb_com.FullName):
            wb_xw = b
            break
    except Exception:
        pass
if wb_xw is None:
    wb_xw = app.books.open(wb_com.FullName)
ws_xw = wb_xw.sheets["Sheet1"]

# convert certain columns to text in source to avoid scientific notation and preserve IDs
for col in ["A", "F", "Y", "AS"]:
    rng_last = ws_xw.range(f"{col}{ws_xw.cells.last_cell.row}").end("up").row
    if rng_last < 1:
        continue
    rng = ws_xw.range(f"{col}1:{col}{rng_last}")
    rng.number_format = "@"
    values = rng.value
    if values is None:
        continue
    if not isinstance(values, list):
        values = [[values]]
    converted = []
    for v in values:
        val = v[0] if isinstance(v, list) else v
        text = "" if val is None else str(val)
        if isinstance(val, (int, float)):
            text = f"{val:.0f}" if float(val).is_integer() else f"{val}"
        converted.append([text])
    rng.value = converted
print("✅ Source text columns converted in Sheet1 (A,F,Y,AS).")

# === PART 3: Open destination workbook and delete last 4 months ===
target_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\data.xlsx"
wb_target_com = excel.Workbooks.Open(target_path)
ws_target_com = wb_target_com.Sheets("Detail1")

last_4_months = [str((today - relativedelta(months=i)).month) for i in range(4)]
ws_target_com.Range("H1").AutoFilter(Field=8, Criteria1=last_4_months, Operator=7)

try:
    used_rows_target = ws_target_com.UsedRange.Rows.Count
    if used_rows_target > 1:
        try:
            visible_range = ws_target_com.Range("A2:A" + str(used_rows_target)).SpecialCells(12)
            visible_range.EntireRow.Delete()
            print(f"🧹 Deleted filtered rows for months: {last_4_months}")
        except Exception as e:
            print(f"⚠️ No visible rows to delete or error occurred: {e}")
    else:
        print("⚠️ Target sheet has no data rows to evaluate for deletion.")
except Exception as e:
    print(f"⚠️ Error while attempting to delete rows: {e}")

try:
    if ws_target_com.FilterMode and ws_target_com.ShowAllData:
        ws_target_com.ShowAllData()
    if ws_target_com.AutoFilterMode:
        ws_target_com.AutoFilterMode = False
    print("🚫 Filters fully cleared on target.")
except Exception as e:
    print(f"⚠️ Error resetting filters: {e}")

# attach target to xlwings
wb_target_xw = None
for b in app.books:
    try:
        if os.path.abspath(b.fullname) == os.path.abspath(wb_target_com.FullName):
            wb_target_xw = b
            break
    except Exception:
        pass
if wb_target_xw is None:
    wb_target_xw = app.books.open(wb_target_com.FullName)
ws_target_xw = wb_target_xw.sheets["Detail1"]

# === PART 4: Read source data, convert date columns to dd.mm.yyyy text, and paste ===
source_last_row = ws_xw.range("A" + str(ws_xw.cells.last_cell.row)).end("up").row
source_last_col = ws_xw.range("A1").end("right").column
if source_last_row < 2:
    print("⚠️ No rows in source to copy. Aborting.")
    exit()

data_range = ws_xw.range((2, 1), (source_last_row, source_last_col)).value
if not data_range:
    print("⚠️ data_range is empty. Aborting.")
    exit()

# build date column indices (zero-based)
date_cols = ["C", "D", "I", "L", "O", "AI", "AJ", "BT", "BU", "BV", "BW"]
date_col_indices = [col_letter_to_index(c) - 1 for c in date_cols]

converted_data = []
for row_idx, row in enumerate(data_range, start=2):
    # xlwings returns single row as tuple, multi as list of tuples
    if not isinstance(row, (list, tuple)):
        row = (row,)
    new_row = []
    for i, val in enumerate(row):
        if i in date_col_indices and val not in (None, ""):
            try:
                if isinstance(val, float):
                    # Excel serial to datetime (handles Excel 1900 system)
                    serial = int(val)
                    py_dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + serial - 2)
                    new_row.append(py_dt.strftime("%d.%m.%Y"))
                elif isinstance(val, datetime):
                    new_row.append(val.strftime("%d.%m.%Y"))
                else:
                    # if it's a string, try to normalize common separators
                    s = str(val).strip()
                    # if already in dd.mm.yyyy keep it
                    new_row.append(s)
            except Exception:
                new_row.append(str(val))
        else:
            new_row.append(val)
    converted_data.append(new_row)

if not converted_data:
    print("⚠️ No converted rows prepared to paste. Aborting.")
    exit()

# determine first empty row in target
target_last_row = ws_target_xw.range("A" + str(ws_target_xw.cells.last_cell.row)).end("up").row + 1

# Paste values
ws_target_xw.range((target_last_row, 1)).value = converted_data
print(f"📥 Pasted {len(converted_data)} rows into {wb_target_xw.fullname} at row {target_last_row}")

# === PART 5: Apply column formats on target AFTER pasting ===
# text columns
for col in ["A", "F", "Y", "AS"]:
    ws_target_xw.range(f"{col}:{col}").number_format = "@"

# numbers
for col in ["H", "AP"]:
    ws_target_xw.range(f"{col}:{col}").number_format = "0"

# date columns
for col in date_cols:
    ws_target_xw.range(f"{col}:{col}").number_format = "dd.mm.yyyy"

print("🎯 Applied destination formats after paste.")

# === PART 6: Add formulas to CP and CQ and fill down if rows exist ===
ws = ws_target_xw
ws.range("CP2").formula = '=IFNA(XLOOKUP(CE2,Maping!B:B,Maping!C:C)," ")'
ws.range("CQ2").formula = '=IF(AND(CD2="Refused",CG2="Sales",CP2="поради отказ на клиента от услуга"),1,0)'

last_row = ws.range("A" + str(ws.cells.last_cell.row)).end("up").row
if last_row >= 2:
    ws.range(f"CP2:CQ{last_row}").api.FillDown()
    print("🧮 Formulas added and filled down in CP and CQ.")
else:
    print("⚠️ No rows to fill formulas into.")

print("📝 Workbooks remain open for manual review. Save when ready.")
