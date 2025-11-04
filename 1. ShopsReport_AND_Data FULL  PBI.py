import os
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
    try:
        field_names = [pf.Name for pf in pivot.PivotFields()]
    except Exception as e:
        continue

    if "ORDER_CREATE_MONTH" in field_names:
        apply_filter(pivot, "ORDER_CREATE_MONTH", months)
    if "Common_Telco" in field_names:
        apply_filter(pivot, "Common_Telco", telco_filter)
    if "CLS_CHANNEL" in field_names:
        apply_filter(pivot, "CLS_CHANNEL", channel_filter)

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

# === PART 3: Open destination workbook and delete last 4 months ===
target_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\PBI-Excel\data.xlsx"
wb_target = excel.Workbooks.Open(target_path)
ws_target = wb_target.Sheets("Detail1")

last_4_months = [str((today - relativedelta(months=i)).month) for i in range(4)]
ws_target.Range("H1").AutoFilter(Field=8, Criteria1=last_4_months, Operator=7)

try:
    visible_range = ws_target.Range("A2:A" + str(ws_target.UsedRange.Rows.Count)).SpecialCells(12)
    visible_range.EntireRow.Delete()
    print(f"🧹 Deleted filtered rows for months: {last_4_months}")
except Exception as e:
    print(f"⚠️ No visible rows to delete or error occurred: {e}")

try:
    if ws_target.FilterMode and ws_target.ShowAllData:
        ws_target.ShowAllData()
    if ws_target.AutoFilterMode:
        ws_target.AutoFilterMode = False
    print("🚫 Filters fully cleared.")
except Exception as e:
    print(f"⚠️ Error resetting filters: {e}")

# === PART 4: Copy Sheet1 data to first empty row in Detail1 as values ===
ws_source = wb_xw.sheets["Sheet1"]
source_last_row = ws_source.range("A" + str(ws_source.cells.last_cell.row)).end("up").row
source_last_col = ws_source.range("A1").end("right").column
data_range = ws_source.range((2, 1), (source_last_row, source_last_col))  # Exclude header

ws_target_xw = xw.books[wb_target.Name].sheets["Detail1"]
target_last_row = ws_target_xw.range("A" + str(ws_target_xw.cells.last_cell.row)).end("up").row + 1

# 🛡️ Format specific destination columns BEFORE pasting
text_cols = ["A", "F", "Y", "AS"]
for col in text_cols:
    ws_target_xw.range(f"{col}:{col}").number_format = "@"

# Format H and AP as numbers with no decimals
for col in ["H", "AP"]:
    ws_target_xw.range(f"{col}:{col}").number_format = "0"

# Paste as values only
ws_target_xw.range(f"A{target_last_row}").value = data_range.value
print(f"📥 Data pasted as text values into {target_path} → Sheet 'Detail1' at row {target_last_row}")

# === PART 5: Add formulas to CP and CQ ===
ws = ws_target_xw
ws.range("CP2").formula = '=IFNA(XLOOKUP(CE2,Maping!B:B,Maping!C:C)," ")'
ws.range("CQ2").formula = '=IF(AND(CD2="Refused",CG2="Sales",CP2="поради отказ на клиента от услуга"),1,0)'
ws.range("CR2").formula = '=XLOOKUP(P2,RollOut!E:E,RollOut!A:A)'
ws.range("CS2").formula = '=XLOOKUP(P2,RollOut!E:E,RollOut!B:B)'
ws.range("CT2").formula = '=XLOOKUP(P2,RollOut!E:E,RollOut!C:C)'
ws.range("CU2").formula = '=XLOOKUP(P2,RollOut!E:E,RollOut!D:D)'
ws.range("CV2").formula = '=XLOOKUP(P2,RollOut!E:E,RollOut!F:F)'
ws.range("CW2").formula = '=XLOOKUP(P2,RollOut!E:E,RollOut!G:G)'
ws.range("CX2").formula = '=XLOOKUP(P2,RollOut!E:E,RollOut!H:H)'
ws.range("CY2").formula = '=XLOOKUP(P2,RollOut!E:E,RollOut!I:I)'
ws.range("CZ2").formula = '=XLOOKUP(P2,RollOut!E:E,RollOut!J:J)'
ws.range("DA2").formula = '=IF(CD2="Performed",1,0)'
ws.range("DB2").formula = '=IF(CD2= "Backlog",1,0)'
ws.range("DC2").formula = '=IF(CD2 = "Refused",1,0)'
ws.range("DD2").formula = '=IFNA(XLOOKUP(AP2,sentToShops!B:B,sentToShops!D:D)," ")'
ws.range("DE2").formula = '=IF(AND(CD2="Refused",DD2="1"),1," ")'
ws.range("DF2").formula = '=IF(AND(CD2="Refused",DD2="0"),1," ")'
ws.range("DG2").formula = '=IF(AND(CD2="Refused",DD2="NULL"),1," ")'
ws.range("DH2").formula = '=IFNA(XLOOKUP(AP2,sentToShops!B:B,sentToShops!F:F)," ")'
ws.range("DI2").formula = '=IFNA(XLOOKUP(AP2,sentToShops!B:B,sentToShops!H:H)," ")'
ws.range("DJ2").formula = '="VL-" & Y2 & "-" & TEXT(NOW(),"yyyymmdd")'

last_row = ws.range("A" + str(ws.cells.last_cell.row)).end("up").row
ws.range(f"CP2:DJ{last_row}").api.FillDown()
print("🧮 Formulas added and filled down in CP and DJ.")

# === PART 5.1: Find & Replace in column P ===
# Read column P values from row 2 down to last_row, perform replacements, write back
p_start_row = 2
p_end_row = last_row  # uses last_row computed above
if p_end_row >= p_start_row:
    p_range = ws.range(f"P{p_start_row}:P{p_end_row}")
    p_values = p_range.value
    # Normalize to list of lists
    if p_values is None:
        p_values = []
    elif not isinstance(p_values, list):
        p_values = [[p_values]]
    # Perform replacements
    updated = []
    for row in p_values:
        val = row[0] if isinstance(row, list) else row
        if val is None:
            updated.append([val])
            continue
        s = str(val)
        # Replacement 1
        if "А1 м-н Бургас,МОЛ Бургас Плаза" in s:
            s = s.replace("А1 м-н Бургас,МОЛ Бургас Плаза", "А1 м-н Бургас, МОЛ Гранд Плаза")
        # Replacement 2
        if "А1 м-н София, Бот. шосе 515 ХОпарк" in s:
            s = s.replace("А1 м-н София, Бот. шосе 515 ХОпарк", "А1 м-н София, ХО Парк")
        updated.append([s])
    # Write back only if there are changes
    p_range.value = updated
    print(f"🔁 Performed Find & Replace in P{p_start_row}:P{p_end_row}")
else:
    print("ℹ️ No rows in column P to process for Find & Replace.")

# === PART 6: Leave workbooks open for manual review ===
print("📝 Workbooks remain open — review, save, or close manually as needed.")
