import win32com.client as win32
import xlwings as xw
from datetime import datetime
from dateutil.relativedelta import relativedelta

# === PART 1: Filter PivotTable and extract Grand Total ===

# Connect to Excel via COM
excel = win32.GetActiveObject("Excel.Application")
wb_com = excel.ActiveWorkbook

# Get last 3 months as numbers (e.g., ["9", "8", "7"])
today = datetime.today()
months = [str((today - relativedelta(months=i)).month) for i in range(3)]

# Define filters
telco_filter = ["Cable Internet", "Unlimited Internet"]
channel_filter = ["A1 Shops", "Franchise"]

# Helper to safely apply filters
def apply_filter(pivot, field_name, allowed_items):
    try:
        pf = pivot.PivotFields(field_name)
        pf.ClearAllFilters()
        pf.EnableMultiplePageItems = True

        # Make all items visible first
        for item in pf.PivotItems():
            try:
                item.Visible = True
            except:
                pass

        # Then hide items not in allowed list
        for item in pf.PivotItems():
            try:
                item.Visible = item.Name in allowed_items
            except:
                pass

        print(f"✅ Filter applied on '{field_name}' → {allowed_items}")
    except Exception as e:
        print(f"❌ Error applying filter on '{field_name}': {e}")

# Target only the sheet named "Main_Report"
try:
    sheet = wb_com.Sheets("Main_Report")
except Exception as e:
    print(f"❌ Sheet 'Main_Report' not found: {e}")
    exit()

# Apply filters to first PivotTable in Main_Report
for pivot in sheet.PivotTables():
    print(f"\n📊 PivotTable: {pivot.Name} on sheet '{sheet.Name}'")
    print("🔍 Available fields:")
    try:
        for pf in pivot.PivotFields():
            print(f"   - {pf.Name}")
    except Exception as e:
        print(f"❌ Error listing fields: {e}")
        continue

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

    # Double-click Grand Total cell to extract data
    try:
        last_cell = pivot.TableRange2.Cells(pivot.TableRange2.Rows.Count, pivot.TableRange2.Columns.Count)
        last_cell.ShowDetail = True
        print("📄 Grand Total detail extracted.")

        # Rename the newly created sheet to "Sheet1"
        new_sheet = wb_com.Sheets(wb_com.Sheets.Count)
        new_sheet.Name = "Sheet1"
        print("📝 Drilldown sheet renamed to 'Sheet1'.")

    except Exception as e:
        print(f"❌ Error extracting Grand Total or renaming sheet: {e}")

    break  # Stop after first PivotTable

# === PART 2: Format columns in Sheet1 using xlwings ===

# Connect to Excel via xlwings
app = xw.apps.active
wb_xw = app.books[wb_com.Name]

try:
    ws = wb_xw.sheets["Sheet1"]

    for col in ["A", "F", "Y", "AS"]:
        ws.range(f"{col}:{col}").number_format = "@"

        last_row = ws.range(f"{col}{ws.cells.last_cell.row}").end("up").row
        rng = ws.range(f"{col}1:{col}{last_row}")
        values = rng.value

        if not isinstance(values, list):
            values = [[values]]

        converted = []
        for v in values:
            val = v[0] if isinstance(v, list) else v
            if val is None:
                converted.append([""])
            else:
                text = str(val)
                if isinstance(val, (int, float)):
                    text = f"{val:.0f}" if float(val).is_integer() else f"{val}"
                converted.append([text])

        rng.value = converted

    print(f"✅ Converted to text without scientific notation: {wb_xw.name}")

except Exception as e:
    print(f"⚠️ Error formatting Sheet1: {e}")
