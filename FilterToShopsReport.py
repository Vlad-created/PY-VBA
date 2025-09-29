import win32com.client as win32
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Connect to Excel
excel = win32.GetActiveObject("Excel.Application")
wb = excel.ActiveWorkbook

# Get last 4 months as numbers (e.g., ["9", "8", "7", "6"])
today = datetime.today()
months = [str((today - relativedelta(months=i)).month) for i in range(4)]

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
    sheet = wb.Sheets("Main_Report")
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

    # Apply filters only if fields exist
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

        # Rename the newly created sheet to "sheet1"
        new_sheet = wb.Sheets(wb.Sheets.Count)
        new_sheet.Name = "Sheet1"
        print("📝 Drilldown sheet renamed to 'Sheet1'.")

    except Exception as e:
        print(f"❌ Error extracting Grand Total or renaming sheet: {e}")

    break  # Stop after first PivotTable
