import time
import win32com.client as win32

def refresh_workbook(path):
    print(f"\nOpening workbook: {path}")

    wb = excel.Workbooks.Open(path)

    # --- DELETE ALL POWER QUERIES ---
    try:
        while wb.Queries.Count > 0:
            wb.Queries.Item(1).Delete()
        print("All Power Queries deleted.")
    except Exception as e:
        print(f"Error deleting Power Queries: {e}")

    # --- DELETE ALL WORKBOOK CONNECTIONS ---
    try:
        while wb.Connections.Count > 0:
            wb.Connections.Item(1).Delete()
        print("All Workbook Connections deleted.")
    except Exception as e:
        print(f"Error deleting Connections: {e}")

    # --- DELETE ALL QUERYTABLES (PQ loaded to table) ---
    try:
        for sheet in wb.Worksheets:
            qt_count = sheet.QueryTables.Count
            if qt_count > 0:
                print(f"Deleting {qt_count} QueryTables in sheet {sheet.Name}")
                for i in range(qt_count, 0, -1):
                    sheet.QueryTables.Item(i).Delete()
    except Exception as e:
        print(f"Error deleting QueryTables: {e}")

    # --- DELETE ALL PIVOT CACHES (PQ loaded to Pivot) ---
    try:
        while wb.PivotCaches().Count > 0:
            wb.PivotCaches().Item(1).Delete()
        print("All PivotCaches deleted.")
    except Exception as e:
        print(f"Error deleting PivotCaches: {e}")

    # --- DELETE DATA MODEL CONTENT ---
    try:
        model = wb.Model

        # Delete relationships
        while model.ModelRelationships.Count > 0:
            model.ModelRelationships.Item(1).Delete()

        # Delete tables
        while model.ModelTables.Count > 0:
            model.ModelTables.Item(1).Delete()

        print("Data Model cleared (tables + relationships).")
    except Exception as e:
        print(f"Error clearing Data Model: {e}")

    # Save and close
    wb.Save()
    wb.Close()

    print("Workbook fully cleaned and saved.")


# Start Excel once
excel = win32.Dispatch("Excel.Application")
excel.Visible = False

# File paths
file_Append_Targets = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Archive Total NBS\Targets_shops_TOTAL_NBS_2026\Append_Targets.xlsx"
file_Targets   = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Archive Total NBS\Targets_shops_TOTAL_NBS_2026\Targets shops march 2018 za nalivane.xlsx"
file_Targets_total_NBS = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Archive Total NBS\Targets_shops_TOTAL_NBS_2026\Targets_shops_TOTAL_NBS.xlsx"

# Process all workbooks
refresh_workbook(file_Append_Targets)
refresh_workbook(file_Targets)
refresh_workbook(file_Targets_total_NBS)

excel.Quit()

print("\nAll workbooks fully cleaned and closed.")
