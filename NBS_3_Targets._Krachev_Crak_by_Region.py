import win32com.client as win32
import os
import shutil

excel = win32.Dispatch("Excel.Application")
excel.Visible = True

# Base file path
base_path = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Krachev_ Crak\NBS_Target_South.xlsx"

# Output folder
output_folder = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Krachev_ Crak"

# Region groups: "Total" removed as requested
region_groups = {
    "Пловдив 1": ["Пловдив 1", "Регион Пловдив 1"],
    "Пловдив 2": ["Пловдив 2", "Регион Пловдив 2"],
    "Пловдив 3": ["Пловдив 3", "Регион Пловдив 3"],
    "Стара Загора": ["Стара Загора", "Регион Стара Загора"]
}

# Values to COPY as values before deletion
copy_targets = [
    "Регион Пловдив 1",
    "Регион Пловдив 2",
    "Регион Пловдив 3",
    "Регион Стара Загора",
    "Total"
]

# Sheets to process
sheets_to_process = [
    "Opportunity",
    "Upsell Revenue RSDN",
    "Voice",
    "Fix",
    "Devices",
    "Other Stacks",
    "Target per FTE",
    "Bonus&Checks",
    "Retention Stack"
]

xlUp = -4162
xlPasteValues = -4163

for region_name, keep_values in region_groups.items():
    print(f"\n=== Creating file for region: {region_name} ===")

    region_file = os.path.join(output_folder, f"{region_name}.xlsx")
    shutil.copy2(base_path, region_file)

    wb = excel.Workbooks.Open(region_file)

    # ------------------------------------------------------------
    # STEP 1: IN ALL SHEETS, COPY ROWS (E in copy_targets) AS VALUES
    # ------------------------------------------------------------
    print("STEP 1: Copy rows as values for copy_targets in ALL sheets")
    for sheet_name in sheets_to_process:
        ws = wb.Sheets(sheet_name)
        print(f"  [Step 1] Sheet: {sheet_name}")

        last_row = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        if last_row <= 20:
            continue

        for row in range(21, last_row + 1):
            val = ws.Cells(row, "E").Value
            if val is None:
                continue

            sval = str(val).strip()
            if sval in copy_targets:
                ws.Rows(row).Copy()
                ws.Rows(row).PasteSpecial(Paste=xlPasteValues)

        excel.CutCopyMode = False

    # ------------------------------------------------------------
    # STEP 2: IN ALL SHEETS, DELETE ROWS WHERE E = "Total"
    # ------------------------------------------------------------
    print("STEP 2: Delete rows where column E = 'Total' in ALL sheets")
    for sheet_name in sheets_to_process:
        ws = wb.Sheets(sheet_name)
        print(f"  [Step 2] Sheet: {sheet_name}")

        last_row = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        if last_row <= 20:
            continue

        for row in range(last_row, 20, -1):
            val = ws.Cells(row, "E").Value
            if val is None:
                continue

            if str(val).strip() == "Total":
                ws.Rows(row).Delete()

    # ------------------------------------------------------------
    # STEP 3: IN ALL SHEETS, DELETE ROWS NOT IN keep_values
    # ------------------------------------------------------------
    print("STEP 3: Delete rows NOT in region keep_values in ALL sheets")
    for sheet_name in sheets_to_process:
        ws = wb.Sheets(sheet_name)
        print(f"  [Step 3] Sheet: {sheet_name}")

        last_row = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        if last_row <= 20:
            continue

        for row in range(last_row, 20, -1):
            val = ws.Cells(row, "E").Value

            if val is None:
                ws.Rows(row).Delete()
                continue

            sval = str(val).strip()
            if sval not in keep_values:
                ws.Rows(row).Delete()

        ws.Rows("21:" + str(ws.Rows.Count)).EntireRow.Hidden = False

    wb.Save()
    wb.Close()

    print(f"Saved: {region_file}")

print("\n✅ All region files created successfully with 3-step logic applied.")
