import win32com.client as win32
import os
import shutil

excel = win32.Dispatch("Excel.Application")
excel.Visible = True

# Base file path
base_path = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Krachev_ Crak\NBS_Target_South.xlsm"

# Output folder
output_folder = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Krachev_ Crak"

# Region groups
region_groups = {
    "Пловдив 1": ["Пловдив 1", "Регион Пловдив 1"],
    "Пловдив 2": ["Пловдив 2", "Регион Пловдив 2"],
    "Пловдив 3": ["Пловдив 3", "Регион Пловдив 3"],
    "Стара Загора": ["Стара Загора", "Регион Стара Загора"]
}

# Only "Total" should be copied as values
copy_targets = ["Total"]

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

    region_file = os.path.join(output_folder, f"{region_name}.xlsm")
    shutil.copy2(base_path, region_file)

    wb = excel.Workbooks.Open(region_file)

    # ------------------------------------------------------------
    # STEP 1: COPY ONLY "Total" ROWS AS VALUES
    # ------------------------------------------------------------
    print("STEP 1: Copy 'Total' rows as values")
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

            if str(val).strip() == "Total":
                ws.Rows(row).Copy()
                ws.Rows(row).PasteSpecial(Paste=xlPasteValues)

        excel.CutCopyMode = False

    # ------------------------------------------------------------
    # STEP 2: HIDE ROWS WHERE E = "Total" (NOT DELETE)
    # ------------------------------------------------------------
    print("STEP 2: Hide rows where column E = 'Total'")
    for sheet_name in sheets_to_process:
        ws = wb.Sheets(sheet_name)
        print(f"  [Step 2] Sheet: {sheet_name}")

        last_row = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        if last_row <= 20:
            continue

        for row in range(21, last_row + 1):
            val = ws.Cells(row, "E").Value
            if val is None:
                continue

            if str(val).strip() == "Total":
                ws.Rows(row).EntireRow.Hidden = True

    # ------------------------------------------------------------
    # STEP 3: DELETE ROWS NOT IN keep_values
    # ------------------------------------------------------------
    print("STEP 3: Delete rows NOT in region keep_values")
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
            if sval not in keep_values and sval != "Total":
                ws.Rows(row).Delete()

    wb.Save()
    wb.Close()

    print(f"Saved: {region_file}")

print("\n✅ All region files created successfully.")