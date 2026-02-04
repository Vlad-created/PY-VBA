import win32com.client as win32
import os
import shutil

excel = win32.Dispatch("Excel.Application")
excel.Visible = True

# Base file path
base_path = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Krachev_ Crak\NBS_Target_South.xlsx"

# Output folder
output_folder = r"\\virgo\Work\OpControl\2016\Regular_Daily_Reports\Georgi Tomov\Execution\Targets_RM\Krachev_ Crak"

# Region groups: values we want to KEEP in each file
region_groups = {
    "Пловдив 1": ["Пловдив 1", "Регион Пловдив 1", "Total"],
    "Пловдив 2": ["Пловдив 2", "Регион Пловдив 2", "Total"],
    "Пловдив 3": ["Пловдив 3", "Регион Пловдив 3", "Total"],
    "Стара Загора": ["Стара Загора", "Регион Стара Загора", "Total"]
}

# Sheets to process
sheets_to_process = [
    "Opportunity",
    "Upsell Revenue RSDN",
    "Voice",
    "Fix",
    "Devices",
    "Retention Stack",
    "Other Stacks",
    "Target per FTE",
    "Bonus&Checks"
]

extra_sheet = "All targets&Daily"  # untouched

xlUp = -4162

for region_name, keep_values in region_groups.items():
    print(f"\n=== Creating file for region: {region_name} ===")

    # Copy base file
    region_file = os.path.join(output_folder, f"{region_name}.xlsx")
    shutil.copy2(base_path, region_file)

    # Open copied workbook
    wb = excel.Workbooks.Open(region_file)

    for sheet_name in sheets_to_process:
        print(f"Cleaning sheet: {sheet_name}")
        ws = wb.Sheets(sheet_name)

        # Find last used row in column E
        last_row = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        if last_row <= 20:
            continue  # nothing below header

        # Loop from bottom to row 21 and delete rows not matching region group
        for row in range(last_row, 20, -1):
            val = ws.Cells(row, "E").Value
            if val is None:
                # Empty region cell -> delete row
                ws.Rows(row).Delete()
            else:
                sval = str(val).strip()
                if sval not in keep_values:
                    ws.Rows(row).Delete()

        # Ensure everything below 21 is visible
        ws.Rows("21:" + str(ws.Rows.Count)).EntireRow.Hidden = False

    # Save and close
    wb.Save()
    wb.Close()

    print(f"Saved: {region_file}")

print("\n✅ All region files created successfully.")
