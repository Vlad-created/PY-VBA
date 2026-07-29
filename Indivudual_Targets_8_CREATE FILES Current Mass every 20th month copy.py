import os
import calendar
from datetime import datetime
import shutil
import time
import win32com.client as win32

# ---------------------------------------------------
# 1. Build dynamic file path
# ---------------------------------------------------
base_path = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets\Regions"

current_month = calendar.month_name[datetime.now().month]
file_name = f"Individual_targets_{current_month}_Total.xlsx"
file_path = os.path.join(base_path, file_name)

print("Main file:", file_path)

# ---------------------------------------------------
# 2. Create a copy (with fallback if original is locked)
# ---------------------------------------------------
base, ext = os.path.splitext(file_path)
copy_path = f"{base}_working{ext}"

try:
    shutil.copy2(file_path, copy_path)
    print("Created working copy with shutil.copy2:", copy_path)
except PermissionError:
    print("shutil.copy2 failed (file locked). Trying Excel SaveCopyAs fallback...")
    excel_tmp = None
    try:
        excel_tmp = win32.DispatchEx("Excel.Application")
        excel_tmp.Visible = False
        excel_tmp.DisplayAlerts = False
        wb_orig = excel_tmp.Workbooks.Open(file_path, ReadOnly=True)
        time.sleep(0.3)
        wb_orig.SaveCopyAs(copy_path)
        wb_orig.Close(SaveChanges=False)
        print("Created working copy with Excel SaveCopyAs:", copy_path)
    except Exception as e:
        print("Excel fallback failed:", e)
        raise
    finally:
        if excel_tmp:
            try:
                excel_tmp.Quit()
            except Exception:
                pass

# ---------------------------------------------------
# 3. Start Excel and open working copy
# ---------------------------------------------------
excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(copy_path)
    ws = wb.Worksheets("individual")

    # ---------------------------------------------------
    # 4. Unhide ALL columns in sheet "individual"
    # ---------------------------------------------------
    ws.Columns.Hidden = False

    # ---------------------------------------------------
    # 5. Copy all and paste as VALUES
    # ---------------------------------------------------
    xlPasteValues = -4163
    used_range = ws.UsedRange
    used_range.Copy()
    ws.Range("A1").PasteSpecial(Paste=xlPasteValues)
    excel.CutCopyMode = False

    # ---------------------------------------------------
    # 6. Delete columns DY → EK (as in original)
    # ---------------------------------------------------
    try:
        ws.Range("DY:EK").Delete()
    except Exception:
        # ignore if range doesn't exist or other issue
        pass

    # ---------------------------------------------------
    # 7. Delete all other sheets except "individual"
    # ---------------------------------------------------
    for sh in list(wb.Worksheets):
        if sh.Name != "individual":
            sh.Delete()

    wb.Save()

    # ---------------------------------------------------
    # 8. Filter by Region and create region files
    # ---------------------------------------------------
    regions = [
        "Северозапад", "София 1", "София 2", "София 3", "София 4",
        "Бургас", "Варна 1", "Варна 2", "Велико Търново",
        "Пловдив 1", "Пловдив 2", "Пловдив 3", "Стара Загора"
    ]

    xlUp = -4162
    xlToLeft = -4159
    xlCellTypeVisible = 12

    last_row = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    last_col = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    data_range = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col))

    for region in regions:

        if ws.AutoFilterMode:
            ws.AutoFilterMode = False

        data_range.AutoFilter(Field=5, Criteria1=region)

        try:
            visible = data_range.SpecialCells(xlCellTypeVisible)
        except Exception:
            print(f"No rows for region: {region}")
            continue

        new_wb = excel.Workbooks.Add()
        new_ws = new_wb.Worksheets(1)
        new_ws.Name = "individual"

        # Copy visible (including header) to new workbook
        visible.Copy(new_ws.Range("A1"))

        # ---------------------------------------------------
        # Delete columns F, C and A in the region workbook
        # Delete in descending order to avoid shifting issues
        # ---------------------------------------------------
        try:
            new_ws.Range("F:F").Delete()
        except Exception:
            pass
        try:
            new_ws.Range("C:C").Delete()
        except Exception:
            pass
        try:
            new_ws.Range("A:A").Delete()
        except Exception:
            pass

        # ---------------------------------------------------
        # Auto-fit columns so the sheet looks visually correct
        # ---------------------------------------------------
        try:
            new_ws.Columns.AutoFit()
        except Exception:
            pass

        # Save region file
        safe_region_name = region.replace("/", "_").replace("\\", "_")
        region_file = f"{base}_{safe_region_name}{ext}"
        new_wb.SaveAs(region_file)
        new_wb.Close(SaveChanges=False)

        print("Created:", region_file)

    if ws.AutoFilterMode:
        ws.AutoFilterMode = False

    wb.Save()

finally:
    # Ensure Excel is closed cleanly
    try:
        wb.Close(SaveChanges=True)
    except Exception:
        pass
    try:
        excel.Quit()
    except Exception:
        pass

print("All done.")
