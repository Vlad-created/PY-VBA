import xlwings as xw

# Master file
file_All_Regions = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_All Regions.xlsx"

# Regional files
regional_files = {
    "Бургас": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_Бургас.xlsx",
    "Варна 1": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_Варна 1.xlsx",
    "Варна 2": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_Варна 2.xlsx",
    "Велико Търново": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_Велико Търново.xlsx",
    "Стара Загора": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_Стара Загора.xlsx",
    "Пловдив 1": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_Пловдив 1.xlsx",
    "Пловдив 2": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_Пловдив 2.xlsx",
    "Пловдив 3": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_Пловдив 3.xlsx",
    "София 1": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_София 1.xlsx",
    "София 2": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_София 2.xlsx",
    "София 3": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_София 3.xlsx",
    "София 4": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_София 4.xlsx",
    "Северозапад": r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_Северозапад.xlsx"
}

columns = ["S", "T", "U", "V"]
start_row = 14
end_row = 27
tolerance = 0.05

# Open master workbook
wb_master = xw.Book(file_All_Regions)
ws_master = wb_master.sheets["By Shop"]

# Read F14:F27 from master
master_values = {row: ws_master.range(f"F{row}").value for row in range(start_row, end_row + 1)}

# Loop through regional files
for region, path in regional_files.items():
    wb_region = xw.Book(path)
    ws_region = wb_region.sheets["By Shop"]
    region_id = ws_region.range("F11").value

    print(f"\n📍 Region: {region}")
    print(f"🔎 F11 value: {region_id}")

    # Find matching row in master
    match_row = None
    for row, value in master_values.items():
        if value == region_id:
            match_row = row
            break

    if match_row:
        print(f"✅ Match found in All Regions at row {match_row}")
        for col in columns:
            master_val = ws_master.range(f"{col}{match_row}").value
            region_val = ws_region.range(f"{col}11").value

            if master_val is None or region_val is None:
                result = "⚠️ Missing data"
            else:
                try:
                    result = abs(master_val - region_val) <= tolerance
                except TypeError:
                    result = "⚠️ Non-numeric"

            print(f"🔸 {col}: {result} (Master: {master_val}, Region: {region_val})")
    else:
        print("❌ No match found in All Regions for F11 value.")

    wb_region.close()

wb_master.close()
