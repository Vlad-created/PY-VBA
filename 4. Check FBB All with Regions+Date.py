import xlwings as xw
from datetime import datetime, timedelta

# Get yesterday's date
today = datetime.today()
file_date = today - timedelta(days=1)
date_str = file_date.strftime("%d.%m")

# Build dynamic path for All Regions file
base_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation"
file_name = f"FBB_Cancellation Rate_{date_str}_All Regions.xlsx"
file_All_Regions = f"{base_path}\\{file_name}"

print("Dynamic file path:", file_All_Regions)

# Regional files
regional_files = {
    "Бургас": f"{base_path}\\FBB_Cancellation Rate_{date_str}_Бургас.xlsx",
    "Варна 1": f"{base_path}\\FBB_Cancellation Rate_{date_str}_Варна 1.xlsx",
    "Варна 2": f"{base_path}\\FBB_Cancellation Rate_{date_str}_Варна 2.xlsx",
    "Велико Търново": f"{base_path}\\FBB_Cancellation Rate_{date_str}_Велико Търново.xlsx",
    "Стара Загора": f"{base_path}\\FBB_Cancellation Rate_{date_str}_Стара Загора.xlsx",
    "Пловдив 1": f"{base_path}\\FBB_Cancellation Rate_{date_str}_Пловдив 1.xlsx",
    "Пловдив 2": f"{base_path}\\FBB_Cancellation Rate_{date_str}_Пловдив 2.xlsx",
    "Пловдив 3": f"{base_path}\\FBB_Cancellation Rate_{date_str}_Пловдив 3.xlsx",
    "София 1": f"{base_path}\\FBB_Cancellation Rate_{date_str}_София 1.xlsx",
    "София 2": f"{base_path}\\FBB_Cancellation Rate_{date_str}_София 2.xlsx",
    "София 3": f"{base_path}\\FBB_Cancellation Rate_{date_str}_София 3.xlsx",
    "София 4": f"{base_path}\\FBB_Cancellation Rate_{date_str}_София 4.xlsx",
    "Северозапад": f"{base_path}\\FBB_Cancellation Rate_{date_str}_Северозапад.xlsx"
}

columns = ["Y", "Z", "AA", "AB", "AC", "AD"]
start_row = 14
end_row = 27
tolerance = 0.15

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

    # 🔗 Match region ID with master
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
