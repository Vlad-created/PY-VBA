import xlwings as xw
import os

def check_text_format_in_open_workbooks():
    print("🔍 Checking column formats in open workbooks...")
    for wb in xw.apps.active.books:
        try:
            ws = wb.sheets["Sheet1"]
            for col in ["A", "F", "Y", "AS"]:
                fmt = ws.range(f"{col}1").number_format
                if fmt == "@":
                    print(f"✅ Column {col} in '{wb.name}' is set to Text.")
                else:
                    print(f"❌ Column {col} in '{wb.name}' is NOT set to Text (found: {fmt}).")
        except Exception as e:
            print(f"⚠️ Error checking workbook '{wb.name}': {e}")

def check_text_format_in_data_file():
    print("🔍 Checking column formats in 'data.xlsx'...")
    path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\data.xlsx"
    if not os.path.exists(path):
        print("❌ File not found:", path)
        return
    try: 
        wb = xw.Book(path)
        ws = wb.sheets["Detail1"]
        for col in ["A", "F", "Y", "AS"]:
            fmt = ws.range(f"{col}1").number_format
            if fmt == "@":
                print(f"✅ Column {col} in '{wb.name}' is set to Text.")
            else:
                print(f"❌ Column {col} in '{wb.name}' is NOT set to Text (found: {fmt}).")
        wb.close()
    except Exception as e:
        print(f"⚠️ Error checking workbook '{path}': {e}")

def check_formulas_in_data_file():
    print("\n🔍 Checking formulas in 'data.xlsx'...")
    path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\data.xlsx"
    if not os.path.exists(path):
        print("❌ File not found:", path)
        return
    try:
        wb = xw.Book(path)
        ws = wb.sheets["Detail1"]

        cp_formula = ws.range("CP2").formula
        cq_formula = ws.range("CQ2").formula

        expected_cp = '=IFNA(XLOOKUP(CE2,Maping!B:B,Maping!C:C)," ")'
        expected_cq = '=IF(AND(CD2="Refused",CG2="Sales",CP2="поради отказ на клиента от услуга"),1,0)'

        print("✅ CP2 formula correct." if cp_formula == expected_cp else f"❌ CP2 formula mismatch: {cp_formula}")
        print("✅ CQ2 formula correct." if cq_formula == expected_cq else f"❌ CQ2 formula mismatch: {cq_formula}")

        wb.close()
    except Exception as e:
        print(f"⚠️ Error checking formulas in workbook '{path}': {e}")

# 🚀 Run all checks
check_text_format_in_open_workbooks()
check_text_format_in_data_file()
check_formulas_in_data_file()
