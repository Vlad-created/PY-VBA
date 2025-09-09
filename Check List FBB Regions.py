import xlwings as xw
import os

def check_format_in_by_shop():
    print("\n🔍 Checking format in 'FBB_Cancellation Rate_03.09_All Regions.xlsx'...")
    path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_08.09_All Regions.xlsx"
    if not os.path.exists(path):
        print("❌ File not found:", path)
        return

    wb = xw.Book(path)

    # Dictionary of sheets and their expected formats
    sheet_format_map = {
        "By Shop": {
            "T": "General",  # Number
            "U": "0%",       # Percentage
            "V": "0%"        # Percentage
        },
        "By Reason and Region": {
            "F": "0%",
            "H": "0%",
            "J": "0%",
            "L": "0%"
        },
        "By Reason": {
            "G": "0%",
            "I": "0%",
            "K": "0%",
            "M": "0%"
        },
        "Sent to Shop": {
            "G": "0%",
            "K": "0%",
            "M": "0%",
            "Q": "0%",
            "S": "0%",
            "W": "0%",
            "Y": "0%",
            "AC": "0%"
        }
    }

    # Loop through each sheet and check formats
    for sheet_name, expected_formats in sheet_format_map.items():
        print(f"\n📄 Checking sheet: '{sheet_name}'")
        ws = wb.sheets[sheet_name]
        for col, expected in expected_formats.items():
            actual = ws.range(f"{col}1").number_format
            if expected in actual:
                print(f"✅ Column {col} format is correct ({actual}).")
            else:
                print(f"❌ Column {col} format mismatch. Expected '{expected}', found '{actual}'.")

    wb.close()



# Run all checks
check_format_in_by_shop()
