import xlwings as xw
import os

# Define file paths
file_new = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\FBB_Cancellation Rate_28.09_All Regions.xlsx"
file_old = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\Archive\FBB_Cancellation Rate_24.09_All Regions.xlsx"

# Define range to compare
columns = ["T", "U", "V"]
start_row = 14
end_row = 27
tolerance = 0.15  # 15%

# Open both workbooks
wb_new = xw.Book(file_new)
wb_old = xw.Book(file_old)

ws_new = wb_new.sheets["By Shop"]
ws_old = wb_old.sheets["By Shop"]

print("🔍 Comparing values with 15% tolerance...\n")

# Loop through each cell in the defined range
for row in range(start_row, end_row + 1):
    for col in columns:
        cell_new = ws_new.range(f"{col}{row}").value
        cell_old = ws_old.range(f"{col}{row}").value

        # Handle None or empty cells
        if cell_new is None or cell_old is None:
            result = "⚠️ Missing data"
        else:
            try:
                # Compare with tolerance
                diff = abs(cell_new - cell_old)
                result = diff <= tolerance
            except TypeError:
                result = "⚠️ Non-numeric"

        print(f"{col}{row}: {result} (New: {cell_new}, Old: {cell_old})")

# Close workbooks
wb_new.close()
wb_old.close()

print("All cells In All Regions checked")