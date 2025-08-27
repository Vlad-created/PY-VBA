import xlwings as xw

# Loop through all open workbooks
for wb in xw.apps.active.books:
    try:
        ws = wb.sheets["Sheet1"]

        for col in ["A", "F", "Y", "AS"]:
            # Set column format to Text
            ws.range(f"{col}:{col}").number_format = "@"

            # Get last used row in the column
            last_row = ws.range(f"{col}{ws.cells.last_cell.row}").end("up").row

            # Convert each cell to a string with a leading apostrophe
            for row in range(1, last_row + 1):
                cell = ws.range(f"{col}{row}")
                value = cell.value
                if value is not None:
                    # Force Excel to treat it as text by adding apostrophe
                    cell.value = f"'{value}"

        print(f"✅ Converted to text: {wb.name}")

    except Exception as e:
        print(f"⚠️ Error in {wb.name}: {e}")
