import xlwings as xw

# Loop through all open workbooks
for wb in xw.apps.active.books:
    try:
        ws = wb.sheets["Sheet1"]

        for col in ["A", "F", "Y", "AS"]:
            cell_ref = f"{col}1"

            # Apply TextToColumns using Excel COM API
            ws.range(cell_ref).api.TextToColumns(
                Destination=ws.range(cell_ref).api,
                DataType=1,  # xlDelimited
                TextQualifier=1,  # xlTextQualifierDoubleQuote
                ConsecutiveDelimiter=False,
                Tab=True,
                Semicolon=False,
                Comma=False,
                Space=False,
                Other=False,
                FieldInfo=[[1, 2]],  # Column 1 as Text
                TrailingMinusNumbers=True
            )

            # Get last used row in the column
            last_row = ws.range(f"{col}{ws.cells.last_cell.row}").end("up").row

            # Convert each cell to string explicitly
            for row in range(1, last_row + 1):
                cell = ws.range(f"{col}{row}")
                value = cell.value
                if value is not None:
                    cell.value = str(value)

        print(f"✅ Converted to text: {wb.name}")

    except Exception as e:
        print(f"⚠️ Error in {wb.name}: {e}")
