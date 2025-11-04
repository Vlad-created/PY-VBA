import xlwings as xw

# Loop through all open workbooks
for wb in xw.apps.active.books:
    try:
        ws = wb.sheets["Sheet1"]

        for col in ["A", "F", "Y", "AS"]:
            # Set column format to Text BEFORE reading/writing
            ws.range(f"{col}:{col}").number_format = "@"

            # Get last used row in the column
            last_row = ws.range(f"{col}{ws.cells.last_cell.row}").end("up").row

            # Read entire column range
            rng = ws.range(f"{col}1:{col}{last_row}")
            values = rng.value

            # Ensure values are in list form
            if not isinstance(values, list):
                values = [[values]]

            # Convert all values to strings
            converted = []
            for v in values:
                val = v[0] if isinstance(v, list) else v
                if val is None:
                    converted.append([""])
                else:
                    # Convert to string without scientific notation
                    text = str(val)
                    if isinstance(val, (int, float)):
                        text = f"{val:.0f}" if float(val).is_integer() else f"{val}"
                    converted.append([text])

            # Write back all values at once
            rng.value = converted

        print(f"✅ Converted to text without scientific notation: {wb.name}")

    except Exception as e:
        print(f"⚠️ Error in {wb.name}: {e}")
