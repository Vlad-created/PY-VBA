import xlwings as xw
from decimal import Decimal, InvalidOperation

# Loop through all open workbooks
for wb in xw.apps.active.books:
    try:
        ws = wb.sheets["Sheet1"]

        for col in ["A", "F", "Y", "AS"]:
            # Get last used row in the column
            last_row = ws.range(f"{col}{ws.cells.last_cell.row}").end("up").row

            # Read entire column range
            rng = ws.range(f"{col}1:{col}{last_row}")
            values = rng.value

            # Ensure values are in list form
            if not isinstance(values, list):
                values = [[values]]

            # Convert all values to precise strings
            converted = []
            for v in values:
                val = v[0] if isinstance(v, list) else v
                if val is None:
                    converted.append([""])
                else:
                    try:
                        # Use Decimal to preserve full precision
                        dec = Decimal(str(val))
                        # Remove trailing .0 if it's an integer
                        text = str(dec.normalize()) if dec == dec.to_integral() else str(dec)
                        converted.append([text])
                    except (InvalidOperation, ValueError):
                        # If not a number, just convert to string
                        converted.append([str(val)])

            # Set column format to Text
            ws.range(f"{col}:{col}").number_format = "@"

            # Write back all values at once
            rng.value = converted

        print(f"✅ Converted to text without data loss: {wb.name}")

    except Exception as e:
        print(f"⚠️ Error in {wb.name}: {e}")
