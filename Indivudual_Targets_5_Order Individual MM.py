import xlwings as xw

# Path to your workbook
file_path = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets\Individual January'26.xlsm"

# Launch Excel without Book1
app = xw.App(visible=True, add_book=False)
wb = app.books.open(file_path)
ws = wb.sheets["All_FTEs"]

# Find last row in column B
last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

# Clear previous sort fields
ws.api.Sort.SortFields.Clear()

# Add sort fields (equivalent to VBA)
ws.api.Sort.SortFields.Add(
    Key=ws.range(f"B2:B{last_row}").api,
    SortOn=0,  # xlSortOnValues
    Order=2,   # xlDescending
    DataOption=0
)

ws.api.Sort.SortFields.Add(
    Key=ws.range(f"C2:C{last_row}").api,
    SortOn=0,
    Order=1,   # xlAscending
    DataOption=0
)

ws.api.Sort.SortFields.Add(
    Key=ws.range(f"E2:E{last_row}").api,
    SortOn=0,
    Order=1,
    DataOption=0
)

ws.api.Sort.SortFields.Add(
    Key=ws.range(f"G2:G{last_row}").api,
    SortOn=0,
    Order=1,
    DataOption=0
)

# Apply sort to full range A1:Z[last_row]
ws.api.Sort.SetRange(ws.range(f"A1:Z{last_row}").api)
ws.api.Sort.Header = 1  # xlYes
ws.api.Sort.MatchCase = False
ws.api.Sort.Orientation = 1  # xlTopToBottom
ws.api.Sort.Apply()

print("✅ Sorting completed successfully.")
