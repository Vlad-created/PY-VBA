import xlwings as xw

file_path = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets\Individual February'26.xlsm"

app = xw.App(visible=True, add_book=False)
wb = app.books.open(file_path)
ws = wb.sheets["All_FTEs"]

# -----------------------------
# 1. FIND LAST ROW
# -----------------------------
last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

# -----------------------------
# 2. SORTING (same as VBA)
# -----------------------------
sort = ws.api.Sort
sort.SortFields.Clear()

sort.SortFields.Add(
    Key=ws.range(f"B2:B{last_row}").api,
    SortOn=0,  # xlSortOnValues
    Order=2,   # xlDescending
    DataOption=0
)

sort.SortFields.Add(
    Key=ws.range(f"C2:C{last_row}").api,
    SortOn=0,
    Order=1,   # xlAscending
    DataOption=0
)

sort.SortFields.Add(
    Key=ws.range(f"E2:E{last_row}").api,
    SortOn=0,
    Order=1,
    DataOption=0
)

sort.SortFields.Add(
    Key=ws.range(f"G2:G{last_row}").api,
    SortOn=0,
    Order=1,
    DataOption=0
)

sort.SetRange(ws.range(f"A1:Z{last_row}").api)
sort.Header = 1
sort.MatchCase = False
sort.Orientation = 1
sort.Apply()

print("✔ Sorting completed")

# -----------------------------
# 3. UPDATE FORMULAS (Unicode OK)
# -----------------------------

for i in range(2, last_row + 1):
    cell = ws.range(f"J{i}")

    # Only update if the cell already contains a formula
    if cell.api.HasFormula:
        formula = (
            f'=IF(AND(G{i}="Мениджър магазин", K{i}=""), '
            f'IF(COUNTIF(D:D, D{i})<=2, 1, '
            f'IF(COUNTIF(D:D, D{i})<=4, 0.5, '
            f'IF(COUNTIF(D:D, D{i})>=5, 0))), '
            f'IF(AND(K{i}="", I{i}=""), 1, '
            f'IF(AND(I{i}<>"", K{i}<>""), VLOOKUP(K{i}, mapping!$B$3:$C$10, 2, FALSE), '
            f'IF(AND(I{i}="", K{i}<>""), VLOOKUP(K{i}, mapping!$B$3:$C$10, 2, FALSE), '
            f'IF(AND(I{i}<>"", K{i}=""), VLOOKUP(I{i}, mapping!$G$3:$H$8, 2, FALSE))))))'
        )

        cell.formula = formula

print("✔ Formulas updated")

# -----------------------------
# 4. SAVE WORKBOOK
# -----------------------------
wb.save()

print("✔ File saved successfully")
