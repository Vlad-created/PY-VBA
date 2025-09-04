import xlwings as xw

# Open the workbook
wb = xw.Book(r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\data.xlsx")
ws = wb.sheets["Detail1"]

# Insert two columns at CO
for _ in range(2):
    ws.range("CO1").entire_column.insert()

# Add headers
ws.range("CP1").value = "Refused_Reasons_Grouped2"
ws.range("CQ1").value = "custRefused"

# Add formulas
ws.range("CP2").formula = '=IFNA(XLOOKUP(CE2,Maping!B:B,Maping!C:C)," ")'
ws.range("CQ2").formula = '=IF(AND(CD2="Refused",CG2="Sales",CP2="поради отказ на клиента от услуга"),1,0)'

# Fill down formulas
last_row = ws.range("A" + str(ws.cells.last_cell.row)).end("up").row
ws.range(f"CP2:CQ{last_row}").api.FillDown()

# Convert formulas to values
ws.range(f"CP2:CQ{last_row}").value = ws.range(f"CP2:CQ{last_row}").value
