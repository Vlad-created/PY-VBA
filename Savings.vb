Sub AddFormula()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Savings 2025")
    
   

    ws.Range("L2").Formula = "=[@[Nessis Salary]]*VLOOKUP(K2,Sheet2!$I$1:$J$3,2,FALSE)"
    ws.Range("Z2").Formula = "=IF(YEAR([@[EFFECTIVE DATE]])>=2026, 0, IFERROR(DAYS($AO$1, [@[EFFECTIVE DATE]]) / 365 * [@[ANNUALISED SAVINGS ]], 0))"
    ws.Range("AA2").Formula = "=IFERROR([@[Annual Salary Cost]]-[@[Total Salary Cost ]],0)"
    ws.Range("AB2").Formula = "=([@[SAVINGS Rest of Yr]])-[@[Paid Exit]]"

ws.Range("L2:AB" & ws.Cells(ws.Rows.Count, "AE").End(xlUp).Row).FillDown

ws.Range("L:AB").Value = ws.Range("L:AB").Value


End Sub
