Sub AddNewColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RRi List ")
    
    Dim i As Integer
    For i = 1 To 3
        ws.Columns("AH:AH").Insert Shift:=xlToRight
    Next i
    
    ws.Range("AH1").Value = "Status Type"
    ws.Range("AI1").Value = "Fin 4"
    ws.Range("AJ1").Value = "Fin 2"
    

    ws.Range("AG:AG").Copy
    ws.Range("AH:AJ").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False


       ws.Range("AH2").Formula = "=IF(F2 =""RRI"",""Applicable"",""Not Applicable"")"
    
ws.Range("AH2:AH" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).FillDown

ws.Range("AH:AH").Value = ws.Range("AH:AH").Value

    End Sub




= Table.SelectRows(Source, each Text.Contains([Folder Path], "_HRSS HC Dash") and not Text.Contains([Folder Path], "Archive") and [Name] = "Headcount "&MonthValue&".xlsx")

Landing HC Column = IF(Table3[HRSSC]<>"Onsite"&&Table3[Intern/Permanent]="Permanent",IF(Table3[Dec-24]="HC"||Table3[Dec-24]="ML IN"||Table3[Dec-24]="ML OUT"||Table3[Dec-24]="IN"||Table3[Dec-24]="OUT",1,0))