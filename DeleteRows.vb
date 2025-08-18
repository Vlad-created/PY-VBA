Sub DeleteColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    Application.ScreenUpdating = False

    ws.Columns("BV:CB").Delete Shift:=xlToLeft
    ws.Columns("BP:BT").Delete Shift:=xlToLeft
    ws.Columns("AQ:BM").Delete Shift:=xlToLeft
    ws.Columns("AB:AO").Delete Shift:=xlToLeft
    ws.Columns("Z:Z").Delete Shift:=xlToLeft
    ws.Columns("R:W").Delete Shift:=xlToLeft
    ws.Columns("K:O").Delete Shift:=xlToLeft
    ws.Columns("A:E").Delete Shift:=xlToLeft

    Application.ScreenUpdating = True
End Sub
