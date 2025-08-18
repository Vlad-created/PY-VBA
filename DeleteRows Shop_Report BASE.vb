Sub DeleteColumns()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

Application.ScreenUpdating = False 

ws.Colums = ("BV:CB").Delete Shift: = xltoLeft
ws.Colums = ("BP:BT").Delete Shift: = xltoLeft
ws.Colums = ("AQ:BM").Delete Shiftt: = xltoLeft
ws.Colums = ("AB:AO"). Delete Shift: = xltoLeft
ws.Colums = ("Z").Delete Shift: = xltoLeft
ws.Colums = ("R:W").Delete Shift: = xltoLeft
ws.Colums = ("K:O").Delete Shift: = xltoLeft
ws.Colums = ("A:E").Delete Shift: = xltoLeft

Application.ScreenUpdating = True

End Sub
