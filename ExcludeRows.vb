Sub ExcludeGermanyRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

     Set ws = ThisWorkbook.Sheets("202505")
     lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

     For i = lastRow To 2 Step -1 
     If Trim(UCase(ws.Cells(i, "F").Value)) = "GERMANY" Then
     ws.Rows(i).EntireRow.Hidden = True
     End If
     Next i

    MsgBox "Ohhh! Again you! Rows with 'GERMANY' in column F have been hidden."
End Sub

Sub UnhideGermanyRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

      Set ws = ThisWorkbook.Sheets("202505") 
      lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

     For i = 2 To lastRow 
     If Trim(UCase(ws.Cells(i, "F").Value)) = "GERMANY" Then
     ws.Rows(i).EntireRow.Hidden = False
     End If
     Next i

     MsgBox "Again you? Rows with 'GERMANY' in column F have been unhidden."
End Sub