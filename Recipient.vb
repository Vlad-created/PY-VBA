im Recipients(31, 4) As Variant
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    Dim sht As Worksheet

Sub BreakFranchise()

    Workbooks.Open (ThisWorkbook.Path & "\" & Cells(2, 5) & Cells(2, 6))
     
    For Each sht In Worksheets
    
    If sht.Name <> "Incentive Dec'20_Shops" Then

    lastRow = sht.Range("B" & sht.Rows.Count).End(xlUp).Row
    
    On Error GoTo ErrorHandler

    Set rng = sht.Range("B" & sht.UsedRange.Find("*Shop Name*", , xlValues, xlWhole).Row & ":B" & lastRow)

    ' filter and delete all but header row
    With rng
        .AutoFilter Field:=sht.UsedRange.Find("*Shop Name*", , xlValues, xlWhole).Column, Criteria1:="<>*ìàãàçèí*"
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End With

    ' turn off the filters
    sht.AutoFilterMode = False
    
    End If
    
    Next
    
ErrorHandler:
        
    Application.DisplayAlerts = False
        
    ActiveWorkbook.SaveAs Filename:= _
        ThisWorkbook.Path & "\" & ThisWorkbook.Sheets(1).Cells(2, 5) & "_Franchise" & ThisWorkbook.Sheets(1).Cells(2, 6) _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = True
    ActiveWindow.Close
        
    Exit Sub
    
End Sub