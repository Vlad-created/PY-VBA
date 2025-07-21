    Dim Recipients(31, 4) As Variant
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    Dim sht As Worksheet

Sub BreakAllRegions()
     
    For j = 1 To 13
    
    Recipients(j, 4) = Sheets("Sheet1").Cells(j + 1, 2)  ' Column B of Sheet1, the regions' names

        Workbooks.Open (ThisWorkbook.Path & "\" & Cells(2, 5) & Cells(2, 6)) ' Returns the folder path (excluding the file name) of the workbook running the macro.
        
    For Each sht In Worksheets
    
    If sht.Name <> "Incentive Dec'20_Shops" Then

    lastRow = sht.Range("B" & sht.Rows.Count).End(xlUp).Row

    Set rng = sht.Range("B" & sht.UsedRange.Find("*REGION*", , xlValues, xlWhole).Row & ":B" & lastRow)

    ' filter and delete all but header row
    With rng
        .AutoFilter Field:=sht.UsedRange.Find("*REGION*", , xlValues, xlWhole).Column, Criteria1:="<>*" & Recipients(j, 4) & "*"
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End With

    ' turn off the filters
    sht.AutoFilterMode = False
    
    End If
    
    Next
    
    Application.DisplayAlerts = False
        
    ActiveWorkbook.SaveAs Filename:= _
        ThisWorkbook.Path & "\" & ThisWorkbook.Sheets(1).Cells(2, 5) & "_" & Recipients(j, 4) & ThisWorkbook.Sheets(1).Cells(2, 6) _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = True
    ActiveWindow.Close
    
    Next j
        
    Exit Sub
    
End Sub


