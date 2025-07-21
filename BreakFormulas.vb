Sub BreakFormulas()

Dim sht As Worksheet

    Application.ScreenUpdating = False
    
    Workbooks.Open (ThisWorkbook.Path & "\" & Cells(2, 5) & Cells(2, 6))
    
    For Each sht In Worksheets
        sht.Visible = xlSheetVisible
        sht.Select
        With sht.UsedRange
            ActiveSheet.Outline.ShowLevels RowLevels:=2
            .Cells.Copy
            .Cells.PasteSpecial xlPasteValues
            .Cells(1).Select
        End With
        Application.CutCopyMode = False
    Next sht

    Application.DisplayAlerts = False
    
    Worksheets("working days").Delete
    Worksheets("days").Delete
    Worksheets("Sheet1").Delete
    Worksheets("CO TA% Bonus").Select
    
    Application.DisplayAlerts = True

    ActiveWorkbook.Save
    ActiveWindow.Close

    Application.ScreenUpdating = True
    
End Sub