Sub BreakAllRegions()
    Dim wbSource As Workbook, wbCopy As Workbook
    Dim wsBase As Worksheet, wsFiltered As Worksheet, sht As Worksheet
    Dim zoneCriteria As String
    Dim lastRow As Long
    Dim filePath As String, copyFilePath As String

    ' Step 1: Get filter criteria
    zoneCriteria = ThisWorkbook.Sheets("Sheet1").Range("B2").Value

    ' Step 2: Get path from Sheet1
    Set sht = ThisWorkbook.Sheets("Sheet1")
    filePath = ThisWorkbook.Path & "\" & sht.Cells(2, 5).Value & sht.Cells(2, 6).Value

    ' Step 3: Open source workbook and refresh data
    Set wbSource = Workbooks.Open(filePath)
    wbSource.RefreshAll
    DoEvents

    Set wsBase = wbSource.Sheets("Sheet2")

    ' Step 4: Apply filter
    lastRow = wsBase.Cells(wsBase.Rows.Count, "W").End(xlUp).Row
    With wsBase
        .AutoFilterMode = False
        .Range("A1:AA" & lastRow).AutoFilter Field:=23, Criteria1:=zoneCriteria
    End With

    ' Step 5: Save a copy of the full workbook
    copyFilePath = ThisWorkbook.Path & "\" & zoneCriteria & "_Shop_Report_" & Format(Date - 1, "dd.mm.yyyy") & ".xlsx"
    wbSource.SaveCopyAs copyFilePath

    ' Step 6: Open the copy and clear Sheet2, replace with filtered values
    Set wbCopy = Workbooks.Open(copyFilePath)
    Set wsFiltered = wbCopy.Sheets("Sheet2")

    On Error Resume Next
    wsFiltered.Cells.ClearContents
    wsBase.Range("A1:AA" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsFiltered.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    wsFiltered.Cells.EntireColumn.AutoFit

    wsFiltered.AutoFilterMode = False

    ' Step 7: Final refresh, save, close
    wbCopy.RefreshAll
    DoEvents
    wbCopy.Save
    wbCopy.Close False
    wbSource.Close False

    MsgBox "Filtered rows copied to Sheet2, and original sheets preserved!", vbInformation
End Sub







