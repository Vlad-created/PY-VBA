Sub BreakAllRegions()
    Dim Recipients(31, 4) As Variant
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    Dim sht As Workasheet
    Dim reionCol As Long
    Dim headerRow As Long 



     
    For j = 1 To 13
    
    Recipients(j, 4) = Sheets("Sheet1").Cells(j + 1, 2)  ' Column B of Sheet1, the regions' names

        Workbooks.Open (ThisWorkbook.Path & "\" & Cells(2, 5) & Cells(2, 6)) ' Returns the folder path (excluding the file name) of the workbook running the macro.
        
     Set sht = Worksheets("Base")

     headerRow = sht.UsedRange.Find("*Region*", , xlValues, xlWhole).Row
     regionCol = sht.UsedRange.Find("*Region*", , xlValues, xlWhole).Column
     sht.Rows(headerRow & ":" & sht.Rows.Count).AutoFilter Field:=regionCol, Criteria1:="*" & Recipients(j, 4) & "*"

     lastRow = sht.Range("W" & sht.Rows.Count).End(xlUp).Row

     Set rng = sht.Range("W" & sht.UsedRange.Find("*Region*", , xlValues, xlWhole).Row & ":W" & lastRow)

     With rng
     .AutoFilter Field:=sht.UsedRange.Find("*Region*", , xlValues, xlWhole).Column, Criteria1:="<>*" & Recipients(j, 4) & "*"
     .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
      End With

      Application.DisplayAlerts = False
      ActiveWorkbook.SaveAs Filename:= _
      ThisWorkbook.Path & "\" & ThisWorkbook.Sheets(1).Cells(2, 5) & "_" & Recipients(j, 4) & ThisWorkbook.Sheets(1).Cells(2, 6), _
      FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
      Application.DisplayAlerts = True

      sht.AutoFilterMode = False
    
    End If
    
    Next
    
    Application.DisplayAlerts = False
        
     ActiveWorkbook.SaveAs Filename:= _
        ThisWorkbook.Path & "\" & ThisWorkbook.Sheets(1).Cells(2, 5) & "_" & Recipients(j, 4) & ThisWorkbook.Sheets(1).Cells(2, 6) _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = True

      On Error Resume Next
      wbCopy.RefreshAll
      On Error GoTo 0

wbCopy.Close SaveChanges:=False
    ActiveWindow.Close
    
    Next j
        
    Exit Sub
    
End Sub


Sub BreakAllRegions()
    Dim Recipients(31, 4) As Variant
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    Dim sht As Worksheet
    Dim regionCol As Long
    Dim headerRow As Long
    Dim wbCopy As Workbook
    Dim j As Long
    Dim fileName As String
    Dim yesterday AS String 

    For j = 1 To 13
        Recipients(j, 4) = ThisWorkbook.Sheets("Sheet1").Cells(j + 1, 2).Value

        Workbooks.Open (ThisWorkbook.Path & "\" & Cells(2, 5) & Cells(2, 6))

        Set wbCopy = Workbooks.Open(fileName)
        Set sht = wbCopy.Worksheets("Base")

        With sht.UsedRange
            headerRow = .Find("Region", , xlValues, xlWhole).Row
            regionCol = .Find("Region", , xlValues, xlWhole).Column
        End With

        sht.Rows(headerRow & ":" & sht.Rows.Count).AutoFilter Field:=regionCol, Criteria1:="*" & Recipients(j, 4) & "*"

        lastRow = sht.Range("W" & sht.Rows.Count).End(xlUp).Row
        Set rng = sht.Range("W" & headerRow & ":W" & lastRow)

        With rng
            .AutoFilter Field:=regionCol, Criteria1:="<>*" & Recipients(j, 4) & "*"
            On Error Resume Next
            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            On Error GoTo 0
        End With

        sht.AutoFilterMode = False

        Application.DisplayAlerts = False
        wbCopy.RefreshAll
        wbCopy.SaveAs Filename:= _
            saveName = ThisWorkbook.Path & "\" & "Shop_Report_" & Format(Date - 1, "dd.mm.yyyy") & "_" & Recipients(j, 4) & ".xlsx"
            FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        wbCopy.Close SaveChanges:=False
        Application.DisplayAlerts = True
    Next j
End Sub









