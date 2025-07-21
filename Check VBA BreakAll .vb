Sub BreakAllRegions()

    Dim Recipients(31, 4) As Variant
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    Dim sht As Worksheet
    Dim wbTarget As Workbook
    Dim regionRow As Range
    Dim fileName As String
    Dim saveName As String
    Dim yesterday As String
    Dim j As Integer

    yesterday = Format(Date - 1, "dd.mm") ' Format for opening file

    For j = 1 To 13
        Recipients(j, 4) = ThisWorkbook.Sheets("Sheet1").Cells(j + 1, 2).Value

        fileName = ThisWorkbook.Path & "\" & "UFC_sent_" & yesterday & ".xlsx"

        On Error Resume Next
        Set wbTarget = Workbooks.Open(fileName)
        On Error GoTo 0

        If wbTarget Is Nothing Then
            MsgBox "Could not open file: " & fileName, vbExclamation
            Exit Sub
        End If

        For Each sht In wbTarget.Worksheets
            If sht.Name <> "Incentive Dec'20_Shops" Then
                lastRow = sht.Cells(sht.Rows.Count, "B").End(xlUp).Row
                Set regionRow = sht.UsedRange.Find("REGION", , xlValues, xlWhole)

                If Not regionRow Is Nothing Then
                    Set rng = sht.Range("B" & regionRow.Row & ":B" & lastRow)

                    With sht
                        rng.AutoFilter Field:=1, Criteria1:="<>*" & Recipients(j, 4) & "*"
                        On Error Resume Next
                        rng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
                        On Error GoTo 0
                        .AutoFilterMode = False
                    End With
                End If
            End If
        Next sht

        ' Save with formatted name including full date
        saveName = ThisWorkbook.Path & "\" & "UFC_sent_" & Format(Date - 1, "dd.mm.yyyy") & "_" & Recipients(j, 4) & ".xlsx"

        Application.DisplayAlerts = False
        wbTarget.SaveAs Filename:=saveName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        wbTarget.Close SaveChanges:=False
        Application.DisplayAlerts = True
    Next j

    MsgBox "All regions processed successfully!"

End Sub

