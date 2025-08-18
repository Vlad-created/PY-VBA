Sub RefreshRegionFiles()
    Dim regionName As String
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim wb As Workbook
    Dim i As Integer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    folderPath = ThisWorkbook.Path & "\"

    For i = 1 To 13
        ' Get the region name from Sheet1 column B
        regionName = Trim(Sheets("Sheet1").Cells(i + 1, 2).Value)
        If regionName <> "" Then
            ' Construct expected file name
            fileName = "FBB_Cancellation Rate_" & Format(Date - 1, "dd.mm") & "_" & regionName & ".xlsx"
            fullPath = folderPath & fileName

            ' Check if the file exists
            If Dir(fullPath) <> "" Then
                Set wb = Workbooks.Open(fullPath)
                wb.RefreshAll
                DoEvents ' Let the refresh complete
                wb.Save
                wb.Close False
            Else
                MsgBox "File not found: " & fileName, vbExclamation
            End If
        End If
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Refresh cycle complete for all available region files.", vbInformation
End Sub



