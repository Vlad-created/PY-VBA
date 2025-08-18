Sub RefreshRegionFiles()
    Dim regionName As String
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim con As WorkbookConnection
    Dim i As Integer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    folderPath = ThisWorkbook.Path & "\"

    For i = 1 To 14
        On Error Resume Next
        regionName = Trim(Sheets("Sheet1").Cells(i + 1, 2).Value)
        On Error GoTo 0

        If regionName <> "" Then
            fileName = "FBB_Cancellation Rate_" & Format(Date - 1, "dd.mm") & "_" & regionName & ".xlsx"
            fullPath = folderPath & fileName

            If Dir(fullPath) <> "" Then
                Set wb = Workbooks.Open(fullPath)

                ' Refresh all OLEDB connections (Access)
                For Each con In wb.Connections
                    If con.Type = xlConnectionTypeOLEDB Then
                        On Error Resume Next
                        con.OLEDBConnection.BackgroundQuery = False
                        con.OLEDBConnection.Refresh
                        On Error GoTo 0
                    End If
                Next con

                ' Refresh all PivotCaches (forces PivotTables to update)
                For Each ws In wb.Worksheets
                    For Each pt In ws.PivotTables
                        On Error Resume Next
                        pt.PivotCache.Refresh
                        On Error GoTo 0
                    Next pt
                Next ws

                ' Wait briefly to ensure refresh completes
                Application.Wait (Now + TimeValue("0:00:30"))

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
