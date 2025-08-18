Sub RenameOldRegionFiles()
    Dim folderPath As String
    Dim fileSystem As Object
    Dim file As Object
    Dim regionName As String
    Dim newFileName As String
    Dim oldFileName As String
    Dim yesterdayDate As String
    Dim i As Integer

    folderPath = ThisWorkbook.Path & "\"
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    yesterdayDate = Format(Date - 1, "dd.mm")

    For i = 1 To 14
        On Error Resume Next
        regionName = Trim(Sheets("Sheet1").Cells(i + 1, 2).Value)
        On Error GoTo 0

        If regionName <> "" Then
            ' Loop through files in the folder
            For Each file In fileSystem.GetFolder(folderPath).Files
                If InStr(file.Name, "FBB_Cancellation Rate_") > 0 And _
                   InStr(file.Name, "_" & regionName & ".xlsx") > 0 And _
                   InStr(file.Name, yesterdayDate) = 0 Then

                    ' Build new name
                    newFileName = "FBB_Cancellation Rate_" & yesterdayDate & "_" & regionName & ".xlsx"

                    ' Rename file
                    Name folderPath & file.Name As folderPath & newFileName
                End If
            Next file
        End If
    Next i

    MsgBox "Files renamed successfully.", vbInformation
End Sub