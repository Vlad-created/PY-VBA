' You will need to add Microsoft Outlook 12.0 Object Library to your references for this code to work
' (or Microsoft Outlook 11.0 Object Library if you are using Office 2003).

Dim Recipients(31, 4) As Variant

Dim myWorkbook As Excel.Workbook
Dim myWorkbookFileName As String
Dim myFilePath As String

Dim myOutlook As Outlook.Application
Dim myMessage As Outlook.MailItem

Sub MassMailing()

For j = 0 To 31
Recipients(j, 1) = Sheets("Sheet1").Cells(j + 2, 1)  ' Column A of Sheet1
Recipients(j, 2) = Sheets("Sheet1").Cells(j + 2, 3)  ' Column C of Sheet1, the recipients' e-mail addresses
Recipients(j, 3) = Sheets("Sheet1").Cells(j + 2, 4)  ' Column D of Sheet1, the recipients' names
Recipients(j, 4) = Sheets("Sheet1").Cells(j + 2, 2)  ' Column B of Sheet1, the regions' names

Next j


myFilePath = ThisWorkbook.Path & "\"
'  (Substitute the file path to the folder with the files you want to send.)

For I = 1 To 31   ' The upper value of i is the number of groups to whom you want to send files.
    For j = 0 To 31
myWorkbookFileName = ThisWorkbook.Sheets(1).Cells(2, 5) & "_" & Recipients(j, 4) & ThisWorkbook.Sheets(1).Cells(2, 6)    '   OR & ".xls" if you are using Office 2003
Set myOutlook = GetObject(, "Outlook.Application")
Set myMessage = myOutlook.CreateItem(ItemType:=olMailItem)

    With myMessage

        If I = Recipients(j, 1) Then
        On Error Resume Next
        .To = Recipients(j, 2)
        .Subject = Sheets("Sheet1").Cells(32, 1)
        .BodyFormat = olFormatHTML
        .HTMLBody = "Çäðàâåéòå,<br><br>" & Sheets("Sheet1").Cells(33, 1) & "<br><br>" & "Ïîçäðàâè!" & .HTMLBody   ' Or whatever you want to insert
        .Attachments.Add myFilePath & myWorkbookFileName
        .SentOnBehalfOfName = "reportsteam@a1.bg"
        .Send
        End If
         
    End With
    
       Next j
       
    Next I
    
    
    Set myMessage = Nothing
    Set myOutlook = Nothing
    Exit Sub
    
End Sub