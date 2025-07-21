Dim Recipients(31, 4) As Variant

Dim myWorkbook As Excel.Workbook
Dim myWorkbookFileName As String
Dim myFilePath As String

Dim myOutlook As Outlook.Application
Dim myMessage As Outlook.MailItem

Sub MailingFranchise()

myFilePath = ThisWorkbook.Path & "\"
'  (Substitute the file path to the folder with the files you want to send.)

myWorkbookFileName = ThisWorkbook.Sheets(1).Cells(2, 5) & "_Franchise" & ThisWorkbook.Sheets(1).Cells(2, 6)    '   OR & ".xls" if you are using Office 2003
Set myOutlook = GetObject(, "Outlook.Application")
Set myMessage = myOutlook.CreateItem(ItemType:=olMailItem)

    With myMessage

        .To = "miroslav.naydenov@a1.bg"
        .Subject = Sheets("Sheet1").Cells(35, 1)
        .BodyFormat = olFormatHTML
        .HTMLBody = "Çäðàâåéòå,<br><br>" & Sheets("Sheet1").Cells(36, 1) & "<br><br>" & "Ïîçäðàâè!" & .HTMLBody   ' Or whatever you want to insert
        .Attachments.Add myFilePath & myWorkbookFileName
        .SentOnBehalfOfName = "reportsteam@a1.bg"
        .Send
         
    End With
    
    Set myMessage = Nothing
    Set myOutlook = Nothing
    Exit Sub
    
End Sub