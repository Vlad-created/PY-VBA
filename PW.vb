
Sub RefreshWithProtection()
   Dim pw As String
   pw = "!@Sofia!@01"

   ThisWorkbook.Unprotect Password:=pw
   ThisWorkbook.RefreshAll
   ThisWorkbook.Protect Structure:=True, Password:=pw
End Sub