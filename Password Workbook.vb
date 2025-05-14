
Sub RefreshWithProtection()
   Dim pw As String
   pw = "QWEqwe!23"

   ThisWorkbook.Unprotect Password:=pw
   ThisWorkbook.RefreshAll
   ThisWorkbook.Protect Structure:=True, Password:=pw
End Sub
