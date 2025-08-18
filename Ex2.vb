Sub AddFormula()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Detail1")
    
   

    ws.Range("CM2").Formula = "=IFNA(XLOOKUP([CE2,Maping!B:B,Maping!C:C),"" "")"
    ws.Range("CN2").Formula = "=IF(AND(CD2=""Refused"",[@Refused_by_Area]=""Sales"",[@Refused_Reasons_Grouped2]=""поради отказ на клиента от услуга""),1,0)"
    

   ws.Range("CM:CN" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).FillDown 

   ws.Range("CM:CN").Value = ws.Range("CM:CN").Value 
   
End Sub




=IFNA(XLOOKUP(CE2,Maping!B:B,Maping!C:C)," ")
=IF(AND(CD2="Refused",CG2="Sales",CM2="поради отказ на клиента от услуга"),1,0)