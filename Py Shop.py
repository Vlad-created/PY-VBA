import openpyxl

wb = openpyxl.load_workbook(r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\FBB Cancellation\Base FBB Cancellation_Rate\Shops Report-20.07.2025.xlsx")

ws = wb["Sheet1"] 

Order_Create_Date = ws["H2"].value 

if Order_Create_Date != 7: 
    print("not ok") 
else: 
    print("ok")