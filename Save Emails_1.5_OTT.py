import time
import win32com.client as win32

file_path = r"\\hope\Shop support\smart_app_ott_watched.xlsx"

excel = win32.Dispatch("Excel.Application")
excel.Visible = True

wb = excel.Workbooks.Open(file_path)

print("Connections found:")
for c in wb.Connections:
    print(" -", c.Name)

print("\nRefreshing Power BI connections...")

for c in wb.Connections:
    try:
        oledb = c.OLEDBConnection
        print(f"Refreshing: {c.Name}")
        oledb.Refresh()
    except Exception as e:
        print(f"Skipping {c.Name}: not an OLEDB connection ({e})")

# Wait until all connections finish
while any(getattr(c, "Refreshing", False) for c in wb.Connections):
    time.sleep(0.5)

time.sleep(1)
wb.Save()

print("\nRefresh complete and workbook saved.")
