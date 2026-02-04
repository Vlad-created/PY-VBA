import subprocess
import win32com.client as win32
import time

# Kill any leftover Access processes
subprocess.call("taskkill /f /im msaccess.exe", shell=True)

db_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\Rank-ACDB.accdb"
macro_name = "Macro1"

# Start Access
access = win32.Dispatch("Access.Application")
access.Visible = False

# Open DB
access.OpenCurrentDatabase(db_path)

print("Running macro...")
access.DoCmd.RunMacro(macro_name)
print("Macro completed.")

# Close DB
access.CloseCurrentDatabase()
access.Quit()

print("Access closed.")
