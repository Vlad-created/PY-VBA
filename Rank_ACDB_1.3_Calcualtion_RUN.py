import subprocess
import win32com.client as win32
import time

# Kill any leftover Access processes
subprocess.call("taskkill /f /im msaccess.exe", shell=True)

db_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\Rank Calcualtions.accdb"
macros = ["1_Macro1", "2_Macro2", "3_Macro3"]

# Start Access
access = win32.Dispatch("Access.Application")
access.Visible = False

# Open DB
access.OpenCurrentDatabase(db_path)

for m in macros:
    print(f"Running macro: {m}")
    access.DoCmd.RunMacro(m)
    print(f"Completed: {m}")
    time.sleep(1)  # optional, gives Access breathing room

# Close DB
access.CloseCurrentDatabase()
access.Quit()

print("Access closed.")
