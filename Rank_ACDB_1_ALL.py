import subprocess
import win32com.client as win32

# Kill any leftover Access processes
subprocess.call("taskkill /f /im msaccess.exe", shell=True)

def run_access_macro(db_path, macro_name):
    print(f"\nOpening database:\n{db_path}")
    
    access = win32.Dispatch("Access.Application")
    access.Visible = False

    access.OpenCurrentDatabase(db_path)

    print(f"Running macro: {macro_name}")
    access.DoCmd.RunMacro(macro_name)
    print("Macro completed.")

    access.CloseCurrentDatabase()
    access.Quit()

    print("Database closed.")


# 1️⃣ First ACCDB
run_access_macro(
    r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\GOAT_FTW.accdb",
    "Macro1 - current month"
)

# 2️⃣ Second ACCDB
run_access_macro(
    r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\Rank-ACDB.accdb",
    "Macro1"
)

# 3️⃣ Third ACCDB — run 3 macros in sequence
third_db = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\Rank Calcualtions.accdb"
third_macros = ["1_Macro1", "2_Macro2", "3_Macro3"]

for m in third_macros:
    run_access_macro(third_db, m)

print("\nAll macros executed successfully.")
