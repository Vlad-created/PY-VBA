import win32com.client as win32

# Path to your Access database
access_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Bonuses\Ranking\Rank_PBI\Rank Calcualtions.accdb"

# Macro name
macro_name = "1_Macro1"

# Start Access
access = win32.Dispatch("Access.Application")

# Open the database
access.OpenCurrentDatabase(access_path)

print(f"Running Access macro: {macro_name}")

# Run the macro
access.DoCmd.RunMacro(macro_name)

print("Macro completed.")

# Close the database and quit Access
access.CloseCurrentDatabase()
access.Quit()

print("Access closed.")
