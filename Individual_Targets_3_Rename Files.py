import os
import datetime

folder = r"\\virgo\Work\OpControl\2016\Bonuses\1_MassCurrMonth\Targets"

# Get current month name (e.g., January)
current_month = datetime.datetime.now().strftime("%B")

# List of all month names to detect and replace
month_names = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Loop through files in folder
for filename in os.listdir(folder):
    old_path = os.path.join(folder, filename)

    if not os.path.isfile(old_path):
        continue

    # Check if filename contains any month name
    for old_month in month_names:
        if old_month in filename:
            new_filename = filename.replace(old_month, current_month)
            new_path = os.path.join(folder, new_filename)

            # Avoid overwriting existing files
            if os.path.exists(new_path):
                print(f"⚠️ Skipping rename (target exists): {new_filename}")
                continue

            print(f"Renaming:\n  {filename}\n→ {new_filename}")
            os.rename(old_path, new_path)
            break

print("\n✅ Renaming completed with no collisions.")
