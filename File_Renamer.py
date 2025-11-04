import os

def rename_documents(folder_path, initials="CV"):
    # Define allowed extensions
    allowed_extensions = ('.xls', '.xlsx', '.doc', '.docx', '.pdf')
    
    # Get matching files
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(allowed_extensions)]
    
    # Sort for consistent numbering
    files.sort()

    for idx, filename in enumerate(files, start=1):
        name, ext = os.path.splitext(filename)
        new_name = f"{name}_{initials}_{idx}{ext}"
        
        old_path = os.path.join(folder_path, filename)
        new_path = os.path.join(folder_path, new_name)
        
        os.rename(old_path, new_path)
        print(f"Renamed: {filename} → {new_name}")

# UNC path to your network folder
folder_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\Reports_Instructions"
rename_documents(folder_path)


######
import os
 
def rename_documents(folder_path, initials="CV"):
    # Define allowed extensions
    allowed_extensions = ('.xls', '.xlsx', '.doc', '.docx', '.pdf')
 
    # Get matching files
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(allowed_extensions)]
 
    # Sort for consistent numbering
    files.sort()
    for idx, filename in enumerate(files, start=1):
        name, ext = os.path.splitext(filename)
        new_name = f"{name}_{initials}_{idx}{ext}"
        old_path = os.path.join(folder_path, filename)
        new_path = os.path.join(folder_path, new_name)
 
        os.rename(old_path, new_path)
        print(f"Renamed: {filename} → {new_name}")
 
# UNC path to your network folder
folder_path = r"C:\Users\vladv\Desktop\New folder"
 
rename_documents(folder_path)
########