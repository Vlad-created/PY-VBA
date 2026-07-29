import zipfile
import shutil
import os
from tempfile import mkdtemp

input_file = r"C:\Users\a1bg542048\OneDrive - A1 Telekom Austria AG\Desktop\Check check .xlsx"
output_file = r"C:\Users\a1bg542048\OneDrive - A1 Telekom Austria AG\Desktop\Check check_unlocked.xlsx"

tmpdir = mkdtemp()

with zipfile.ZipFile(input_file, 'r') as zin:
    zin.extractall(tmpdir)

rels_path1 = os.path.join(tmpdir, "xl", "workbook.xml")
rels_path2 = os.path.join(tmpdir, "xl", "workbook.xml.rels")
rels_path3 = os.path.join(tmpdir, "xl", "_rels", "workbook.xml.rels")

paths = [rels_path1, rels_path2, rels_path3]

for path in paths:
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            xml = f.read()

        if "fileSharing" in xml:
            xml = xml.replace(
                xml[xml.find("<fileSharing"):xml.find("/>", xml.find("<fileSharing"))+2],
                ""
            )

            with open(path, "w", encoding="utf-8") as f:
                f.write(xml)

shutil.make_archive(output_file.replace(".xlsx", ""), 'zip', tmpdir)
os.rename(output_file.replace(".xlsx", "") + ".zip", output_file)
