import magic
import os
import sys

dir_path = sys.argv[1]
file_list = []
for folder, subfolder, files in os.walk(dir_path):
    for f in files:
        full_path = os.path.join(folder, f)
        file_list.append(full_path)

for item in file_list:
    File_To_Scan = item
    m = magic.open(magic.MAGIC_NONE)
    m.load()
    ftype = m.file(File_To_Scan)
    print(item)
    print(ftype)