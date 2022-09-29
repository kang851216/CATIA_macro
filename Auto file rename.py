import glob
import os

path = "C:\\Work\\Compactor\\Mobile Compactor_W200908(SS2009007)_HKJC\\D01_CATIA\\As-build_final\\Manufacturing drawing\\"  # Drawing path
pattern = path + "APC10" + "*.CATDrawing"                                                                                   # Filter condition

# List of the files that match the pattern
result = glob.glob(pattern)

# Iterating the list with the count
count = 1
for file_name in result:
    old_name = file_name
    new_name = old_name.replace("APC10","APC10JC")
    os.rename(old_name, new_name)
    count = count + 1
