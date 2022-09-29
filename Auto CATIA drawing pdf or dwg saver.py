import glob
import os
from win32com.client import Dispatch
import time

path = "C:\\Work\\ARCS\\ARCS_Macau Island Hospital\\D01_Design\\Linen separator_Macau\\Linen_Collector_Rotate_Drum_220308_new\\manufacturing drawing\\" 


pattern = path + "LCA_A" + "*.CATDrawing"       # Name pattern for bulk process
CATIA = Dispatch('CATIA.Application') 

# List of the files that match the pattern
result = glob.glob(pattern)


for file in result:
    print("Opening file")
    newfilename = file.replace('CATDrawing','dwg') # File name expansion change
    doc = CATIA.Documents.Open(file)
    print(newfilename)

    doc.ExportData(newfilename, "dwg")             # Save as dwg or pdf
    time.sleep(3)
    doc.Close()
    print("Completed!")