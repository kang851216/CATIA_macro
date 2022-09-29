import glob
import os
from win32com.client import Dispatch

CATIA = Dispatch('CATIA.Application')   # optional CATIA visibility
'''
file_name = "C:\\Users\\paul.kang\\Desktop\\Manufacturing drawing\\AP10JC.CATDrawing"
doc = CATIA.Documents.Open(file_name)
'''
doc = CATIA.ActiveDocument
DrawingSheets = doc.Sheets
DrawingSheet = DrawingSheets.Item(1)
DrawingViews = DrawingSheet.Views
DrawingView = DrawingViews.ActiveView
viewLinks = DrawingView.GenerativeLinks
dele = viewLinks.FirstLink
print(dele)



'''
Myview = actdoc.ActiveSheet.ActiveView
viewLinks = Myview.DrawingViewGenerativeLinks
del_update = viewLinks.RemoveAllLinks()
'''
