import os
from win32com.client import Dispatch
import pyautogui as auto
import time
import pandas

info = pandas.read_excel('BOM1.xlsx', sheet_name='Sheet')                                                                                         # Drawing list for batch process
path = "C:\\Users\\paul.kang\\Desktop\\Linen separator_Macau\\Linen_Collector_Rotate_Drum_220308_new\\"                                           # Original Part File path
drawingpath = "C:\\Users\\paul.kang\\Desktop\\Linen separator_Macau\\Linen_Collector_Rotate_Drum_220308_new\\manufacturing drawing\\test\\part\\" # Path to save draiwng 

CATIA = Dispatch('CATIA.Application')   
CATIA.Visible = True 

filelists = os.listdir(path)
files = []
filenames =[]


#len(drawinglist)
for i in range (18, 400, 1): 
    partname = info.iloc[i,2] #model name 
    quantity = int(info.iloc[i,3]) #Quantity
    inde = info.iloc[i,4] #翻滚

    exten = '.CATPart'
    drawingexten = '.CATDrawing'
    model = partname + exten
    file = path + partname + exten
    #print(partname)
    #print(i)
    #print(file)
    doc = CATIA.Documents.Open(file) # CATIA part open

    documents1 = CATIA.Documents

    partDocument1 = documents1.Item(model)
    product1 = partDocument1.GetItem(partname)
    
    part_parameters = partDocument1.part.parameters
    try:
        if part_parameters.item("L").istrueparameter:
            Len = part_parameters.item("L").value
    except:
        Len = 1
   
  
    
    scale_factor=int(Len/310)
    scale = 1
    if scale_factor <= 1:
        scale = 1
    elif scale_factor > 1 and scale_factor <= 2:
        scale = 1/2
    elif scale_factor > 2 and scale_factor <= 3:
        scale = 1/3
    elif scale_factor > 3 and scale_factor <= 5:
        scale = 1/5
    elif scale_factor > 5 and scale_factor <= 8:
        scale = 1/8
    elif scale_factor > 8 and scale_factor <= 10:
        scale = 1/10
    elif scale_factor > 10 and scale_factor <= 12:
        scale = 1/12
    elif scale_factor > 12 and scale_factor <= 15:
        scale = 1/15
    elif scale_factor > 15 and scale_factor <= 18:
        scale = 1/18
    elif scale_factor > 18 and scale_factor <= 20:
        scale = 1/20
    elif scale_factor > 20 and scale_factor <= 25:
        scale = 1/25
    elif scale_factor > 25 and scale_factor <= 30:
        scale = 1/30
    else:
        scale = 1/40
    #print(scale_factor)
    #print(scale)


    drawingDocument1 = documents1.Add('Drawing')
    drawingDocument1.Standard = 1
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item('Sheet.1')
    drawingSheet1.PaperSize = 5 #A3
    #drawingSheet1.Orientation = catPaperLandscape
    drawingViews1 = drawingSheet1.Views
    partDocument1 = documents1.Item(model)
    product1 = partDocument1.GetItem(partname)
    req1 = "技术要求"
    req2 = "1.未注明公差为±0.1mm"
    text1=drawingViews1.Activeview.Texts.Add(req1, 25, 35)
    text2=drawingViews1.Activeview.Texts.Add(req2, 25, 25)
    text1.SetFontName(0,0,'Arial Unicode MS (TrueType)')
    text2.SetFontName(0,0,'Arial Unicode MS (TrueType)')

    #Add iso view
    drawingView0 = drawingViews1.Add('AutomaticNaming')
    drawingView0.X = 315
    drawingView0.Y = 210
    drawingView0.Scale = scale
    drawingViewGenerativeLinks0 = drawingView0.GenerativeLinks
    drawingViewGenerativeBehavior0 = drawingView0.GenerativeBehavior
    drawingViewGenerativeBehavior0.Document = product1
    drawingViewGenerativeBehavior0.DefineIsometricView(-0.707, 0.707, 0.707, 0, 0, 0)
    drawingViewGenerativeBehavior0.ColorInheritanceMode = 1
    drawingViewGenerativeBehavior0.RepresentationMode = 0
    drawingViewGenerativeBehavior0.Update()

    #Add front view
    drawingView1 = drawingViews1.Add('AutomaticNaming')
    drawingView1.X = 210
    drawingView1.Y = 148.5
    drawingView1.Scale = scale
    drawingViewGenerativeLinks1 = drawingView1.GenerativeLinks
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    drawingViewGenerativeBehavior1.Document = product1
    drawingViewGenerativeBehavior1.DefineFrontView(1, 0, 0, 0, 1, 0)
    drawingViewGenerativeBehavior1.HiddenLineMode=1 # 0: hide hidden line, 1: show hidden line
    drawingViewGenerativeBehavior1.Update()

    if inde == '翻滚' or inde == '折弯':
        #Add unfolded view
        print(partname)
        print(inde)
        drawingView4 = drawingViews1.Add('AutomaticNaming')
        drawingView4.X = 315
        drawingView4.Y = 148.5
        drawingView4.Scale = scale
        drawingViewGenerativeLinks4 = drawingView4.GenerativeLinks
        drawingViewGenerativeBehavior4 = drawingView4.GenerativeBehavior
        drawingViewGenerativeBehavior4.Document = product1
        drawingViewGenerativeBehavior4.DefineUnfoldedView(0., 0., 1., 1., 0., 0.)
        drawingViewGenerativeBehavior4.Update()
   
    time.sleep(2)
    os.startfile("C:\\Work\\CAD material\\CATIA_macro\\Shop drawing macro\\Macau.CATScript")
    time.sleep(5)
    quan = drawingViews1.Activeview.Texts.Add(quantity, 314, 41)
    quan.SetFontName(0,0,'Arial Unicode MS (TrueType)')

    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    drawingSheet1.GenerateDimensions
    drawingView1.Activate()
    
    
    drawingfilename = drawingpath + partname 
    drawingDocument1.SaveAs(drawingfilename)
    drawingDocument1.Close()
    partDocument1.Close()
    #print(partname)

