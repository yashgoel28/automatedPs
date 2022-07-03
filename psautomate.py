import win32com.client
import os
from datetime import datetime
import calendar
import pandas as pd

options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
options.Format = 13   # PNG
options.PNG8 = False  # Sets it to PNG-24 bit

psApp =win32com.client.Dispatch("Photoshop.Application")

df = pd.read_excel(r"C:\Users\ishu\Desktop\db.xlsx", sheet_name="Sheet1")
for i in range(0,len(df)):
    nameval= df.values[i][0]
    mrpval = df.values[i][1]
    barname=df.values[i][2]
    wt=df.values[i][4]
    exp =df.values[i][3]



    template=psApp.Open(r"C:\Users\ishu\Desktop\anaaj mish\automate_template.psd")
    doc = psApp.Application.ActiveDocument

    psApp.Load(r"C:/Users/ishu/Desktop/anaaj mish/barcodes(3)/"+barname)
    psApp.ActiveDocument.Selection.SelectAll()
    psApp.ActiveDocument.Duplicate()
    psApp.ActiveDocument.Selection.SelectAll()
    psApp.ActiveDocument.Selection.Resize(30,30)
    psApp.ActiveDocument.Selection.Copy()
    psApp.ActiveDocument.Close(2)

    psApp.ActiveDocument.Close()

    SmartObjBarcode = psApp.Open(r"C:\Users\ishu\Desktop\anaaj mish\autoBarcode.psb")
    psApp.ActiveDocument = SmartObjBarcode
    SmartObjBarcode.Paste()
    for l in psApp.ActiveDocument.layers:
        if l!=psApp.ActiveDocument.ActiveLayer:
            l.Delete()
    SmartObjBarcode.Save()
    SmartObjBarcode.Close()
    psApp.ActiveDocument = template
    name = doc.ArtLayers["name"]
    name_txt =name.TextItem
    name_txt.contents=nameval
    name_txt.size=18

    detail = doc.ArtLayers["Net Wt"]
    detail_txt =detail.TextItem
    detail_txt.contents=wt
    detail_txt.size=6.5
    detail_txt.font="Arial"
    detail_txt.color.rgb.red=0
    detail_txt.color.rgb.blue=0
    detail_txt.color.rgb.green=0

    detail = doc.ArtLayers["MRP"]
    detail_txt =detail.TextItem
    detail_txt.contents=str(mrpval)+"/-"
    detail_txt.size=6.5
    detail_txt.font="Arial"
    detail_txt.color.rgb.red=0
    detail_txt.color.rgb.blue=0
    detail_txt.color.rgb.green=0

    detail = doc.ArtLayers["MFD"]
    detail_txt =detail.TextItem
    detail_txt.contents=str(calendar.month_name[datetime.now().month])+", "+str(datetime.now().year)
    detail_txt.size=6.5
    detail_txt.font="Arial"
    detail_txt.color.rgb.red=0
    detail_txt.color.rgb.blue=0
    detail_txt.color.rgb.green=0

    detail = doc.ArtLayers["EXP"]
    detail_txt =detail.TextItem
    detail_txt.contents=exp
    detail_txt.size=6.5
    detail_txt.font="Arial"
    detail_txt.color.rgb.red=0
    detail_txt.color.rgb.blue=0
    detail_txt.color.rgb.green=0

    psApp.ActiveDocument.Selection.SelectAll()
    widthRef=1476
    psApp.ActiveDocument.ActiveLayer= doc.ArtLayers["name"]
    name=doc.ArtLayers["name"].bounds
    widthchild = name[2]-name[0]
    psApp.ActiveDocument.ActiveLayer.Translate((widthRef-widthchild)/2-name[0])
    doc.Export(ExportIn="C:/Users/ishu/Desktop/sticker/"+nameval+" "+wt+".png", ExportAs=2, Options=options)
    doc.Close(2)
