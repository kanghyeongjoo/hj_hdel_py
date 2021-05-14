import win32com.client
import math
import string
import fnmatch

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

doc.SendCommand('setxdata ')

for entity in doc.ModelSpace:
    if entity.EntityName == "AcDbRotatedDimension" and entity.TextOverride.strip(string.punctuation2) == "출입구 유효폭":
        Xdata = entity.GetXData("", "Type", "Data")
        pt1 = Xdata[1][len(Xdata[1])-2]
        pt2 = Xdata[1][len(Xdata[1])-1]
        print(pt1, pt2)

for entity in doc.ModelSpace:
    if entity.EntityName == "AcDbRotatedDimension" and entity.TextOverride.strip(string.punctuation2) == "승강로 내부":
        Xdata = entity.GetXData("", "Type", "Data")
        pt1 = Xdata[1][len(Xdata[1])-2]
        pt2 = Xdata[1][len(Xdata[1])-1]
        print(pt1, pt2)