import win32com.client
import math
import string

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

def Find_Duct_Hole_pst():

    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-HOISTWAY-MP-AC-TYPE3-H":
            hoistway_x_cdnt = int(entity.InsertionPoint[0])
            hoistway_y_cdnt = int(entity.InsertionPoint[1])

    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-HBTN-HOLE":
            duct_hole_x_cdnt = int(entity.InsertionPoint[0])
            duct_hole_y_cdnt = int(entity.InsertionPoint[1])

    if hoistway_x_cdnt < duct_hole_x_cdnt:
        duct_hole_x_pst = "Left"
    else:
        duct_hole_x_pst = "Right"

    if hoistway_y_cdnt < duct_hole_y_cdnt:
        duct_hole_y_pst = "Top"
    else:
        duct_hole_y_pst = "Bottom"

    return duct_hole_x_pst, duct_hole_y_pst

def Set_Hoistway_Dim_cdnt():

    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-HOISTWAY-MP-AC-TYPE3":
            entity.Explode()

    doc.SendCommand('setxdata ')

    for entity in doc.ModelSpace:
        if entity.EntityName == "AcDbRotatedDimension" and entity.TextOverride.strip(string.punctuation2) == "승강로 내부":
            Xdata = entity.GetXData("", "Type", "Data")
            pt1 = Xdata[1][len(Xdata[1]) - 2]
            pt2 = Xdata[1][len(Xdata[1]) - 1]
            if int(pt1[0]) == int(pt2[0]):
                hoistway_ver = (pt1[1], pt2[1])
            else:
                hoistway_hor = (pt1[0], pt2[0])

    return hoistway_hor, hoistway_ver

def Find_Duct_Hole_cdnt(x_pst, y_pst, hor_cdnt, ver_cdnt):

    if x_pst == "Left":
        duct_x_cdnt = min(hor_cdnt)
    else:
        duct_x_cdnt = max(hor_cdnt)

    if y_pst == "Bottom":
        duct_y_cdnt = min(ver_cdnt)
    else:
        duct_y_cdnt = max(ver_cdnt)

    duct_hole_cdnt = (duct_x_cdnt, duct_y_cdnt)

    return duct_hole_cdnt

def Find_CP_Cdnt():

    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-CP-MP-L":
            entity.Explode()

    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbAttributeDefinition' and entity.TagString == "@TEXT":
            CP_cdnt = (entity.TextAlignmentPoint[0], entity.TextAlignmentPoint[1])

    return CP_cdnt

def Cal_CP_to_Duct_Hole_dis(Duct_Hole_cdnt, CP_cdnt):

    cal_x_dis = abs(int(Duct_Hole_cdnt[0]) - int(CP_cdnt[0]))
    cal_y_dis = abs(int(Duct_Hole_cdnt[1]) - int(CP_cdnt[1]))

    CP_to_Duct_Hole_dis = cal_x_dis + cal_y_dis

    return CP_to_Duct_Hole_dis

Duct_Hole_cdnt = Find_Duct_Hole_cdnt(*Find_Duct_Hole_pst(), *Set_Hoistway_Dim_cdnt())
CP_cdnt = Find_CP_Cdnt()
CP_TO_DUCT_HOLE = Cal_CP_to_Duct_Hole_dis(Duct_Hole_cdnt, CP_cdnt)
print(CP_TO_DUCT_HOLE)

