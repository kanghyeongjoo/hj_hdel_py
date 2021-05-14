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
        duct_hole_x_pst = "Right"
    else:
        duct_hole_x_pst = "Left"

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


def Find_Gov_cdnt():

    for entity in doc.ModelSpace:
        if entity.EntityName == "AcDbBlockReference" and entity.EffectiveName == "LAD-GOV-MP":
            Gov_insert_cdnt = entity.InsertionPoint
            for Gov_properties in entity.GetDynamicBlockProperties(): #동적블럭 속성 가져오기
                if Gov_properties.propertyname == "@DIST":
                    Gov_y_cdnt = Gov_insert_cdnt[1]+Gov_properties.value
                    Gov_cdnt = (Gov_insert_cdnt[0], Gov_y_cdnt)

    return Gov_cdnt


def Find_Machine_cdnt(Hoistway_top_pnt):
    for entity in doc.ModelSpace:
        if entity.EntityName == "AcDbBlockReference" and entity.EffectiveName == "LAD-TM-GT50-70-LX":
            Machine_x_cdnt = (entity.InsertionPoint[0])
            Machine_y_cdnt = max(Hoistway_top_pnt[1])
            Machine_cdnt = (Machine_x_cdnt, Machine_y_cdnt)

    return Machine_cdnt


def Cal_CP_to_Duct_Hole_dis(Duct_Hole_cdnt, CP_cdnt):

    cal_x_dis = abs(int(Duct_Hole_cdnt[0]) - int(CP_cdnt[0]))
    cal_y_dis = abs(int(Duct_Hole_cdnt[1]) - int(CP_cdnt[1]))

    CP_to_Duct_Hole_dis = cal_x_dis + cal_y_dis

    return CP_to_Duct_Hole_dis


def Cal_CP_to_Gov_dis(Gov_cdnt, CP_cdnt):

    cal_x_dis = abs(int(Gov_cdnt[0]) - int(CP_cdnt[0]))
    cal_y_dis = abs(int(Gov_cdnt[1]) - int(CP_cdnt[1]))

    CP_to_Gov_dis = cal_x_dis + cal_y_dis

    return CP_to_Gov_dis

def Cal_CP_to_Machine_dis(Gov_cdnt, CP_cdnt):

    cal_x_dis = abs(int(Gov_cdnt[0]) - int(CP_cdnt[0]))
    cal_y_dis = abs(int(Gov_cdnt[1]) - int(CP_cdnt[1]))

    CP_to_Gov_dis = cal_x_dis + cal_y_dis

    return CP_to_Gov_dis


Duct_Hole_cdnt = Find_Duct_Hole_cdnt(*Find_Duct_Hole_pst(), *Set_Hoistway_Dim_cdnt())
CP_cdnt = Find_CP_Cdnt()
Gov_cdnt = Find_Gov_cdnt()
Machine_cdnt = Find_Machine_cdnt(Set_Hoistway_Dim_cdnt())
CP_TO_DUCT_HOLE = Cal_CP_to_Duct_Hole_dis(Duct_Hole_cdnt, CP_cdnt)
CP_TO_GOV = Cal_CP_to_Gov_dis(Gov_cdnt, CP_cdnt)
CP_TO_MACHINE = Cal_CP_to_Machine_dis(Machine_cdnt, CP_cdnt)
print("CP TO DUCT :", CP_TO_DUCT_HOLE)
print("CP TO GOV :", CP_TO_GOV)
print("CP TO MACHINE :", CP_TO_MACHINE)

