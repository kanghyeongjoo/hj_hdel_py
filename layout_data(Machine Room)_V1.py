import win32com.client
import math
import string
import fnmatch

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

def Find_Hoistway_cdnt(ent):
    hoistway_x_cdnt = int(ent.InsertionPoint[0])
    hoistway_y_cdnt = int(ent.InsertionPoint[1])
    return hoistway_x_cdnt, hoistway_y_cdnt


def Find_Duct_cdnt(ent):
    duct_hole_x_cdnt = int(ent.InsertionPoint[0])
    duct_hole_y_cdnt = int(ent.InsertionPoint[1])
    return duct_hole_x_cdnt, duct_hole_y_cdnt


def Find_Duct_Hole_pst(hoistway_x_cdnt,hoistway_y_cdnt,duct_hole_x_cdnt,duct_hole_y_cdnt ):

    if hoistway_x_cdnt < duct_hole_x_cdnt:
        duct_hole_x_pst = "Right"
    else:
        duct_hole_x_pst = "Left"

    if hoistway_y_cdnt < duct_hole_y_cdnt:
        duct_hole_y_pst = "Top"
    else:
        duct_hole_y_pst = "Bottom"

    return duct_hole_x_pst, duct_hole_y_pst


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


def Set_Hoistway_Dim_cdnt():

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


def Find_CP_Cdnt():

    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbAttributeDefinition' and entity.TagString == "@TEXT":
            CP_cdnt = (entity.TextAlignmentPoint[0], entity.TextAlignmentPoint[1])
    return CP_cdnt


def Find_Gov_cdnt(ent):

    Gov_insert_cdnt = ent.InsertionPoint
    for Gov_properties in ent.GetDynamicBlockProperties(): #동적블럭 속성 가져오기
        if Gov_properties.propertyname == "@DIST":
            Gov_y_cdnt = Gov_insert_cdnt[1]+Gov_properties.value
            Gov_cdnt = (Gov_insert_cdnt[0], Gov_y_cdnt)
    return Gov_cdnt


def Find_Machine_cdnt(Machine_x_cdnt, Hoistway_top_pnt):

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


def Cal_Machineroom_cdnt():

    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-HOISTWAY-MP-AC-TYPE3-H":
            Hoistway_cdnt = Find_Hoistway_cdnt(entity)
        elif entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-HBTN-HOLE":
            Duct_hole_cdnt = Find_Duct_cdnt(entity)
        elif entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-HOISTWAY-MP-AC-TYPE3":
            entity.Explode()
            Hoistway_size = Set_Hoistway_Dim_cdnt()
        elif entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-CP-MP-L":
            entity.Explode()
            CP_cdnt = Find_CP_Cdnt()
        elif entity.EntityName == "AcDbBlockReference" and entity.EffectiveName == "LAD-GOV-MP":
            Gov_cdnt = Find_Gov_cdnt(entity)
        elif entity.EntityName == "AcDbBlockReference" and entity.EffectiveName == "LAD-TM-GT50-70-LX":
            Machine_x_cdnt = (entity.InsertionPoint[0])
        elif entity.EntityName == "AcDbBlockReference" and entity.EffectiveName == "LAD-TM-GT101C":
            Machine_x_cdnt = (entity.InsertionPoint[0])

    Duct_hole_pst = Find_Duct_Hole_pst(*Hoistway_cdnt, *Duct_hole_cdnt)
    Duct_Hole_cdnt = Find_Duct_Hole_cdnt(*Duct_hole_pst, *Hoistway_size)
    Machine_cdnt = Find_Machine_cdnt(Machine_x_cdnt, Hoistway_size)

    CP_TO_DUCT_HOLE = Cal_CP_to_Duct_Hole_dis(Duct_Hole_cdnt, CP_cdnt)
    CP_TO_GOV = Cal_CP_to_Gov_dis(Gov_cdnt, CP_cdnt)
    CP_TO_MACHINE = Cal_CP_to_Machine_dis(Machine_cdnt, CP_cdnt)

    print(Gov_cdnt, Machine_cdnt)

    return CP_TO_DUCT_HOLE, CP_TO_GOV, CP_TO_MACHINE

Cal_Mchineroom_data = Cal_Machineroom_cdnt()
print("CP TO DUCT :", Cal_Mchineroom_data[0])
print("CP TO GOV :", Cal_Mchineroom_data[1])
print("CP TO MACHINE :", Cal_Mchineroom_data[2])
