import pandas as pd
import win32com.client
import re
import time
import math
import requests
from bs4 import BeautifulSoup

start = time.time()


def get_entity(layout_kind):
    acad = win32com.client.Dispatch("AutoCAD.Application")
    global doc, ent_group
    doc = acad.ActiveDocument
    if layout_kind == "M":
        ent_blo_name = ["LAD-HOLE", "U107", "LAD-HBTN-HOLE", "LAD-CP", "LAD-GOV", "LAD-TM", "DIM_ENT", "hst_cent", "hst_hole_cent", "door_cdnt"]
        ent_group = dict.fromkeys(ent_blo_name)
        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbBlockReference':
                if "LAD-HOISTWAY" in entity.EffectiveName and entity.EffectiveName[-1] == "H":
                    ent_group.update({"hst_hole_cent": entity.InsertionPoint})
                    entity.explode()
                elif "LAD-HOISTWAY" in entity.EffectiveName:
                    ent_group.update({"hst_cent": entity.InsertionPoint})
                    entity.explode()
                elif entity.EffectiveName == "A$C4bb02043":
                    entity.explode()
                else:
                    for ent_name in ent_blo_name:
                        if ent_name in entity.EffectiveName:
                            if ent_group[ent_name] == None:
                                ent_group[ent_name] = []
                                ent_group[ent_name].append(entity)
                            else:
                                ent_group[ent_name].append(entity)

        for entity in doc.ModelSpace:
            if entity.EntityName == "AcDbRotatedDimension" and entity.TextOverride != "":
                if ent_group["DIM_ENT"] == None:
                    ent_group["DIM_ENT"] = []
                    ent_group["DIM_ENT"].append(entity)
                else:
                    ent_group["DIM_ENT"].append(entity)
            elif entity.EntityName == "AcDbText" and "기계실 출입문" in entity.TextString:
                ent_group.update({"door_cdnt":entity.InsertionPoint})
            elif ent_group["LAD-CP"] == None and entity.EntityName == 'AcDbBlockReference' and "LAD-CP" in entity.EffectiveName:
                ent_group.update({"LAD-CP":[entity]})

        if ent_group["hst_cent"] == None and ent_group["LAD-TM"] != None:
            hst_cent = ent_group["LAD-TM"][0].InsertionPoint
            ent_group.update({"hst_cent": hst_cent})
        else:
            print("hoistway 중심점을 확인할 수 없습니다.")  # 이 부분은 나중에 직접 입력 또는 CAD에서 직접 선택할 수 있도록 한다.

        if ent_group["hst_hole_cent"] == None and ent_group["LAD-HOLE"] != None:
            hst_hole_cent = ent_group["LAD-HOLE"][0].InsertionPoint
            ent_group.update({"hst_hole_cent": hst_hole_cent})
        elif ent_group["hst_hole_cent"] == None and ent_group["U107"] != None:
            hst_hole_cent = ent_group["U107"][0].InsertionPoint
            ent_group.update({"hst_hole_cent": hst_hole_cent})
        else:
            print("hoistway hole 중심점을 확인할 수 없습니다.")  # 이 부분은 나중에 직접 입력 또는 CAD에서 직접 선택할 수 있도록 한다.

    return ent_group


def get_machine_room_data(entity):
    doc.SendCommand('setxdata ')
    m_room_data = {}

    for dim_ent in ent_group["DIM_ENT"]:
        del_s = dim_ent.TextOverride.replace(" ", "")
        size_name = re.findall("[가-힣]+", del_s)[0]
        Xdata = dim_ent.GetXData("", "Type", "Data")
        pt1 = Xdata[1][-2]
        pt2 = Xdata[1][-1]
        if int(pt1[0]) == int(pt2[0]):
            size_name = size_name + "(세로)"
        elif int(pt1[1]) == int(pt2[1]):
            size_name = size_name + "(가로)"
        else:
            gaps = {}
            gaps.update({abs(int(pt1[0]) - int(pt2[0])): "(가로)"})
            gaps.update({abs(int(pt1[1]) - int(pt2[1])): "(세로)"})
            size_name = size_name + gaps[round(dim_ent.Measurement)]

        size = round(dim_ent.Measurement)
        if size_name == "승강로내부(가로)":
            hoist_lft_x = min(int(pt1[0]), int(pt2[0]))
            if int(ent_group["hst_cent"][0]) < hoist_lft_x:
                size_name = "EL_EHH_CHK"
            else:
                size_name = "EL_EHH"
        elif size_name == "승강로내부(세로)":
            ver_dim_x = int(pt1[0])
            if ver_dim_x < int(ent_group["hst_cent"][0]):
                size_name = "EL_EHV"
            elif ver_dim_x > int(ent_group["hst_hole_cent"][0]):
                size_name = "EL_EHV_CHK"
            else:
                gap_hoistway = abs(ver_dim_x - int(ent_group["hst_cent"][0]))
                gap_hoistway_hole = abs(ver_dim_x - int(ent_group["hst_hole_cent"][0]))
                if min(gap_hoistway, gap_hoistway_hole) == gap_hoistway:
                    size_name = "EL_EHV"
                else:
                    size_name = "EL_EHV_CHK"
        m_room_data.update({size_name: size})

    if m_room_data["EL_EHH"] == m_room_data["EL_EHH_CHK"]:
        del m_room_data["EL_EHH_CHK"]
    if m_room_data["EL_EHV"] == m_room_data["EL_EHV_CHK"]:
        del m_room_data["EL_EHV_CHK"]

    if ent_group["LAD-CP"] != None:
        cp_ent = ent_group["LAD-CP"][0]
        for cp_att in cp_ent.GetAttributes():
            if cp_att.TagString == "@TEXT":
                cp_cdnt = cp_att.TextAlignmentPoint

        if ent_group["LAD-HBTN-HOLE"] != None:
            duct_ent = ent_group["LAD-HBTN-HOLE"][0]
            hole_ent_x_gap = int(ent_group["hst_cent"][0] - ent_group["hst_hole_cent"][0])
            duct_x = int(duct_ent.InsertionPoint[0]) + hole_ent_x_gap
            cp_duct_x = abs(int(cp_cdnt[0]) - duct_x)
            cp_duct_y = abs(int(cp_cdnt[1] - duct_ent.InsertionPoint[1]))
            cp_to_duct = round(cp_duct_x + cp_duct_y, -3) + 1250
            m_room_data.update({"EL_EDTA":cp_to_duct})

        if ent_group["LAD-TM"] != None:
            tm_ent = ent_group["LAD-TM"][0]
            for tm_prt in tm_ent.GetDynamicBlockProperties():
                if tm_prt.propertyname == "@PP":
                    m_room_data.update({"EL_EPPY":int(tm_prt.value)})
                    break
            cp_tm_x = abs(int(cp_cdnt[0] - tm_ent.InsertionPoint[0]))
            cp_tm_y = abs(int(cp_cdnt[1] - tm_ent.InsertionPoint[1]))
            cp_to_tm = round(cp_tm_x + cp_tm_y + 1500, -3) + 1000
            m_room_data.update({"EL_EDTB": cp_to_tm})

        if ent_group["door_cdnt"] != None:
            cp_door_x = abs(int(cp_cdnt[0] - ent_group["door_cdnt"][0]))
            cp_door_y = abs(int(cp_cdnt[1] - ent_group["door_cdnt"][1]))
            cp_to_pwr = round(cp_door_x + cp_door_y + 1650, -3) + 1000
            m_room_data.update({"EL_EDTC":cp_to_pwr})

        if ent_group["LAD-GOV"] != None:
            for gov_ent in ent_group["LAD-GOV"]:
                if gov_ent.EffectiveName[-1] == "H":
                    for gov_h_prt in gov_ent.GetDynamicBlockProperties():
                        if gov_h_prt.propertyname == "@DIST":
                            gov_name = "EL_ECGV_CHK"
                            gov_spec = "DG" + str(int(gov_h_prt.value))
                            break
                else:
                    gov_y_cdnt = gov_ent.InsertionPoint[1]
                    for gov_prt in gov_ent.GetDynamicBlockProperties():
                        if gov_prt.propertyname == "@DIST":
                            gov_name = "EL_ECGV"
                            gov_spec = "DG" + str(int(gov_prt.value))
                            break
                    cp_gov_x = abs(int(cp_cdnt[0] - gov_ent.InsertionPoint[0]))
                    cp_gov_y = abs(int(cp_cdnt[1] - (gov_y_cdnt + gov_prt.value)))
                    gov_cc = abs(int(ent_group["hst_cent"][1] - gov_ent.InsertionPoint[1]))
                    cp_to_gov = round(cp_gov_x + cp_gov_y + 150, -3) + 1000
                    m_room_data.update({"EL_EDTE":cp_to_gov, "EL_ECCC":gov_cc})
                m_room_data.update({gov_name:gov_spec})

        if m_room_data["EL_ECGV"] == m_room_data["EL_ECGV_CHK"]:
            del m_room_data["EL_ECGV_CHK"]

    return m_room_data

entity = get_entity("M")
data = get_machine_room_data(entity)
print(data)
print("걸린 시간 : ", time.time() - start)
