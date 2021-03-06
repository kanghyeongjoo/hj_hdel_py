import win32com.client
import re

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

def get_property():
    att_dict={}
    jamb_ord = 0
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TITLE":
            dwg_boundary = (238 * 388) * entity.XEffectiveScaleFactor # 도면 경계 SIZE * 축적
            jamb_ord = jamb_ord + 1
            jamb_kind, btn_type = get_jamb_btn(entity.InsertionPoint, dwg_boundary, dwg_boundary)
            print(btn_type)
            for att in entity.GetAttributes():
                if att.tagstring == "@TITLE-T" and jamb_kind == "GENERAL":
                    jamb_spec = "JP" + re.findall("\d+", att.textstring)[0]
                    print(jamb_spec)
                elif att.tagstring == "@TITLE-T" and jamb_kind == "CP_JAMB":
                    jamb_spec == "CP"+ re.findall("\d+", att.textstring)[0]
                    print(jamb_spec)
                elif att.textstring == "@TITLE-B":
                    jamb_type, app_floor = split_floor(att.textstring, jamb_kind, jamb_ord)# JAMB 적용층 표기 분리 작업 필요!!!!
                    print(jamb_kind, app_floor)
        att_list=[att_dict]
    return att_list

def get_jamb_btn(titile_pnt, jamb_min_gap, btn_min_gap):
    tit_x, tit_y, tit_z = titile_pnt
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and "LAD-DOOR-JAMB" in entity.EffectiveName:
            jamb_x, jamb_y, jamb_z = entity.InsertionPoint
            tit_jamb_gap = abs(tit_x - jamb_x)+abs(tit_y - jamb_y)
            if jamb_min_gap > tit_jamb_gap:
                jamb_min_gap = tit_jamb_gap
                if "CP" in entity.EffectiveName.upper():
                    jamb_kind = "CP_JAMB"
                else:
                    jamb_kind = "GENERAL"
        elif entity.EntityName == 'AcDbBlockReference' and "LAD-HBTN" in entity.EffectiveName:
            btn_x, btn_y, btn_z = entity.InsertionPoint
            tit_btn_gap = abs(tit_x-btn_x)+abs(tit_y+btn_y)
            if btn_min_gap > tit_btn_gap:
                btn_min_gap = tit_btn_gap
                if "SMALL" in entity.EffectiveName.upper():
                    btn_type = "HPB"
                elif "LARGE" in entity.EffectiveName.upper():
                    btn_type = "HIP"

    return jamb_kind, btn_type



def split_floor(bf_floor, jamb_kind, jamb_ord):

    if "기준층" in bf_floor:
        jamb_type = "JAMB(1);종류"
        app_floor = re.findall("기준층.?(\w+)층", bf_floor)
    elif "기타층" in bf_floor:
        jamb_type = "JAMB(" + str(jamb_ord) + ");종류"
        app_floor = ["기타층"]
    elif "최상층" in bf_floor and jamb_kind == "GENERAL":
        jamb_type = "JAMB(" + str(jamb_ord) + ");종류"
        app_floor = ["최상층"]
    elif "최상층" in bf_floor and jamb_kind == "CP_JAMB":
        jamb_kind = "JAMB(CP);종류"
        app_floor = ["최상층"]
    else:
        jamb_type = "JAMB(" + str(jamb_ord) + ");종류"
        app_floor = re.findall(r"\w+\b|\w+(?=층)", bf_floor)

    return jamb_kind, app_floor

print(get_property())