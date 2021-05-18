import win32com.client
import re

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

def get_property():
    att_dict={}
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TITLE":
            dwg_boundary = (238 * 388) * entity.XEffectiveScaleFactor # 도면 경계 SIZE * 축적
            jamb, btn = get_jamb_btn(entity.InsertionPoint, dwg_boundary, dwg_boundary)
            print(jamb, btn)
            for att in entity.GetAttributes():
                tagstring = att.tagstring
                textstring = att.textstring
                if tagstring == "@TITLE-T":
                    jamb_type = re.findall("\d+", textstring)
                elif tagstring == "@TITLE-B":
                    app_floor = split_floor(textstring)# JAMB 적용층 표기 분리 작업 필요!!!!
                    print(jamb_type, app_floor)
        att_list=[att_dict]
    return att_list

def get_jamb_btn(titile_pnt, jamb_min_gap, btn_min_gap):
    tit_x, tit_y, tit_z = titile_pnt
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and "LAD-DOOR-JAMB" in entity.EffectiveName:
            # for att in entity.GetDynamicBlockProperties():
            #     if att.propertyname == "@HH":
            #         print("HH, ", att.value)
            jamb_x, jamb_y, jamb_z = entity.InsertionPoint
            tit_jamb_gap = abs(tit_x - jamb_x)+abs(tit_y - jamb_y)
            if jamb_min_gap > tit_jamb_gap:
                jamb_min_gap = tit_jamb_gap
                jamb_type = entity.EffectiveName
        elif entity.EntityName == 'AcDbBlockReference' and "LAD-HBTN" in entity.EffectiveName:
            btn_x, btn_y, btn_z = entity.InsertionPoint
            tit_btn_gap = abs(tit_x-btn_x)+abs(tit_y+btn_y)
            if btn_min_gap > tit_btn_gap:
                btn_min_gap = tit_btn_gap
                if "SMALL" in entity.EffectiveName.upper():
                    btn_type = "HPB"
                elif "LARGE" in entity.EffectiveName.upper():
                    btn_type = "HIP"

    return jamb_type, btn_type



def split_floor(bf_floor):
    floor_list = ["1","2","3","4","5","6","7"]

    if "기준층" in bf_floor:
        app_floor = re.findall("기준층.?(\w+)층", bf_floor)
    elif "기타층" in bf_floor:
        app_floor = ["기타층"]
    elif "최상층" in bf_floor:
        app_floor = ["최상층"]
    else:
        app_floor = re.findall(r"\w+\b|\w+(?=층)", bf_floor)

    return app_floor

print(get_property())