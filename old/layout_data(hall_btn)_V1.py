import win32com.client
import math
import string
import fnmatch
import re
import pandas as pd

acad = win32com.client.Dispatch("AutoCAD.Application")

for doc in acad.Documents:
    layout_kind = re.findall("(\w)[.]DWG", doc.Name.upper())
    if "E" in layout_kind:
        doc.Activate()
doc = acad.ActiveDocument

def test_df():
    #floor_list = [['B2', '3200'], ['B1', '5250'], ['1', '5500'], ['2', '4530'], ['3', '4220'], ['4', '4220'], ['5', '4220'], ['6', '4220'], ['7', '4220'], ['8', '5400'], ['PH', '4400']]
    floor_list = [['1', '3600'], ['2', '2700'], ['3', '2700'], ['4', '2700'], ['5', '2700'], ['6', '2700'], ['7', '2700'],['8', '2700'], ['PH', '4300']]
    df_floor_data = pd.DataFrame(floor_list, columns=["층표기", "층고"])

    return df_floor_data

def get_hbtn(floor_data):
    jamb_dict={}
    floor_data["홀버튼"] = ""
    floor_data["위치"] = "기타층"
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TITLE":
            dwg_boundary = (238 * 388) * entity.XEffectiveScaleFactor # 도면 경계 SIZE * 축적
            btn_type = get_btn_type(entity.InsertionPoint, dwg_boundary)
            for att in entity.GetAttributes():
                if att.tagstring == "@TITLE-B":
                    textstring = att.textstring.replace(" ", "")
                    if "기준층" in textstring:
                        app_floor = re.findall("기준층.?(\w+)층", textstring)[0]
                        fl_idx = floor_data.index[floor_data["층표기"] == app_floor]
                        floor_data.loc[fl_idx, "홀버튼"] = btn_type
                        floor_data.loc[fl_idx, "위치"] = "기준층"
                    elif "기타층" in textstring:
                        fl_idx = floor_data.index[floor_data["위치"] == "기타층"]
                        floor_data.loc[fl_idx, "홀버튼"] = btn_type
                    elif "최상층" in textstring:
                        floor_data.loc[floor_data.index[-1], "홀버튼"] = btn_type
                        floor_data.loc[floor_data.index[-1], "위치"] = "최상층"
                    else:
                        textstring = textstring.replace("층", "")
                        floor_list = special_str_split(textstring)
                        for app_floor in floor_list:
                            fl_idx = floor_data.index[floor_data["층표기"] == app_floor]
                            floor_data.loc[fl_idx, "홀버튼"] = btn_type
    return floor_data

def special_str_split(textstring):
    floor_list = []
    if "," not in textstring and "." not in textstring:
        comma_split_floor = [textstring]
    elif "," in textstring or "." in textstring:
        comma_split_floor = re.split("[,.]", textstring)

    for split_floor in comma_split_floor:
        if "~" not in split_floor and "-" not in split_floor:
            floor_list.append(split_floor)
        elif "~" in split_floor or "-" in split_floor:
            str_floor = re.findall("(\w+)\W", split_floor)[0]
            str_text = re.findall("(\D+)\d+", str_floor)
            end_floor = re.findall("\W(\w+)", split_floor)[0]
            end_text = re.findall("(\D+)\d+", end_floor)
            if len(str_text) == 0: # start 층표기에 B2~3과 같은 문자가 있는지 확인
                st_no = int(str_floor)
                end_no = int(end_floor) + 1
                for floor in range(st_no, end_no):
                    floor_list.append(str(floor))
            elif len(end_text) == 0: #start 층표기에는 문자가 있고, end 층표기에는 문자가 없을 때
                text = str_text[0]
                st_no = re.findall("\d+", str_floor)[0]
                st_no = int(st_no)
                end_no = int(end_floor) + 1
                for floor in range(st_no, 0, -1):
                    floor_list.append(text+str(floor))
                for floor in range(1, end_no):
                    floor_list.append(str(floor))
            elif len(end_text) > 0: #start, end 모두 층표기에 문자가 있을 때
                text = str_text[0]
                st_no = re.findall("\d+", str_floor)[0]
                st_no = int(st_no)
                end_no = re.findall("\d+", end_floor)[0]
                end_no = int(end_no) - 1
                for floor in range(st_no, end_no, -1):
                    floor_list.append(text+str(floor))

    return floor_list

def get_btn_type(titile_pnt, btn_min_gap):
    tit_x, tit_y, tit_z = titile_pnt
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and "LAD-HBTN" in entity.EffectiveName:
            btn_x, btn_y, btn_z = entity.InsertionPoint
            tit_btn_gap = abs(tit_x-btn_x)+abs(tit_y+btn_y)
            if btn_min_gap > tit_btn_gap:
                btn_min_gap = tit_btn_gap
                if "SMALL" in entity.EffectiveName.upper():
                    btn_type = "HPB"
                elif "LARGE" in entity.EffectiveName.upper():
                    btn_type = "HIP"

    return btn_type

print(get_hbtn(test_df()))