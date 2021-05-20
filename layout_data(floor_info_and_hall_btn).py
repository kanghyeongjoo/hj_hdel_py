import win32com.client
import re
import pandas as pd


acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument
layout_kind = re.findall("(\w)[.]DWG", doc.Name.upper())
if "E" not in layout_kind:
    for doc in acad.Documents:
        layout_kind = re.findall("(\w)[.]DWG", doc.Name.upper())
        if "E" in layout_kind:
            doc.Activate()
            doc = acad.ActiveDocument
        else:
            print("ok")
            open() # 현장번호에 맞춰 폴더로 접근하고, 파일명을 get하고 open하는데 오류가 생기면... 경로를 input해라
            #input("출입구 의장도를 찾을 수 없습니다. 경로를 지정하세요.")


def test_df():
    #floor_list = [['B2', '3200'], ['B1', '5250'], ['1', '5500'], ['2', '4530'], ['3', '4220'], ['4', '4220'], ['5', '4220'], ['6', '4220'], ['7', '4220'], ['8', '5400'], ['PH', '4400']]
    floor_list = [['1', '3600'], ['2', '2700'], ['3', '2700'], ['4', '2700'], ['5', '2700'], ['6', '2700'], ['7', '2700'],
     ['8', '2700'], ['PH', '4300']]
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
                        for fl_idx in range(len(floor_data)):
                            if app_floor == floor_data.loc[fl_idx, "층표기"]:
                                floor_data.loc[fl_idx, "홀버튼"] = btn_type
                                floor_data.loc[fl_idx, "위치"] = "기준층"
                    elif "기타층" in textstring:
                        for fl_idx in range(len(floor_data)):
                            if floor_data.loc[fl_idx, "위치"] == "기타층":
                                floor_data.loc[fl_idx, "홀버튼"] = btn_type
                    elif "최상층" in textstring:
                        floor_data.loc[len(floor_data)-1, "홀버튼"] = btn_type
                        floor_data.loc[len(floor_data)-1, "위치"] = "최상층"
                    else:
                        app_floors = []
                        spc_chr = re.findall("\W", textstring)
                        if len(spc_chr) == 0:
                            app_floors.append(re.findall("(\w+)층", textstring)[0])
                        else:
                            spc_app_floor = re.split("[,.]", textstring.replace("층", ""))
                            for chk_floor in spc_app_floor:
                                if "~" not in chk_floor and "-" not in chk_floor:
                                    app_floors.append(chk_floor)
                                else:
                                    st_end_no = re.findall("\d+", chk_floor)
                                    st_no = int(st_end_no[0])
                                    end_no = int(st_end_no[1]) + 1
                                    for floor_ord in range(st_no, end_no):
                                        app_floors.append(str(floor_ord))
                        for app_floor in app_floors:
                             for fl_idx in range(len(floor_data)):
                                 if app_floor == floor_data.loc[fl_idx, "층표기"]:
                                     floor_data.loc[fl_idx, "홀버튼"] = btn_type

    return floor_data

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