import win32com.client
import math
import string
import fnmatch
import re
import pandas as pd

# acad = win32com.client.Dispatch("AutoCAD.Application")

# for doc in acad.Documents:
#     layout_kind = re.findall("(\w)[.]DWG", doc.Name.upper())
#     if "S" in layout_kind:
#         doc.Activate()
# doc = acad.ActiveDocument
#
# def get_table_cdnt():
#
#     table_blo_y_cdnt = None
#     for entity in doc.ModelSpace: # 층고 테이블 분해
#         if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TABLE-FLOOR-HEIGHT":
#             table_blo_y_cdnt = entity.InsertionPoint[1]
#             entity.Explode()
#
#     if table_blo_y_cdnt != None:
#         for entity in doc.ModelSpace:  # 테이블 좌표 구하기
#             if entity.EntityName == 'AcDbPolyline' and entity.Coordinates[1] == table_blo_y_cdnt:
#                 start_x_cdnt = entity.Coordinates[0]
#                 end_x_cdnt = entity.Coordinates[2]
#                 start_y_cdnt = entity.Coordinates[1]
#                 end_y_cdnt = entity.Coordinates[5]
#
#     elif table_blo_y_cdnt == None:
#         for entity in doc.ModelSpace:
#             if entity.EntityName == 'AcDbText' and entity.TextString == "FL / ST":
#                 text_base_y_cdnt = entity.TextAlignmentPoint[1]
#
#         for entity in doc.ModelSpace:
#             if entity.EntityName == 'AcDbText' and entity.TextString == "층" and entity.TextAlignmentPoint[1] == text_base_y_cdnt:
#                 text_base_x_cdnt = entity.TextAlignmentPoint[0]
#
#         cal_x_gap = {}
#         cal_y_gap = {}
#         for entity in doc.ModelSpace:
#             if entity.EntityName == 'AcDbPolyline' and entity.Layer == "LAD-OUTLINE":  # 대신 방화도어 table 하고 겹치지 않도록 별도의 좌표 비교값 넣기
#                 x_gap = abs(text_base_x_cdnt - entity.Coordinates[0])
#                 y_gap = abs(text_base_y_cdnt - entity.Coordinates[1])
#                 cal_x_gap.update({x_gap: entity.Coordinates[0]})
#                 cal_y_gap.update({y_gap: entity.Coordinates[1]})
#
#         cal_x_gap = sorted(cal_x_gap.items())
#         cal_y_gap = sorted(cal_y_gap.items())
#
#         table_x_cdnt = cal_x_gap[0][1]
#         table_y_cdnt = cal_y_gap[0][1]
#
#         for entity in doc.ModelSpace:  # 테이블 좌표 구하기
#             if entity.EntityName == 'AcDbPolyline' and entity.Layer == "LAD-OUTLINE" and entity.Coordinates[0] == table_x_cdnt and entity.Coordinates[1] == table_y_cdnt:
#                 start_x_cdnt = entity.Coordinates[0]
#                 end_x_cdnt = entity.Coordinates[2]
#                 start_y_cdnt = entity.Coordinates[1]
#                 end_y_cdnt = entity.Coordinates[5]
#
#     return start_x_cdnt, end_x_cdnt, start_y_cdnt, end_y_cdnt
#
# def get_fl_and_fl_h(s_x_cdnt, e_x_cdnt, s_y_cdnt, e_y_cdnt):
#
#     table_datas_list = {}
#     for entity in doc.ModelSpace:  # 테이블에 있는 모든 TEXT는 정렬좌표와 Dictionary로 get
#         if entity.EntityName == 'AcDbText':
#             x_cdnt = entity.InsertionPoint[0]
#             y_cdnt = entity.InsertionPoint[1]
#             if x_cdnt > s_x_cdnt and x_cdnt < e_x_cdnt and y_cdnt < s_y_cdnt and y_cdnt > e_y_cdnt:
#                 table_datas_list.update({entity.TextAlignmentPoint: entity.TextString})
#                 str_text = entity.TextString
#                 if str_text == "층": # 윗행(층)과 아래행(층고) 나누는 기준
#                     floor_row_y_cdnt = entity.TextAlignmentPoint[1]
#                 elif str_text == "층고":  # 윗행(층)과 아래행(층고) 나누는 기준
#                     floor_height_row_y_cdnt = entity.TextAlignmentPoint[1]
#
#     floor_row_datas = {}
#     floor_height_row_datas = {}
#     for data_cdnt, table_data in table_datas_list.items():
#         if data_cdnt[1] == floor_row_y_cdnt: # 층행
#             floor_row_datas.update({table_data:data_cdnt[0]}) # 층표기 : 좌표(층표기는 고유하지만 층표기 분리 시 세로 좌표와 층고는 중복이 가능)
#         elif data_cdnt[1] == floor_height_row_y_cdnt: # 층고행
#             floor_height_row_datas.update({data_cdnt[0]:table_data}) # 좌표 : 층고
#
#     floors_data_with_x_cdnt = {}
#     for floor_data, floor_data_cdnt in floor_row_datas.items(): # 층행 Dictionary
#         if floor_data != "층" and floor_data != "FL / ST":
#             for floor_height_data_cdnt, floor_height_data in floor_height_row_datas.items(): # 층고행 Dictionary
#                 if floor_data_cdnt == floor_height_data_cdnt: # 가로 좌표 비교하여 층표시와 층고 Matching
#                     floors_data_with_x_cdnt.update({floor_data_cdnt:{floor_data:floor_height_data}})
#
#     floors_data_with_x_cdnt = sorted(floors_data_with_x_cdnt.items()) # 가로 좌표 기준 정렬(단, 층표기 분리시 세로 좌표가 동일하므로 세로 층표기를 정렬이후에 나누는 것도 고려해볼 것)
#
#     floor_and_floor_height={}
#     for floor_data_with_x_cdnt in floors_data_with_x_cdnt:
#         floor_data = floor_data_with_x_cdnt[1]
#         for floor_mark, floor_height in floor_data.items():
#             floor_and_floor_height.update({floor_mark:floor_height})
#
#     return floor_and_floor_height
#
#
# def special_str_split(floor_and_floor_height):
#     comma_split = {}
#     tilde_split = []
#     for before_floor, height in floor_and_floor_height.items():
#         if "," not in before_floor and "." not in before_floor:
#             comma_split.update({before_floor: height})
#         elif "," in before_floor or "." in before_floor:
#             comma_split_list = re.split("[,.]", before_floor)
#             for split_floor in comma_split_list:
#                 comma_split.update({split_floor: height})
#
#     for before_floor, height in comma_split.items():
#         if "~" not in before_floor and "-" not in before_floor:
#             tilde_split.append([before_floor, height])
#         elif "~" in before_floor or "-" in before_floor:
#             st_end_no = re.findall("\d+", before_floor)
#             st_no = int(st_end_no[0])
#             end_no = int(st_end_no[1]) + 1
#             for tilde_floor in range(st_no, end_no):
#                 tilde_split.append([str(tilde_floor), height])
#
#     return tilde_split
#
# fl_and_fl_h = get_fl_and_fl_h(*get_table_cdnt())
# final_floor_data = special_str_split(fl_and_fl_h)
# print(final_floor_data)
#
# df_floor_data = pd.DataFrame(final_floor_data, columns=["층표기", "층고"])
# print(df_floor_data)
# doc.close()

#fl / st도 구하기(함수 나누기)#and floor_name != "FL / ST" 이거 어떻게 할꺼??

acad = win32com.client.Dispatch("AutoCAD.Application")

for doc in acad.Documents:
    layout_kind = re.findall("(\w)[.]DWG", doc.Name.upper())
    if "E" in layout_kind:
        doc.Activate()
doc = acad.ActiveDocument

def test_df():
    #floor_list = [['B2', '3200'], ['B1', '5250'], ['1', '5500'], ['2', '4530'], ['3', '4220'], ['4', '4220'], ['5', '4220'], ['6', '4220'], ['7', '4220'], ['8', '5400'], ['PH', '4400']]
    #floor_list = [['1', '3600'], ['2', '2700'], ['3', '2700'], ['4', '2700'], ['5', '2700'], ['6', '2700'], ['7', '2700'],['8', '2700'], ['PH', '4300']]
    floor_list = [['B1', '6000'], ['1', '5600'], ['2', '5100'], ['3', '5100'], ['4', '5100'], ['5', '5100'], ['6', '4500']]
    df_floor_data = pd.DataFrame(floor_list, columns=["층표기", "층고"])

    return df_floor_data

def get_hbtn(floor_data):
    jamb_dict={}
    floor_data[["JAMB", "JAMB_SPEC","홀버튼","위치"]] = ""
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TITLE":
            dwg_boundary = (238 * 388) * entity.XEffectiveScaleFactor # 도면 경계 SIZE * 축적
            btn_type = get_btn_type(entity.InsertionPoint, dwg_boundary)
            for att in entity.GetAttributes():
                if att.tagstring == "@TITLE-T":
                    app_jamb, jamb_spec = get_jamb_type(entity.InsertionPoint, dwg_boundary, att.textstring)
            for att in entity.GetAttributes():
                if att.tagstring == "@TITLE-B":
                    textstring = att.textstring.replace(" ", "")
                    if "기준층" in textstring:
                        app_floor = re.findall("기준층.?(\w+)층", textstring)[0]
                        fl_idx = floor_data.index[floor_data["층표기"] == app_floor] # 빈값을 넣고, 비어 있으면 기타층으론 넣는다
                        floor_data.loc[fl_idx, "홀버튼"] = btn_type
                        floor_data.loc[fl_idx, "JAMB"] = app_jamb
                        floor_data.loc[fl_idx, "JAMB_SPEC"] = jamb_spec
                        floor_data.loc[fl_idx, "위치"] = "기준층"
                    elif "기타층" in textstring:
                        fl_idx = floor_data.index[(floor_data["위치"] == "")]
                        floor_data.loc[fl_idx, "홀버튼"] = btn_type
                        floor_data.loc[fl_idx, "JAMB"] = app_jamb
                        floor_data.loc[fl_idx, "JAMB_SPEC"] = jamb_spec
                        floor_data.loc[fl_idx, "위치"] = "기타층"
                    elif "최상층" in textstring:
                        floor_data.loc[floor_data.index[-1], "홀버튼"] = btn_type
                        floor_data.loc[floor_data.index[-1], "위치"] = "최상층"
                        floor_data.loc[floor_data.index[-1], "JAMB"] = app_jamb
                        floor_data.loc[floor_data.index[-1], "JAMB_SPEC"] = jamb_spec
                    else:
                        textstring = textstring.replace("층", "")
                        floor_list = special_str_split(textstring)
                        for app_floor in floor_list:
                            fl_idx = floor_data.index[floor_data["층표기"] == app_floor]
                            floor_data.loc[fl_idx, "JAMB"] = app_jamb
                            floor_data.loc[fl_idx, "JAMB_SPEC"] = jamb_spec
                            floor_data.loc[fl_idx, "홀버튼"] = btn_type
                            floor_data.loc[fl_idx, "위치"] = "기타층"

    main_fl_idx = floor_data.index[(floor_data["위치"] == "기준층")][0]
    if main_fl_idx > 0:
        floor_data.loc[:main_fl_idx - 1, "위치"] = "지하층"

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

def get_jamb_type(titile_pnt, jamb_min_gap, textstring):
    tit_x, tit_y, tit_z = titile_pnt
    jamb_ord = 0
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and "LAD-DOOR-JAMB" in entity.EffectiveName:
            jamb_x, jamb_y, jamb_z = entity.InsertionPoint
            tit_jamb_gap = abs(tit_x - jamb_x)+abs(tit_y - jamb_y)
            if jamb_min_gap > tit_jamb_gap:
                jamb_min_gap = tit_jamb_gap
                if "CP" in entity.EffectiveName.upper():
                    jamb_spec = "CP" + re.findall("\d+", textstring)[0]
                    app_jamb = "JAMB(CP);"
                else:
                    jamb_spec = "JP" + re.findall("\d+", textstring)[0]
                    jamb_ord = jamb_ord + 1
                    app_jamb = "JAMB(" + str(jamb_ord) + ");"

    return app_jamb, jamb_spec

print(get_hbtn(test_df()))