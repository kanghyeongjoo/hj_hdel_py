import win32com.client
import math
import string
import fnmatch
import re
import pandas as pd

acad = win32com.client.Dispatch("AutoCAD.Application")

for doc in acad.Documents:
    layout_kind = re.findall("(\w)[.]DWG", doc.Name.upper())
    if "S" in layout_kind:
        doc.Activate()
doc = acad.ActiveDocument

def get_table_cdnt():

    table_blo_y_cdnt = None
    for entity in doc.ModelSpace: # 층고 테이블 분해
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TABLE-FLOOR-HEIGHT":
            table_blo_y_cdnt = entity.InsertionPoint[1]
            entity.Explode()

    if table_blo_y_cdnt != None:
        for entity in doc.ModelSpace:  # 테이블 좌표 구하기
            if entity.EntityName == 'AcDbPolyline' and entity.Coordinates[1] == table_blo_y_cdnt:
                start_x_cdnt = entity.Coordinates[0]
                end_x_cdnt = entity.Coordinates[2]
                start_y_cdnt = entity.Coordinates[1]
                end_y_cdnt = entity.Coordinates[5]

    elif table_blo_y_cdnt == None:
        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbText' and entity.TextString == "FL / ST":
                text_base_y_cdnt = entity.TextAlignmentPoint[1]

        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbText' and entity.TextString == "층" and entity.TextAlignmentPoint[1] == text_base_y_cdnt:
                text_base_x_cdnt = entity.TextAlignmentPoint[0]

        cal_x_gap = {}
        cal_y_gap = {}
        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbPolyline' and entity.Layer == "LAD-OUTLINE":  # 대신 방화도어 table 하고 겹치지 않도록 별도의 좌표 비교값 넣기
                x_gap = abs(text_base_x_cdnt - entity.Coordinates[0])
                y_gap = abs(text_base_y_cdnt - entity.Coordinates[1])
                cal_x_gap.update({x_gap: entity.Coordinates[0]})
                cal_y_gap.update({y_gap: entity.Coordinates[1]})

        cal_x_gap = sorted(cal_x_gap.items())
        cal_y_gap = sorted(cal_y_gap.items())

        table_x_cdnt = cal_x_gap[0][1]
        table_y_cdnt = cal_y_gap[0][1]

        for entity in doc.ModelSpace:  # 테이블 좌표 구하기
            if entity.EntityName == 'AcDbPolyline' and entity.Layer == "LAD-OUTLINE" and entity.Coordinates[0] == table_x_cdnt and entity.Coordinates[1] == table_y_cdnt:
                start_x_cdnt = entity.Coordinates[0]
                end_x_cdnt = entity.Coordinates[2]
                start_y_cdnt = entity.Coordinates[1]
                end_y_cdnt = entity.Coordinates[5]

    return start_x_cdnt, end_x_cdnt, start_y_cdnt, end_y_cdnt

def get_fl_and_fl_h(s_x_cdnt, e_x_cdnt, s_y_cdnt, e_y_cdnt):

    table_datas_list = {}
    for entity in doc.ModelSpace:  # 테이블에 있는 모든 TEXT는 정렬좌표와 Dictionary로 get
        if entity.EntityName == 'AcDbText':
            x_cdnt = entity.InsertionPoint[0]
            y_cdnt = entity.InsertionPoint[1]
            if x_cdnt > s_x_cdnt and x_cdnt < e_x_cdnt and y_cdnt < s_y_cdnt and y_cdnt > e_y_cdnt:
                table_datas_list.update({entity.TextAlignmentPoint: entity.TextString})
                str_text = entity.TextString
                if str_text == "층": # 윗행(층)과 아래행(층고) 나누는 기준
                    floor_row_y_cdnt = entity.TextAlignmentPoint[1]
                elif str_text == "층고":  # 윗행(층)과 아래행(층고) 나누는 기준
                    floor_height_row_y_cdnt = entity.TextAlignmentPoint[1]

    floor_row_datas = {}
    floor_height_row_datas = {}
    for data_cdnt, table_data in table_datas_list.items():
        if data_cdnt[1] == floor_row_y_cdnt: # 층행
            floor_row_datas.update({table_data:data_cdnt[0]}) # 층표기 : 좌표(층표기는 고유하지만 층표기 분리 시 세로 좌표와 층고는 중복이 가능)
        elif data_cdnt[1] == floor_height_row_y_cdnt: # 층고행
            floor_height_row_datas.update({data_cdnt[0]:table_data}) # 좌표 : 층고

    floors_data_with_x_cdnt = {}
    for floor_data, floor_data_cdnt in floor_row_datas.items(): # 층행 Dictionary
        if floor_data != "층" and floor_data != "FL / ST":
            for floor_height_data_cdnt, floor_height_data in floor_height_row_datas.items(): # 층고행 Dictionary
                if floor_data_cdnt == floor_height_data_cdnt: # 가로 좌표 비교하여 층표시와 층고 Matching
                    floors_data_with_x_cdnt.update({floor_data_cdnt:{floor_data:floor_height_data}})

    floors_data_with_x_cdnt = sorted(floors_data_with_x_cdnt.items()) # 가로 좌표 기준 정렬(단, 층표기 분리시 세로 좌표가 동일하므로 세로 층표기를 정렬이후에 나누는 것도 고려해볼 것)

    floor_and_floor_height={}
    for floor_data_with_x_cdnt in floors_data_with_x_cdnt:
        floor_data = floor_data_with_x_cdnt[1]
        for floor_mark, floor_height in floor_data.items():
            floor_and_floor_height.update({floor_mark:floor_height})

    return floor_and_floor_height


def special_str_split(floor_and_floor_height):
    comma_split = {}
    tilde_split = []
    for before_floor, height in floor_and_floor_height.items():
        if "," not in before_floor and "." not in before_floor:
            comma_split.update({before_floor: height})
        elif "," in before_floor or "." in before_floor:
            comma_split_list = re.split("[,.]", before_floor)
            for split_floor in comma_split_list:
                comma_split.update({split_floor: height})

    for before_floor, height in comma_split.items():
        if "~" not in before_floor and "-" not in before_floor:
            tilde_split.append([before_floor, height])
        elif "~" in before_floor or "-" in before_floor:

            st_end_no = re.findall("\d+", before_floor) #B2~3과 같은 층표기 정제 필요!!!!!
            st_no = int(st_end_no[0])
            end_no = int(st_end_no[1]) + 1
            for tilde_floor in range(st_no, end_no):
                tilde_split.append([str(tilde_floor), height])

    return tilde_split

fl_and_fl_h = get_fl_and_fl_h(*get_table_cdnt())
final_floor_data = special_str_split(fl_and_fl_h)
print(final_floor_data)

df_floor_data = pd.DataFrame(final_floor_data, columns=["층표기", "층고"])
print(df_floor_data)
# print(df_floor_data.index)

# doc.close()

#fl / st도 구하기(함수 나누기)#and floor_name != "FL / ST" 이거 어떻게 할꺼??