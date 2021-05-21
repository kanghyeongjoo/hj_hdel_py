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
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-FORM-A3-SIMPLE":
            palette_area = (238 * 388) * entity.XEffectiveScaleFactor # palett 면적
        elif entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TABLE-FLOOR-HEIGHT":
            table_blo_y_cdnt = entity.InsertionPoint[1] # 층고 테이블 분해
            entity.Explode()

    if table_blo_y_cdnt != None:
        for entity in doc.ModelSpace:  # 테이블 좌표 구하기
            if entity.EntityName == 'AcDbPolyline' and entity.Coordinates[1] == table_blo_y_cdnt:
                start_x_cdnt = entity.Coordinates[0]
                end_x_cdnt = entity.Coordinates[2]
                start_y_cdnt = entity.Coordinates[1]
                end_y_cdnt = entity.Coordinates[5]

    elif table_blo_y_cdnt == None:
        text_insert_cdnt = {}
        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbText' and entity.TextString == "FL / ST":
                text_base_y_cdnt = entity.TextAlignmentPoint[1]
            elif entity.EntityName == 'AcDbText' and entity.TextString == "층":
                text_insert_cdnt.update({entity.TextAlignmentPoint[1]:entity.TextAlignmentPoint[0]}) # text y:x
        text_base_x_cdnt = text_insert_cdnt.get(text_base_y_cdnt)

        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbPolyline' and entity.Layer == "LAD-OUTLINE":
                x_gap = abs(text_base_x_cdnt - entity.Coordinates[0])
                y_gap = abs(text_base_y_cdnt - entity.Coordinates[1])
                gap  = x_gap + y_gap
                if palette_area > gap: # 방화도어 table 하고 겹치지 않도록 gap 비교
                    palette_area = gap
                    start_x_cdnt = entity.Coordinates[0]
                    end_x_cdnt = entity.Coordinates[2]
                    start_y_cdnt = entity.Coordinates[1]
                    end_y_cdnt = entity.Coordinates[5]

    return start_x_cdnt, end_x_cdnt, start_y_cdnt, end_y_cdnt


def get_floor_data(s_x_cdnt, e_x_cdnt, s_y_cdnt, e_y_cdnt):
    table_datas_list = {}
    for entity in doc.ModelSpace:  # 테이블에 있는 모든 TEXT는 정렬좌표와 Dictionary로 get
        if entity.EntityName == 'AcDbText':
            x_cdnt = entity.InsertionPoint[0]
            y_cdnt = entity.InsertionPoint[1]
            if x_cdnt > s_x_cdnt and x_cdnt < e_x_cdnt and y_cdnt < s_y_cdnt and y_cdnt > e_y_cdnt:
                table_datas_list.update({entity.TextAlignmentPoint: entity.TextString})
                if entity.TextString == "층":  # 윗행(층)과 아래행(층고) 나누는 기준
                    floor_row_y_cdnt = entity.TextAlignmentPoint[1]
                elif entity.TextString == "층고":  # 윗행(층)과 아래행(층고) 나누는 기준
                    floor_height_row_y_cdnt = entity.TextAlignmentPoint[1]

    floor_row_datas = {}
    floor_height_row_datas = {}
    for data_cdnt, table_data in table_datas_list.items():
        if data_cdnt[1] == floor_row_y_cdnt:  # 층행
            floor_row_datas.update({table_data: data_cdnt[0]})  # 층표기 : 좌표(층표기는 고유하지만 층표기 분리 시 세로 좌표와 층고는 중복이 가능)
        elif data_cdnt[1] == floor_height_row_y_cdnt:  # 층고행
            floor_height_row_datas.update({data_cdnt[0]: table_data})  # 좌표 : 층고

    floors_data_with_x_cdnt = {}
    for floor_data, floor_data_cdnt in floor_row_datas.items():  # 층행 Dictionary
        if floor_data != "층" and floor_data != "FL / ST":
            for floor_height_data_cdnt, floor_height_data in floor_height_row_datas.items():  # 층고행 Dictionary
                if floor_data_cdnt == floor_height_data_cdnt:  # 가로 좌표 비교하여 층표시와 층고 Matching
                    floors_data_with_x_cdnt.update({floor_data_cdnt: {floor_data: floor_height_data}})
        elif floor_data == "FL / ST": # 층수 구하기
            for floor_height_data_cdnt, floor_height_data in floor_height_row_datas.items():
                if floor_data_cdnt == floor_height_data_cdnt:
                    total_floor = re.findall("(\d+)/", floor_height_data)[0]
                    stop_floor = re.findall("/(\d+)", floor_height_data)[0]

    floors_data_with_x_cdnt = sorted(floors_data_with_x_cdnt.items())  # 가로 좌표 기준 정렬(층표기 분리 시 세로 좌표가 동일)

    floor_and_floor_height = {}
    for floor_data_with_x_cdnt in floors_data_with_x_cdnt:
        floor_data = floor_data_with_x_cdnt[1]
        for floor_mark, floor_height in floor_data.items():
            floor_and_floor_height.update({floor_mark: floor_height})

    final_floor_data = special_str_split(floor_and_floor_height)

    return final_floor_data, total_floor, stop_floor


def special_str_split(floor_and_floor_height):
    comma_split = {}
    tilde_split = []
    for before_floor, height in floor_and_floor_height.items():
        before_floor=before_floor.replace(" ","")
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
            str_floor = re.findall("(\w+)\W", before_floor)[0]
            str_text = re.findall("(\D+)\d+", str_floor)
            end_floor = re.findall("\W(\w+)", before_floor)[0]
            end_text = re.findall("(\D+)\d+", end_floor)
            if len(str_text) == 0: # start 층표기에 B2~3과 같은 문자가 있는지 확인
                st_no = int(str_floor)
                end_no = int(end_floor) + 1
                for floor in range(st_no, end_no):
                    tilde_split.append([str(floor), height])
            elif len(end_text) == 0: #start 층표기에는 문자가 있고, end 층표기에는 문자가 없을 때
                text = str_text[0]
                st_no = re.findall("\d+", str_floor)[0]
                st_no = int(st_no)
                end_no = int(end_floor) + 1
                for floor in range(st_no, 0, -1):
                    tilde_split.append([text+str(floor), height])
                for floor in range(1, end_no):
                    tilde_split.append([str(floor), height])
            elif len(end_text) > 0: #start, end 모두 층표기에 문자가 있을 때
                text = str_text[0]
                st_no = re.findall("\d+", str_floor)[0]
                st_no = int(st_no)
                end_no = re.findall("\d+", end_floor)[0]
                end_no = int(end_no) - 1
                for floor in range(st_no, end_no, -1):
                    tilde_split.append([text+str(floor), height])

    return tilde_split


floor_and_height, total_floor, stop_floor = get_floor_data(*get_table_cdnt())
print(floor_and_height, total_floor, stop_floor )

df_floor_data = pd.DataFrame(floor_and_height , columns=["층표기", "층고"])
print(df_floor_data)


# doc.close()