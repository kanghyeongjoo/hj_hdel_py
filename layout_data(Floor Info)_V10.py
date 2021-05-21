import win32com.client
import re
import pandas as pd

acad = win32com.client.Dispatch("AutoCAD.Application")

for doc in acad.Documents:
    layout_kind = re.findall("(\w)[.]DWG", doc.Name.upper())
    if "S" in layout_kind:
        doc.Activate()
doc = acad.ActiveDocument


def get_floor_data():
    floor_table_y_cdnt = None
    fdoor_table_y_cdnt = None
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-FORM-A3-SIMPLE":
            palette_area = (238 * 388) * entity.XEffectiveScaleFactor # palett 면적
        elif entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TABLE-FLOOR-HEIGHT":
            floor_table_y_cdnt = entity.InsertionPoint[1]
            entity.Explode() # 층고 테이블 분해
        elif entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TABLE-FIRE-DOOR":
            fdoor_table_y_cdnt = entity.InsertionPoint[1]
            entity.Explode() # 방화도어 테이블 분해

    floor_table_cdnt = get_table_cdnt("floor", floor_table_y_cdnt, palette_area)
    floor_data, total_floor, stop_floor = get_floor_height(*floor_table_cdnt)
    df_floor_data = pd.DataFrame(floor_data, columns=["층표기", "층고"])

    fire_door_table_cdnt = get_table_cdnt("door", fdoor_table_y_cdnt, palette_area)
    fire_door_data = get_fire_door(*fire_door_table_cdnt)
    df_floor_data["방화도어"] = ""
    if len(set(fire_door_data.values())) == 1:
        df_floor_data["방화도어"] = fire_door_data.values()
    else:
        for app_floor, fire_door in fire_door_data.items():
            fl_idx = df_floor_data.index[df_floor_data["층표기"] == app_floor]
            df_floor_data.loc[fl_idx, "방화도어"] = fire_door

    return df_floor_data, total_floor, stop_floor


def get_table_cdnt(table_type, table_y_cdnt, palette_area):

    if table_type == "floor" and table_y_cdnt != None:
        for entity in doc.ModelSpace:  # 테이블 좌표 구하기
            if entity.EntityName == 'AcDbPolyline' and table_y_cdnt in entity.Coordinates:
                start_x_cdnt = entity.Coordinates[0]
                end_x_cdnt = entity.Coordinates[2]
                start_y_cdnt = entity.Coordinates[1]
                end_y_cdnt = entity.Coordinates[5]

    elif table_type == "floor" and table_y_cdnt == None:
        text_insert_cdnt = {}
        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbText' and entity.TextString == "FL / ST":
                text_base_y_cdnt = entity.TextAlignmentPoint[1]
            elif entity.EntityName == 'AcDbText' and entity.TextString == "층":
                text_insert_cdnt.update({entity.TextAlignmentPoint[1]:entity.TextAlignmentPoint[0]}) # text Y:X(Y값은 상이함)
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

    if table_type == "door" and table_y_cdnt != None:
        for entity in doc.ModelSpace:  # 테이블 좌표 구하기
            if entity.EntityName == 'AcDbPolyline' and table_y_cdnt in entity.Coordinates:
                start_x_cdnt = entity.Coordinates[2]
                end_x_cdnt = entity.Coordinates[0]
                start_y_cdnt = entity.Coordinates[5]
                end_y_cdnt = entity.Coordinates[1]

    elif table_type == "door" and table_y_cdnt == None:
        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbText' and entity.TextString == "방화도어 유무":
                text_base_x_cdnt = entity.TextAlignmentPoint[0]
                text_base_y_cdnt = entity.TextAlignmentPoint[1]

        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbPolyline' and entity.Layer == "LAD-OUTLINE":
                x_gap = abs(text_base_x_cdnt - entity.Coordinates[2])
                y_gap = abs(text_base_y_cdnt - entity.Coordinates[1])
                gap  = x_gap + y_gap
                if palette_area > gap: # 방화도어 table 하고 겹치지 않도록 gap 비교
                    palette_area = gap
                    start_x_cdnt = entity.Coordinates[2]
                    end_x_cdnt = entity.Coordinates[0]
                    start_y_cdnt = entity.Coordinates[5]
                    end_y_cdnt = entity.Coordinates[1]

    return start_x_cdnt, end_x_cdnt, start_y_cdnt, end_y_cdnt


def get_floor_height(s_x_cdnt, e_x_cdnt, s_y_cdnt, e_y_cdnt):
    table_data = {}
    x_cdnt_list = []
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbText':
            x_cdnt = entity.InsertionPoint[0]
            y_cdnt = entity.InsertionPoint[1]
            if x_cdnt > s_x_cdnt and x_cdnt < e_x_cdnt and y_cdnt < s_y_cdnt and y_cdnt > e_y_cdnt:
                table_data.update({entity.TextAlignmentPoint: entity.TextString}) # 좌표안에 있는 테이블에 있는 모든 TEXT get
                if entity.TextString == "층":  # 윗행(층)과 아래행(층고) 나누는 기준
                    floor_y_cdnt = entity.TextAlignmentPoint[1]
                elif entity.TextString == "층고":  # 윗행(층)과 아래행(층고) 나누는 기준
                    floor_hei_y_cdnt = entity.TextAlignmentPoint[1]
                elif entity.TextString == "FL / ST":  # 층수 구하기
                    flst_x_cdnt = entity.TextAlignmentPoint[0]
                else:
                    x_cdnt_list.append(entity.TextAlignmentPoint[0])

    fl_st_data = table_data.get((flst_x_cdnt, floor_hei_y_cdnt, 0.0))
    total_floor = re.findall("(\d+)/", fl_st_data)[0]
    stop_floor = re.findall("/(\d+)", fl_st_data)[0]

    x_cdnt_list = list(set(x_cdnt_list)) # 중복 좌표 삭제
    x_cdnt_list.remove(flst_x_cdnt)
    x_cdnt_list.sort() # x좌표 순서대로 정리

    floor_and_height = []
    for x in x_cdnt_list:
        floor_text = table_data.get((x, floor_y_cdnt, 0.0))
        floor_mark_list = special_str_split(floor_text)
        floor_height = table_data.get((x, floor_hei_y_cdnt, 0.0))
        for floor_mark in floor_mark_list:
            floor_and_height.append([floor_mark, floor_height])

    return floor_and_height, total_floor, stop_floor

def get_fire_door(s_x_cdnt, e_x_cdnt, s_y_cdnt, e_y_cdnt):
    table_data = {}
    x_cdnt_list = []
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbText':
            x_cdnt = entity.InsertionPoint[0]
            y_cdnt = entity.InsertionPoint[1]
            if x_cdnt > s_x_cdnt and x_cdnt < e_x_cdnt and y_cdnt < s_y_cdnt and y_cdnt > e_y_cdnt:
                table_data.update({entity.TextAlignmentPoint: entity.TextString}) # 좌표안에 있는 테이블에 있는 모든 TEXT get
                if entity.TextString == "층":  # 윗행(층)과 아래행(층고) 나누는 기준
                    floor_y_cdnt = entity.TextAlignmentPoint[1]
                elif "방화도어" in entity.TextString:  # 윗행(층)과 아래행(층고) 나누는 기준
                    fire_door_y_cdnt = entity.TextAlignmentPoint[1]
                else:
                    x_cdnt_list.append(entity.TextAlignmentPoint[0])

    x_cdnt_list = list(set(x_cdnt_list)) # 중복 좌표 삭제
    x_cdnt_list.sort() # x좌표 순서대로 정리

    floor_and_fire_door = {}
    for x in x_cdnt_list:
        floor_text = table_data.get((x, floor_y_cdnt, 0.0))
        floor_mark_list = special_str_split(floor_text)

        fire_door = table_data.get((x, fire_door_y_cdnt, 0.0))
        fire_door = fire_door.upper()
        if fire_door == "O" or fire_door == "YES":
            fire_door = re.sub("O|YES", "Y", fire_door)[0]
        elif fire_door == "X" or fire_door == "NO":
            fire_door = re.sub("X|NO", "N", fire_door)[0]

        for floor_mark in floor_mark_list:
            floor_and_fire_door.update({floor_mark: fire_door})

    return floor_and_fire_door

def special_str_split(floor_mark):
    floor_mark_list = []
    if "," not in floor_mark and "." not in floor_mark:
        comma_split_floor = [floor_mark]
    elif "," in floor_mark or "." in floor_mark:
        comma_split_floor = re.split("[,.]", floor_mark)

    for split_floor in comma_split_floor:
        if "~" not in split_floor and "-" not in split_floor:
            floor_mark_list.append(split_floor)
        elif "~" in split_floor or "-" in split_floor:
            str_floor = re.findall("(\w+)\W", split_floor)[0]
            str_text = re.findall("(\D+)\d+", str_floor)
            end_floor = re.findall("\W(\w+)", split_floor)[0]
            end_text = re.findall("(\D+)\d+", end_floor)
            if not len(str_text): # start 층표기에 B2~3과 같은 문자가 있는지 확인
                st_no = int(str_floor)
                end_no = int(end_floor) + 1
                for floor in range(st_no, end_no):
                    floor_mark_list.append(str(floor))
            elif not len(end_text): #start 층표기에는 문자가 있고, end 층표기에는 문자가 없을 때
                text = str_text[0]
                st_no = re.findall("\d+", str_floor)[0]
                st_no = int(st_no)
                end_no = int(end_floor) + 1
                for floor in range(st_no, 0, -1):
                    floor_mark_list.append(text+str(floor))
                for floor in range(1, end_no):
                    floor_mark_list.append(str(floor))
            elif len(end_text) > 0: #start, end 모두 층표기에 문자가 있을 때
                text = str_text[0]
                st_no = re.findall("\d+", str_floor)[0]
                st_no = int(st_no)
                end_no = re.findall("\d+", end_floor)[0]
                end_no = int(end_no) - 1
                for floor in range(st_no, end_no, -1):
                    floor_mark_list.append(text+str(floor))

    return floor_mark_list


floor_and_height, total_floor, stop_floor = get_floor_data()
print(floor_and_height)
print("층수 : ", total_floor)
print("정지층수 : ", stop_floor )


# doc.close()