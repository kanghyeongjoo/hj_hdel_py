import win32com.client
import math
import string
import fnmatch

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

def get_floor_datas():

    for entity in doc.ModelSpace: # 층고 테이블 분해
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TABLE-FLOOR-HEIGHT":
            table_blo_y_cdnt = entity.InsertionPoint[1]
            entity.Explode()

    for entity in doc.ModelSpace: # 테이블 좌표 구하기
        if entity.EntityName == 'AcDbPolyline' and entity.Coordinates[1] == table_blo_y_cdnt:
            start_x_cdnt = entity.Coordinates[0]
            end_x_cdnt = entity.Coordinates[2]
            start_y_cdnt = entity.Coordinates[1]
            end_y_cdnt = entity.Coordinates[5]

    table_datas_list = {}
    for entity in doc.ModelSpace:  # 테이블에 있는 모든 TEXT는 정렬좌표와 Dictionary로 get
        if entity.EntityName == 'AcDbText':
            x_cdnt = entity.InsertionPoint[0]
            y_cdnt = entity.InsertionPoint[1]
            if x_cdnt > start_x_cdnt and x_cdnt < end_x_cdnt and y_cdnt < start_y_cdnt and y_cdnt > end_y_cdnt:
                table_datas_list.update({entity.TextAlignmentPoint: entity.TextString})
                str_text = entity.TextString
                if str_text == "층": # 윗행(층)과 아래행(층고) 나누는 기준
                    floor_row_y_cdnt = entity.TextAlignmentPoint[1]

    floor_datas = {}
    floor_hight_datas = {}
    for data_cdnt, table_data in table_datas_list.items():
        if data_cdnt[1] == floor_row_y_cdnt: # 층행
            floor_datas.update({table_data:data_cdnt}) # 층표기 : 좌표(층표기는 고유하지만 층표기 분리 시 세로 좌표와 층고는 중복이 가능)
        elif data_cdnt[1] < floor_row_y_cdnt: # 층고행
            floor_hight_datas.update({data_cdnt:table_data}) # 좌표 : 층고

    floors_with_cdnt = {}
    for floor_name, floor_name_cdnt in floor_datas.items(): # 층행 Dictionary
        if floor_name != "층" and floor_name != "FL / ST":
            for floor_hight_cdnt, floor_hight in floor_hight_datas.items(): # 층고행 Dictionary
                if floor_name_cdnt[0] == floor_hight_cdnt[0]: # 세로좌표 비교하여 층표시와 층고 Matching
                    floor_dic = {floor_name:floor_hight}
                    floors_with_cdnt.update({floor_name_cdnt[0]:floor_dic})

    floors_with_cdnt = sorted(floors_with_cdnt.items()) # 세로 좌표 기준 정렬(단, 층표기 분리시 세로 좌표가 동일하므로 세로 층표기를 정렬이후에 나누는 것도 고려해볼 것)

    floors=[]
    for floor in floors_with_cdnt:
        flo = floor[1]
        floors.append(flo)

    return floors

#fl / st도 구하기(함수 나누기)
#층표기 분리하기(함수 나누기)

floor_data = get_floor_datas()
print(floor_data)