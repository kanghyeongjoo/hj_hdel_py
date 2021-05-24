import win32com.client
import re
import glob
import tkinter
import pandas as pd
from tkinter import filedialog

acad = win32com.client.Dispatch("AutoCAD.Application")

def layout_open(prjt_no, layout_kind):

    for filename in glob.glob("D:\DAILY\*.dwg"):
        file_kind = re.findall("(\w)[.]DWG", filename.upper())[0]
        if prjt_no in filename and layout_kind in file_kind:
            layout_path = filename
    try:
        doc = acad.Documents.Open(layout_path)
    except:
        root = tkinter.Tk()
        root.withdraw()
        filename = filedialog.askopenfilename(initialdir=r"C:\Users\Administrator\Downloads", title= prjt_no+"현장 Layout을 선택 바랍니다.",
                                                   filetypes=(("dwg files", "*.dwg"), ("all files", "*.*")))
        filename_split = filename.split("/")
        sel_prjt_no = re.findall("\w?\d+", filename_split[-1].upper())[0]
        sel_kind = re.findall("(\w)[.]DWG", filename.upper())
        if prjt_no in sel_prjt_no   and layout_kind in sel_kind:
            doc = acad.Documents.Open(filename)
        else:
            print("선택한 도면이 올바르지 않습니다. 다시 진행 바랍니다.")
            return

    return doc

def layout_find(prjt_no, layout_kind):
    doc = None
    if acad.Documents.Count == 0:
        doc = layout_open(prjt_no, layout_kind)
    else:
        for document in acad.Documents:
            dwg_f_kind = re.findall("(\w)[.]DWG", document.Name.upper())
            dwg_f_prjt_no = re.findall("\w?\d+", document.Name.upper())[0]
            if dwg_f_prjt_no == prjt_no and layout_kind in dwg_f_kind:
                doc = document

    if doc == None:
        doc = layout_open(prjt_no, layout_kind)

    return doc

def get_floor_data(proj_no):

    global doc
    doc = layout_find(proj_no, "S")
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

    doc = layout_find(proj_no, "E")
    df_floor_data, remote_cp = get_hall_part(df_floor_data)

    return df_floor_data, total_floor, stop_floor, remote_cp


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

def get_hall_part(floor_data):
    jamb_dict={}
    remote_cp = "N"
    floor_data[["JAMB", "JAMB_SPEC","홀버튼", "HPI","위치"]] = ""
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-REMOTE-CP":
            remote_cp = "Y"
        elif entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TITLE":
            dwg_boundary = (238 * 388) * entity.XEffectiveScaleFactor # 도면 경계 SIZE * 축적
            btn_type, lantern = get_hall_item(entity.InsertionPoint, dwg_boundary)
            for att in entity.GetAttributes():
                if att.tagstring == "@TITLE-T":
                    app_jamb, jamb_spec, hpi = get_jamb_type(entity.InsertionPoint, dwg_boundary, att.textstring)
            for att in entity.GetAttributes():
                if att.tagstring == "@TITLE-B":
                    textstring = att.textstring.replace(" ", "")
                    if "기준층" in textstring:
                        app_floor = re.findall("기준층.?(\w+)층", textstring)[0]
                        fl_idx = floor_data.index[floor_data["층표기"] == app_floor]
                        floor_data.loc[fl_idx, "JAMB"] = app_jamb
                        floor_data.loc[fl_idx, "JAMB_SPEC"] = jamb_spec
                        floor_data.loc[fl_idx, "홀버튼"] = btn_type
                        floor_data.loc[fl_idx, "HPI"] = hpi
                        floor_data.loc[fl_idx, "홀랜턴"] = lantern
                        floor_data.loc[fl_idx, "위치"] = "기준층"
                    elif "기타층" in textstring:
                        fl_idx = floor_data.index[(floor_data["위치"] == "")]
                        floor_data.loc[fl_idx, "JAMB"] = app_jamb
                        floor_data.loc[fl_idx, "JAMB_SPEC"] = jamb_spec
                        floor_data.loc[fl_idx, "홀버튼"] = btn_type
                        floor_data.loc[fl_idx, "HPI"] = hpi
                        floor_data.loc[fl_idx, "홀랜턴"] = lantern
                        floor_data.loc[fl_idx, "위치"] = "기타층"
                    elif "최상층" in textstring:
                        floor_data.loc[floor_data.index[-1], "JAMB"] = app_jamb
                        floor_data.loc[floor_data.index[-1], "JAMB_SPEC"] = jamb_spec
                        floor_data.loc[floor_data.index[-1], "홀버튼"] = btn_type
                        floor_data.loc[floor_data.index[-1], "HPI"] = hpi
                        floor_data.loc[floor_data.index[-1], "홀랜턴"] = lantern
                        floor_data.loc[floor_data.index[-1], "위치"] = "최상층"
                    else:
                        textstring = textstring.replace("층", "")
                        floor_list = special_str_split(textstring)
                        for app_floor in floor_list:
                            fl_idx = floor_data.index[floor_data["층표기"] == app_floor]
                            floor_data.loc[fl_idx, "JAMB"] = app_jamb
                            floor_data.loc[fl_idx, "JAMB_SPEC"] = jamb_spec
                            floor_data.loc[fl_idx, "홀버튼"] = btn_type
                            floor_data.loc[fl_idx, "HPI"] = hpi
                            floor_data.loc[fl_idx, "홀랜턴"] = lantern
                            floor_data.loc[fl_idx, "위치"] = "기타층"

    main_fl_idx = floor_data.index[(floor_data["위치"] == "기준층")][0]
    if main_fl_idx > 0:
        floor_data.loc[:main_fl_idx - 1, "위치"] = "지하층"

    if floor_data["홀랜턴"].isnull().sum() > 0:
        del floor_data["홀랜턴"]

    return floor_data, remote_cp

def get_hall_item(titile_pnt, btn_min_gap):
    tit_x, tit_y, tit_z = titile_pnt
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and "LAD-HBTN" in entity.EffectiveName:
            btn_x, btn_y, btn_z = entity.InsertionPoint
            tit_btn_gap = abs(tit_x-btn_x)+abs(tit_y+btn_y)
            if btn_min_gap > tit_btn_gap:
                btn_min_gap = tit_btn_gap
                app_btn_x = btn_x
                if "SMALL" in entity.EffectiveName.upper():
                    btn_type = "HPB"
                elif "LARGE" in entity.EffectiveName.upper():
                    btn_type = "HIP"
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and "LAD-HALL-LANTERN" in entity.EffectiveName:
            lantern_x = entity.InsertionPoint[0]
            if lantern_x == app_btn_x:
                lantern = "Y"
            elif lantern != "Y":
                lantern = "N"
        else:
            lantern = None

    return btn_type, lantern


def get_jamb_type(titile_pnt, jamb_min_gap, textstring):
    tit_x, tit_y, tit_z = titile_pnt
    jamb_ord = 0
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and "LAD-DOOR-JAMB" in entity.EffectiveName:
            jamb_x, jamb_y, jamb_z = entity.InsertionPoint
            tit_jamb_gap = abs(tit_x - jamb_x)+abs(tit_y - jamb_y)
            if jamb_min_gap > tit_jamb_gap:
                jamb_min_gap = tit_jamb_gap
                for att in entity.GetDynamicBlockProperties():  # 동적블럭 속성 가져오기
                    if att.propertyname == "@VISIBLE" and att.value == "Visible":
                        hpi = "Y"
                    elif att.propertyname == "@VISIBLE" and att.value == "Invisible":
                        hpi = "N"
                if "CP" in entity.EffectiveName.upper():
                    jamb_spec = "CP" + re.findall("\d+", textstring)[0]
                    app_jamb = "JAMB(CP);"
                else:
                    jamb_spec = "JP" + re.findall("\d+", textstring)[0]
                    jamb_ord = jamb_ord + 1
                    app_jamb = "JAMB(" + str(jamb_ord) + ");"

    return app_jamb, jamb_spec, hpi

floor_and_height, total_floor, stop_floor, remote_cp = get_floor_data("190580")
print(floor_and_height)
print("층수 : ", total_floor)
print("정지층수 : ", stop_floor )
print("보조제어반 : ", remote_cp)