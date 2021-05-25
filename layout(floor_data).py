import win32com.client
import re
import glob
import tkinter
import pandas as pd
from tkinter import filedialog
import time

start = time.time()

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

def get_entity(layout_kind):
    ent_list=[]

    if layout_kind == "E":
        ent_blo_name = ["LAD-TITLE", "LAD-HBTN-SMALL", "LAD-HBTN-LONG", "LAD-HALL-LANTERN", "LAD-REMOTE-CP", "LAD-DOOR-JAMB"]
        for entity in doc.ModelSpace:
            for ent_name in ent_blo_name:
                if entity.EntityName == 'AcDbBlockReference' and ent_name in entity.EffectiveName:
                    ent_list.append(entity)

    elif layout_kind == "S":
        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-FORM-A3-SIMPLE":
                palette_area = (238 * 388) * entity.XEffectiveScaleFactor
                ent_list.append({"palette_area" : palette_area})
            elif entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TABLE-FLOOR-HEIGHT":
                ent_list.append({"floor_table_y_cdnt" : entity.InsertionPoint[1]})
                entity.Explode()
                for entity in doc.ModelSpace:
                    if entity.EntityName == 'AcDbText' or entity.EntityName == 'AcDbPolyline':
                        ent_list.append(entity)
            elif entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TABLE-FIRE-DOOR":
                ent_list.append({"fdoor_table_y_cdnt" : entity.InsertionPoint[1]})
                entity.Explode()
                for entity in doc.ModelSpace:
                    if entity.EntityName == 'AcDbText' or entity.EntityName == 'AcDbPolyline':
                        ent_list.append(entity)

    return ent_list

def get_floor_data(proj_no):

    global doc, ent_list
    doc = layout_find(proj_no, "S")
    ent_list = get_entity("S")

    floor_table_cdnt, fire_door_table_cdnt = get_table_cdnt()
    floor_data, total_floor, stop_floor = get_floor_height(*floor_table_cdnt)
    df_floor_data = pd.DataFrame(floor_data, columns=["층표기", "층고"])

    fire_door_data = get_fire_door(*fire_door_table_cdnt)
    df_floor_data["방화도어"] = ""
    if len(set(fire_door_data.values())) == 1:
        df_floor_data["방화도어"] = fire_door_data.values()
    else:
        for app_floor, fire_door in fire_door_data.items():
            fl_idx = df_floor_data.index[df_floor_data["층표기"] == app_floor]
            df_floor_data.loc[fl_idx, "방화도어"] = fire_door

    doc = layout_find(proj_no, "E")
    ent_list = get_entity("E")
    df_floor_data, remote_cp = get_hall_part(df_floor_data)

    return df_floor_data, total_floor, stop_floor, remote_cp


def get_table_cdnt():
    floor_table_y_cdnt = None
    fdoor_table_y_cdnt = None
    for ent in ent_list:
        if type(ent) == dict:
            if "palette_area" in ent.keys():
                palette_area = ent_list.pop(ent_list.index(ent)).get("palette_area")
            elif "floor_table_y_cdnt" in ent.keys():
                floor_table_y_cdnt = ent_list.pop(ent_list.index(ent)).get("floor_table_y_cdnt")
            elif "fdoor_table_y_cdnt" in ent.keys():
                fdoor_table_y_cdnt = ent_list.pop(ent_list.index(ent)).get("fdoor_table_y_cdnt")

    if floor_table_y_cdnt != None:
        for ent in ent_list:  # 층 테이블 좌표 구하기
            if ent.EntityName == 'AcDbPolyline' and floor_table_y_cdnt in ent.Coordinates:
                fl_st_x_cdnt = ent.Coordinates[0]
                fl_ed_x_cdnt = ent.Coordinates[2]
                fl_st_y_cdnt = ent.Coordinates[1]
                fl_ed_y_cdnt = ent.Coordinates[5]
                floor_cdnt = [fl_st_x_cdnt, fl_ed_x_cdnt, fl_st_y_cdnt, fl_ed_y_cdnt]

    elif floor_table_y_cdnt == None:
        text_insert_cdnt = {}
        for ent in ent_list: # 층 테이블의 데이터 좌표 구하기
            if ent.EntityName == 'AcDbText' and ent.TextString == "FL / ST":
                text_base_y_cdnt = ent.TextAlignmentPoint[1]
            elif ent.EntityName == 'AcDbText' and ent.TextString == "층":
                text_insert_cdnt.update({ent.TextAlignmentPoint[1]:ent.TextAlignmentPoint[0]}) # text Y:X(Y값은 상이함)
        text_base_x_cdnt = text_insert_cdnt.get(text_base_y_cdnt)

        for ent in ent_list: # 데이터와 가까운 Line 좌표 구하기
            if ent.EntityName == 'AcDbPolyline' and ent.Layer == "LAD-OUTLINE":
                x_gap = abs(text_base_x_cdnt - ent.Coordinates[0])
                y_gap = abs(text_base_y_cdnt - ent.Coordinates[1])
                gap  = x_gap + y_gap
                if palette_area > gap: # 방화도어 TABLE과 하고 겹치지 않도록 gap 비교
                    palette_area = gap
                    fl_st_x_cdnt = ent.Coordinates[0]
                    fl_ed_x_cdnt = ent.Coordinates[2]
                    fl_st_y_cdnt = ent.Coordinates[1]
                    fl_ed_y_cdnt = ent.Coordinates[5]
                    floor_cdnt = [fl_st_x_cdnt, fl_ed_x_cdnt, fl_st_y_cdnt, fl_ed_y_cdnt]
                    print(floor_cdnt)

    if fdoor_table_y_cdnt != None:
        for ent in ent_list:  # 방화도어 TABLE 좌표 구하기
            if ent.EntityName == 'AcDbPolyline' and fdoor_table_y_cdnt in ent.Coordinates:
                fd_st_x_cdnt = ent.Coordinates[2]
                fd_ed_x_cdnt = ent.Coordinates[0]
                fd_st_y_cdnt = ent.Coordinates[5]
                fd_ed_y_cdnt = ent.Coordinates[1]
                fire_door_cdnt = [fd_st_x_cdnt, fd_ed_x_cdnt, fd_st_y_cdnt, fd_ed_y_cdnt]


    elif fdoor_table_y_cdnt == None:
        for ent in ent_list: # 방화도어 TABLE의 데이터 좌표 구하기
            if ent.EntityName == 'AcDbText' and ent.TextString == "방화도어 유무":
                text_base_x_cdnt = ent.TextAlignmentPoint[0]
                text_base_y_cdnt = ent.TextAlignmentPoint[1]

        for ent in ent_list:  # 데이터와 가까운 Line 좌표 구하기
            if ent.EntityName == 'AcDbPolyline' and ent.Layer == "LAD-OUTLINE":
                x_gap = abs(text_base_x_cdnt - ent.Coordinates[2])
                y_gap = abs(text_base_y_cdnt - ent.Coordinates[1])
                gap  = x_gap + y_gap
                if palette_area > gap: # 방화도어 table 하고 겹치지 않도록 gap 비교
                    palette_area = gap
                    fd_st_x_cdnt = ent.Coordinates[2]
                    fd_ed_x_cdnt = ent.Coordinates[0]
                    fd_st_y_cdnt = ent.Coordinates[5]
                    fd_ed_y_cdnt = ent.Coordinates[1]
                    fire_door_cdnt = [fd_st_x_cdnt, fd_ed_x_cdnt, fd_st_y_cdnt, fd_ed_y_cdnt]

    for ent in ent_list:
        if ent.EntityName == 'AcDbPolyline':
            ent_list.remove(ent)

    return floor_cdnt, fire_door_cdnt


def get_floor_height(s_x_cdnt, e_x_cdnt, s_y_cdnt, e_y_cdnt):
    table_data = {}
    x_cdnt_list = []
    for ent in ent_list:
        if ent.EntityName == 'AcDbText':
            x_cdnt = ent.InsertionPoint[0]
            y_cdnt = ent.InsertionPoint[1]
            if x_cdnt > s_x_cdnt and x_cdnt < e_x_cdnt and y_cdnt < s_y_cdnt and y_cdnt > e_y_cdnt:
                table_data.update({ent.TextAlignmentPoint: ent.TextString}) # 좌표안에 있는 테이블에 있는 모든 TEXT get
                if ent.TextString == "층":  # 윗행(층)과 아래행(층고) 나누는 기준
                    floor_y_cdnt = ent.TextAlignmentPoint[1]
                elif ent.TextString == "층고":  # 윗행(층)과 아래행(층고) 나누는 기준
                    floor_hei_y_cdnt = ent.TextAlignmentPoint[1]
                elif ent.TextString == "FL / ST":  # 층수 구하기
                    flst_x_cdnt = ent.TextAlignmentPoint[0]
                else:
                    x_cdnt_list.append(ent.TextAlignmentPoint[0])

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
    for ent in ent_list:
        if ent.EntityName == 'AcDbText':
            x_cdnt = ent.InsertionPoint[0]
            y_cdnt = ent.InsertionPoint[1]
            if x_cdnt > s_x_cdnt and x_cdnt < e_x_cdnt and y_cdnt < s_y_cdnt and y_cdnt > e_y_cdnt:
                table_data.update({ent.TextAlignmentPoint: ent.TextString}) # 좌표안에 있는 테이블에 있는 모든 TEXT get
                if ent.TextString == "층":  # 윗행(층)과 아래행(층고) 나누는 기준
                    floor_y_cdnt = ent.TextAlignmentPoint[1]
                elif "방화도어" in ent.TextString:  # 윗행(층)과 아래행(층고) 나누는 기준
                    fire_door_y_cdnt = ent.TextAlignmentPoint[1]
                else:
                    x_cdnt_list.append(ent.TextAlignmentPoint[0])

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
    for ent in ent_list:
        if ent.EffectiveName == "LAD-REMOTE-CP":
            remote_cp = "Y"
        elif ent.EffectiveName == "LAD-TITLE":
            dwg_boundary = (238 * 388) * ent.XEffectiveScaleFactor # 도면 경계 SIZE * 축적
            btn_type, lantern = get_hall_item(ent.InsertionPoint, dwg_boundary)
            for att in ent.GetAttributes():
                if att.tagstring == "@TITLE-T":
                    app_jamb, jamb_spec, hpi = get_jamb_type(ent.InsertionPoint, dwg_boundary, att.textstring)
            for att in ent.GetAttributes():
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
    for ent in ent_list:
        if "LAD-HBTN" in ent.EffectiveName:
            btn_x, btn_y, btn_z = ent.InsertionPoint
            tit_btn_gap = abs(tit_x-btn_x)+abs(tit_y+btn_y)
            if btn_min_gap > tit_btn_gap:
                btn_min_gap = tit_btn_gap
                app_btn_x = btn_x
                if "SMALL" in ent.EffectiveName.upper():
                    btn_type = "HPB"
                elif "LARGE" in ent.EffectiveName.upper():
                    btn_type = "HIP"
    for ent in ent_list:
        if "LAD-HALL-LANTERN" in ent.EffectiveName:
            lantern_x = ent.InsertionPoint[0]
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
    for ent in ent_list:
        if "LAD-DOOR-JAMB" in ent.EffectiveName:
            jamb_x, jamb_y, jamb_z = ent.InsertionPoint
            tit_jamb_gap = abs(tit_x - jamb_x)+abs(tit_y - jamb_y)
            if jamb_min_gap > tit_jamb_gap:
                jamb_min_gap = tit_jamb_gap
                for att in ent.GetDynamicBlockProperties():  # 동적블럭 속성 가져오기
                    if att.propertyname == "@VISIBLE" and att.value == "Visible":
                        hpi = "Y"
                    elif att.propertyname == "@VISIBLE" and att.value == "Invisible":
                        hpi = "N"
                if "CP" in ent.EffectiveName.upper():
                    jamb_spec = "CP" + re.findall("\d+", textstring)[0]
                    app_jamb = "JAMB(CP);"
                else:
                    jamb_spec = "JP" + re.findall("\d+", textstring)[0]
                    jamb_ord = jamb_ord + 1
                    app_jamb = "JAMB(" + str(jamb_ord) + ");"

    return app_jamb, jamb_spec, hpi

floor_and_height, total_floor, stop_floor, remote_cp = get_floor_data("190390")
print(floor_and_height)
print("층수 : ", total_floor)
print("정지층수 : ", stop_floor )
print("보조제어반 : ", remote_cp)

print("걸린 시간 : ", time.time() - start)