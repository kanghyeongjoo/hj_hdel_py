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

    if layout_kind == "E":
        ent_blo_name = ["LAD-TITLE", "LAD-DOOR-JAMB", "LAD-HBTN", "LAD-OPEN-AC", "LAD-HALL-LANTERN", "LAD-EMCY-SWITCH",
                        "LAD-REMOTE-CP"]
        ent_group = dict.fromkeys(ent_blo_name)
        for entity in doc.ModelSpace:
            for ent_name in ent_blo_name:
                if entity.EntityName == 'AcDbBlockReference' and ent_name in entity.EffectiveName:
                    if ent_group[ent_name] == None:
                        ent_group[ent_name] = {entity.InsertionPoint[0]: entity}
                    else:
                        ent_group[ent_name].update({entity.InsertionPoint[0]: entity})
                elif entity.EntityName == 'AcDbBlockReference' and "LAD-OPEN-HOLE" in entity.EffectiveName:
                    cdnt = str(entity.InsertionPoint[0]) + "_" + str(entity.InsertionPoint[1])
                    if "LAD-OPEN-HOLE" not in ent_group.keys():
                        ent_group.update({"LAD-OPEN-HOLE": {cdnt: entity}})
                    else:
                        ent_group["LAD-OPEN-HOLE"].update({cdnt: entity})

    elif layout_kind == "S":
        ent_group = []
        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-FORM-A3-SIMPLE":
                palette_area = (238 * 388) * entity.XEffectiveScaleFactor
                ent_group.append({"palette_area" : palette_area})
            elif entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TABLE-FLOOR-HEIGHT":
                ent_group.append({"floor_table_y_cdnt" : entity.InsertionPoint[1]})
                entity.Explode()
                for entity in doc.ModelSpace:
                    if entity.EntityName == 'AcDbText' or entity.EntityName == 'AcDbPolyline':
                        ent_group.append(entity)
            elif entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TABLE-FIRE-DOOR":
                ent_group.append({"fdoor_table_y_cdnt" : entity.InsertionPoint[1]})
                entity.Explode()
                for entity in doc.ModelSpace:
                    if entity.EntityName == 'AcDbText' or entity.EntityName == 'AcDbPolyline':
                        ent_group.append(entity)

    return ent_group

def get_floor_data(proj_no):

    global doc, ent_group
    doc = layout_find(proj_no, "S")
    ent_group = get_entity("S")

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
    ent_group = get_entity("E")
    df_floor_data = get_hall_data(df_floor_data)


    return df_floor_data, total_floor, stop_floor, fireman_sw, remote_dv


def get_table_cdnt():
    floor_table_y_cdnt = None
    fdoor_table_y_cdnt = None
    for ent in ent_group:
        if type(ent) == dict:
            if "palette_area" in ent.keys():
                palette_area = ent_group.pop(ent_group.index(ent)).get("palette_area")
            elif "floor_table_y_cdnt" in ent.keys():
                floor_table_y_cdnt = ent_group.pop(ent_group.index(ent)).get("floor_table_y_cdnt")
            elif "fdoor_table_y_cdnt" in ent.keys():
                fdoor_table_y_cdnt = ent_group.pop(ent_group.index(ent)).get("fdoor_table_y_cdnt")

    if floor_table_y_cdnt != None:
        for ent in ent_group:  # 층 테이블 좌표 구하기
            if ent.EntityName == 'AcDbPolyline' and floor_table_y_cdnt in ent.Coordinates:
                fl_st_x_cdnt = int(ent.Coordinates[0])
                fl_ed_x_cdnt = int(ent.Coordinates[2])
                fl_st_y_cdnt = int(ent.Coordinates[1])
                fl_ed_y_cdnt = int(ent.Coordinates[5])
                floor_cdnt = [fl_st_x_cdnt, fl_ed_x_cdnt, fl_st_y_cdnt, fl_ed_y_cdnt]

    elif floor_table_y_cdnt == None:
        text_insert_cdnt = {}
        for ent in ent_group: # 층 테이블의 데이터 좌표 구하기
            if ent.EntityName == 'AcDbText' and ent.TextString == "FL / ST":
                text_base_y_cdnt = ent.TextAlignmentPoint[1]
            elif ent.EntityName == 'AcDbText' and ent.TextString == "층":
                text_insert_cdnt.update({ent.TextAlignmentPoint[1]:ent.TextAlignmentPoint[0]}) # text Y:X(Y값은 상이함)
        text_base_x_cdnt = text_insert_cdnt.get(text_base_y_cdnt)

        for ent in ent_group: # 데이터와 가까운 Line 좌표 구하기
            if ent.EntityName == 'AcDbPolyline' and ent.Layer == "LAD-OUTLINE":
                x_gap = abs(text_base_x_cdnt - ent.Coordinates[0])
                y_gap = abs(text_base_y_cdnt - ent.Coordinates[1])
                gap  = x_gap + y_gap
                if palette_area > gap: # 방화도어 TABLE과 하고 겹치지 않도록 gap 비교
                    palette_area = gap
                    fl_st_x_cdnt = int(ent.Coordinates[0])
                    fl_ed_x_cdnt = int(ent.Coordinates[2])
                    fl_st_y_cdnt = int(ent.Coordinates[1])
                    fl_ed_y_cdnt = int(ent.Coordinates[5])
                    floor_cdnt = [fl_st_x_cdnt, fl_ed_x_cdnt, fl_st_y_cdnt, fl_ed_y_cdnt]

    if fdoor_table_y_cdnt != None:
        for ent in ent_group:  # 방화도어 TABLE 좌표 구하기
            if ent.EntityName == 'AcDbPolyline' and fdoor_table_y_cdnt in ent.Coordinates:
                fd_st_x_cdnt = int(ent.Coordinates[2])
                fd_ed_x_cdnt = int(ent.Coordinates[0])
                fd_st_y_cdnt = int(ent.Coordinates[5])
                fd_ed_y_cdnt = int(ent.Coordinates[1])
                fire_door_cdnt = [fd_st_x_cdnt, fd_ed_x_cdnt, fd_st_y_cdnt, fd_ed_y_cdnt]


    elif fdoor_table_y_cdnt == None:
        for ent in ent_group: # 방화도어 TABLE의 데이터 좌표 구하기
            if ent.EntityName == 'AcDbText' and ent.TextString == "방화도어 유무":
                text_base_x_cdnt = ent.TextAlignmentPoint[0]
                text_base_y_cdnt = ent.TextAlignmentPoint[1]

        for ent in ent_group:  # 데이터와 가까운 Line 좌표 구하기
            if ent.EntityName == 'AcDbPolyline' and ent.Layer == "LAD-OUTLINE":
                x_gap = abs(text_base_x_cdnt - ent.Coordinates[2])
                y_gap = abs(text_base_y_cdnt - ent.Coordinates[1])
                gap  = x_gap + y_gap
                if palette_area > gap: # 방화도어 table 하고 겹치지 않도록 gap 비교
                    palette_area = gap
                    fd_st_x_cdnt = int(ent.Coordinates[2])
                    fd_ed_x_cdnt = int(ent.Coordinates[0])
                    fd_st_y_cdnt = int(ent.Coordinates[5])
                    fd_ed_y_cdnt = int(ent.Coordinates[1])
                    fire_door_cdnt = [fd_st_x_cdnt, fd_ed_x_cdnt, fd_st_y_cdnt, fd_ed_y_cdnt]

    for ent in ent_group:
        if ent.EntityName == 'AcDbPolyline':
            ent_group.remove(ent)

    return floor_cdnt, fire_door_cdnt


def get_floor_height(s_x_cdnt, e_x_cdnt, s_y_cdnt, e_y_cdnt):
    table_data = {}
    x_cdnt_list = []
    for ent in ent_group:
        if ent.EntityName == 'AcDbText':
            x_cdnt = int(ent.InsertionPoint[0])
            y_cdnt = int(ent.InsertionPoint[1])
            if x_cdnt in range(s_x_cdnt, e_x_cdnt) and y_cdnt in range(e_y_cdnt, s_y_cdnt):
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
    for ent in ent_group:
        if ent.EntityName == 'AcDbText':
            x_cdnt = ent.InsertionPoint[0]
            y_cdnt = ent.InsertionPoint[1]
            if x_cdnt in range(s_x_cdnt, e_x_cdnt) and y_cdnt in range(e_y_cdnt, s_y_cdnt):
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
    floor_mark = floor_mark.replace(" ", "")
    floor_mark = floor_mark.upper()
    floor_mark_list = []
    if "," not in floor_mark and "." not in floor_mark:
        comma_split_floor = [floor_mark]
    elif "," in floor_mark or "." in floor_mark:
        comma_split_floor = re.split("[,.]", floor_mark)

    for del_comma in comma_split_floor:
        if "~" not in del_comma and "-" not in del_comma and "∼" not in del_comma: #물결 문자 상이함
            floor_mark_list.append(del_comma)
        elif "~" in del_comma or "-" in del_comma or "∼" in del_comma:
            del_tilde = re.split("~|-|∼", del_comma)
            convert_floor = []
            prefix_dic = {}
            chk_F = "N"
            for split_ord in range(len(del_tilde)):
                try:
                    split_text = del_tilde[split_ord]
                    chk_floor = int(split_text)
                except:
                    if "F" in split_text:
                        if len(split_text) == 1:
                            chk_floor = 4
                            chk_F = "Y"
                            convert_floor.append(chk_floor)
                        elif len(split_text) > 1 and split_text[-1] =="F":
                            del_F = re.findall("(^\d+)F", split_text)
                            if len(del_F): #del_F에 인덱스로 추출할 때 에러가 나지 않도록 list로 가져와서 값이 있는지 확인
                                chk_floor = int(del_F[0])
                                convert_floor.append(chk_floor)
                        else:
                            print(del_comma, "는 연속된 층표기가 아닙니다. 수정 후 다시 진행해주세요.")
                    elif "F" not in split_text:
                        split_int = re.findall("(^\D+)(\d+$)", split_text)
                        if len(split_int ) > 0:# split_int에 덱스로 추출할 때 에러가 나지 않도록 list로 가져와서 값이 있는지 확인
                            prefix = split_int[0][0]
                            chk_floor = int(split_int[0][1])
                            prefix_dic.update({split_ord:prefix})
                            convert_floor.append(chk_floor)
                        else:
                            print(del_comma, "는 연속된 층표기가 아닙니다. 수정 후 다시 진행해주세요.")
                else:
                    convert_floor.append(chk_floor)

            start_no = convert_floor[0]
            end_no = convert_floor[1]
            if len(prefix_dic):
                if 0 in prefix_dic.keys() and 1 in prefix_dic.keys(): #start, end 모두 접두사가 있을 때
                    if prefix_dic[0] != prefix_dic[1]: #start, end 접두사가 상이할 때
                        print(del_comma, "는 연속된 층표기가 아닙니다. 수정 후 다시 진행해주세요.")
                    else:
                        prefix = prefix_dic[0]
                        if start_no > end_no:
                            for floor in range(start_no, end_no-1, -1):
                                floor_mark_list.append(prefix + str(floor))
                        elif end_no > start_no:
                            for floor in range(start_no, end_no+1):
                                floor_mark_list.append(prefix + str(floor))
                elif 0 in prefix_dic.keys():#start에 접두사가 있고, end에는 없을 때
                    prefix = prefix_dic[0]
                    for floor in range(start_no, 0, -1):
                        floor_mark_list.append(prefix + str(floor))
                    for floor in range(1, end_no+1):
                        floor_mark_list.appen(str(floor))
            else:
                for floor in range(start_no, end_no + 1):
                    floor_mark_list.append(str(floor))

            if chk_F == "Y":
                floor_mark_list[floor_mark_list.index("4")] = "F"


    return floor_mark_list


def get_hall_data(floor_data):
    tit_ents = ent_group["LAD-TITLE"]
    hall_ord = 0
    floor_data["위치"] = "기타층"
    floor_data.loc[floor_data.index[-1], "위치"] = "최상층"
    for tit_ent in tit_ents.values():
        for att in tit_ent.GetAttributes():
            if att.tagstring == "@TITLE-T":
                jamb_spec = re.findall("\d+", att.textstring)[0]
            elif att.tagstring == "@TITLE-B":
                app_floor_info = att.textstring.replace(" ", "")
        hall_items = get_hall_items(hall_ord, jamb_spec)
        if "소방스위치" in hall_items.keys():
            fireman_sw = hall_items["소방스위치"]
        if "분리형 보조제어반" in hall_items.keys():
             remote_dv = hall_items["분리형 보조제어반"]
        hall_ord = hall_ord + 1
        if "기준층" in app_floor_info:
            app_floor = re.findall("기준층.?(\w+)층", app_floor_info)[0]
            idx = floor_data.index[floor_data["층표기"] == app_floor]
            for item, spec in hall_items.items():
                floor_data.loc[idx, item] = spec
                floor_data.loc[idx, "위치"] = "기준층"
        elif "기타층" in app_floor_info:
            idx = floor_data.index[floor_data["위치"] == "기타층"]
            for item, spec in hall_items.items():
                floor_data.loc[idx, item] = spec
        elif "최상층" in app_floor_info:
            idx = floor_data.index[-1]
            for item, spec in hall_items.items():
                floor_data.loc[idx, item] = spec
        else:
            app_floor_info = app_floor_info.replace("층", "")
            floor_list = special_str_split(app_floor_info)
            for app_floor in floor_list:
                idx = floor_data.index[floor_data["층표기"] == app_floor]
                for item, spec in hall_items.items():
                    floor_data.loc[idx, item] = spec

    return floor_data, fireman_sw, remote_dv

def get_hall_items(hall_odr, jamb_spec):
    cable_x_cdnt = {}
    for ent_name, ents in ent_group.items():
        name = ent_name.replace("-","_")
        if ents == None: # 객체가 없는 데이터 제외
            globals()[name] = None
        elif ent_name == "LAD-OPEN-HOLE":
            for cable_hole_cdnt, ent in ents.items():
                cable_x = ent.InsertionPoint[0]
                for att in ent.GetDynamicBlockProperties():
                    if att.propertyname == "@OFFSET-Y" and att.value < 1400:# hole 높이
                        cable_x_cdnt.update({cable_x:"HBTN"})
                    elif att.propertyname == "@OFFSET-Y" and att.value >= 1400:
                        cable_x_cdnt.update({cable_x:"OTHER"})
        elif ent_name == "LAD-EMCY-SWITCH":
            LAD_EMCY_SWITCH = "Y"
            firesw_x_cdnt = list(ents.keys())[0]
            jamb_cdnt_list = list(ent_group["LAD-OPEN-AC"].keys())
            jamb_cdnt_list.sort()
            firesw_app_jamb_cdnt = min(jamb_cdnt_list, key=lambda x: abs(x-firesw_x_cdnt))
            firesw_app_jamb_ord = jamb_cdnt_list.index(firesw_app_jamb_cdnt)
            if firesw_app_jamb_cdnt - firesw_x_cdnt < 0:
                firesw_pst = "RIGHT"
            else:
                firesw_pst = "LEFT"
        else:
            sort_ents=[]
            for sort_cdnt, sort_ent in sorted(ents.items()):
                sort_ents.append(sort_ent)
            globals()[name] = sort_ents

    jamb_hole_ent = LAD_OPEN_AC[hall_odr]
    hole_dic = {"@EMSW-H": "LEFT", "@HBTN-H": "HBTN", "@HPI-H": "HPI", "@LTRN-H": "RIGHT"}
    box_hole = []
    for hole_att in jamb_hole_ent.GetDynamicBlockProperties():
        if hole_att.propertyname in hole_dic.keys() and int(hole_att.value) > 0:
            box_hole.append(hole_dic[hole_att.propertyname])

    jamb_ent = LAD_DOOR_JAMB[hall_odr]
    if "CP" in jamb_ent.EffectiveName.upper():
        jamb_type = "CP" + jamb_spec
        app_jamb = "JAMB(CP);"
    else:
        jamb_type = "JP" + jamb_spec
        app_jamb = "JAMB(" + str(hall_odr+1) + ");"
    for att in jamb_ent.GetDynamicBlockProperties():
        if att.propertyname == "@VISIBLE" and att.value == "Visible":
            hpi = "Y"
            break
        elif att.propertyname == "@VISIBLE" and att.value == "Invisible":
            hpi = "N"
            break

    btn_ent = LAD_HBTN[hall_odr]
    btn_x_cdnt = btn_ent.InsertionPoint[0]
    if "SMALL" in btn_ent.EffectiveName.upper():
        btn_spec = "HPB"
    elif "LARGE" in btn_ent.EffectiveName.upper():
        btn_spec = "HIP"
    if btn_x_cdnt in cable_x_cdnt.keys() and "HBTN" in cable_x_cdnt.values() and "HBTN" not in box_hole:
        btn_type = "BOXLESS"
    elif "HBTN" in box_hole:
        btn_type = "BOX"
    else:
        btn_type = "확인할 수 없습니다."

    floor_items = {"JAMB": jamb_type, "JAMB_ORD": app_jamb, "홀버튼": btn_spec, "홀버튼_취부": btn_type, "HPI": hpi}

    if hpi == "Y":
        if "2" in jamb_spec:
            hpi_type = "JAMB 취부"
        else:
            if "HPI" in box_hole:
                hpi_type = "BOX"
            elif "HPI" not in box_hole:
                hpi_type = "BOXLESS"
        floor_items.update({"HPI_취부": hpi_type})


    if LAD_HALL_LANTERN != None:
        for lant_ent in LAD_HALL_LANTERN:
            lant_cdnt = lant_ent.InsertionPoint[0]
            jamb_cdnt_list = list(ent_group["LAD-OPEN-AC"].keys())
            jamb_cdnt_list.sort()
            lant_app_jamb_cdnt = min(jamb_cdnt_list, key=lambda x: abs(x-lant_cdnt))
            lant_app_jamb_ord = jamb_cdnt_list.index(lant_app_jamb_cdnt)
            if hall_odr == lant_app_jamb_ord:
                if lant_app_jamb_cdnt - lant_cdnt < 0:
                    lant_pst = "LEFT"
                else:
                    lant_pst = "RIGHT"
                if lant_pst in box_hole:
                    lantern = "BOX"
                elif lant_cdnt in cable_x_cdnt.keys() and "OTHER" in cable_x_cdnt.values():
                    lantern = "BOXLESS"
                else:
                    lantern = "홀랜턴 type을 확인할 수 없습니다."

                floor_items.update({"홀랜턴": lantern})


    # firema sw type 찾기 jamb ord 번호가 필요하므로 당 for문에서 구할 것.
    if LAD_EMCY_SWITCH != None:
        if hall_odr == firesw_app_jamb_ord:
            if firesw_pst in box_hole:
                firesw = "BOX"
            elif firesw_x_cdnt in cable_x_cdnt.keys() and "OTHER" in cable_x_cdnt.values():
                firesw = "BOXLESS"
            else:
                firesw = "소방스위치 type을 확인할 수 없습니다."
            floor_items.update({"소방스위치": firesw})
        else:
            pass



    if hall_odr == len(LAD_DOOR_JAMB)-1 and LAD_REMOTE_CP != None:#마지막 jamb(=최상층 jamb)일 떄
        remote_cp = "Y"
        floor_items.update({"분리형 보조제어반": remote_cp})

    return floor_items


floor_and_height, total_floor, stop_floor = get_floor_data("181226")
print(floor_and_height)
print("층수 : ", total_floor)
print("정지층수 : ", stop_floor )

print("걸린 시간 : ", time.time() - start)