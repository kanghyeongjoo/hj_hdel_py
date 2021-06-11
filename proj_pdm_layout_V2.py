import win32com.client
import tkinter
from tkinter import filedialog
import re
import glob
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

start = time.time()
acad = win32com.client.Dispatch("AutoCAD.Application")


# get ouid
def get_ouid(proj_no):
    ouid_url = "http://plm.hdel.co.kr/jsp/help/ouidList.jsp?md%24number={}".format(proj_no)
    ouid_data = requests.get(ouid_url)
    ouid_address = BeautifulSoup(ouid_data.content, "html.parser")
    ouid_get = ouid_address.find("form").next_sibling.strip()
    return ouid_get  # get spec을 진행할 때 sepc url을 여기서 return하는 것도 고려해볼 것.


# 코드별 값 확인 URL에서 현장 정보 가져오기
def get_spec(ouid):
    spec_url = "http://plm.hdel.co.kr/jsp/plmetc/elvInfo/elvinfomation.jsp?cOuid=860c9bb8&iOuid={}".format(ouid)
    spec_data = requests.get(spec_url)
    spec_address = BeautifulSoup(spec_data.content, "html.parser")
    spec_list = spec_address.find_all("tr", "01-cell")
    spec_get = {}
    el_code_list = ["EL_A", "EL_B", "EL_C", "EL_D", "EL_E"]
    for spec in spec_list:
        code = spec.find_all("td")[3].get_text()
        if "TEXT" not in code and any(el_code in code for el_code in el_code_list):
            val = spec.find_all("td")[4].get_text()
            spec_get.update({code: val})
    return spec_get


def get_pdm_spec(proj_no):
    ouid = get_ouid(proj_no)
    pdm_spec = get_spec(ouid)
    df_pdm_spec = pd.DataFrame([pdm_spec])
    df_pdm_spec.index = ["PDM_DATA"]
    return df_pdm_spec


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
        filename = filedialog.askopenfilename(initialdir=r"C:\Users\Administrator\Downloads",
                                              title=prjt_no + "현장 Layout을 선택 바랍니다.",
                                              filetypes=(("dwg files", "*.dwg"), ("all files", "*.*")))
        filename_split = filename.split("/")
        sel_prjt_no = re.findall("\w?\d+", filename_split[-1].upper())[0]
        sel_kind = re.findall("(\w)[.]DWG", filename.upper())
        if prjt_no in sel_prjt_no and layout_kind in sel_kind:
            doc = acad.Documents.Open(filename)
        else:
            print("선택한 도면이 올바르지 않습니다. 다시 진행 바랍니다.")
            return

    return doc


def layout_find(prjt_no, layout_kind):
    global doc
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
    global ent_group
    if layout_kind == "H":
        ent_blo_name = ["LAD-RAIL", "LAD-OPB", "LAD-CWT", "LAD-GOV", "DIM_ENT", "LAD-CP"]
        ent_group = dict.fromkeys(ent_blo_name)
        for entity in doc.ModelSpace:
            if entity.EntityName == "AcDbBlockReference" and entity.EffectiveName == "LAD-HOISTWAY-HP-SC":
                ent_group.update({"hoistway_m": "CEMEN"})
                entity.explode()
            elif entity.EntityName == "AcDbBlockReference" and entity.EffectiveName == "LAD-HOISTWAY-HP-SS":
                ent_group.update({"hoistway_m": "ST"})
                entity.explode()
            elif entity.EntityName == "AcDbBlockReference" and entity.EffectiveName == "LAD-CAR-1SCO":
                ent_group.update({"car_center": entity.InsertionPoint})
                entity.explode()
            elif entity.EntityName == "AcDbBlockReference" and entity.EffectiveName == "LAD-CAR-1SCO-CP":
                ent_group.update({"platform_cp": entity.InsertionPoint})
            elif entity.EntityName == 'AcDbBlockReference' and entity.Name == "LAD-FORM-A3-DETAIL":
                ent_group.update({"spec_data": entity})

        for entity in doc.ModelSpace:
            if entity.EntityName == "AcDbBlockReference":
                for ent_name in ent_blo_name:
                    if ent_name in entity.EffectiveName:
                        if ent_group[ent_name] == None:
                            ent_group[ent_name] = []
                            ent_group[ent_name].append(entity)
                        else:
                            ent_group[ent_name].append(entity)
            elif entity.EntityName == "AcDbRotatedDimension" and entity.TextOverride != "":
                if ent_group["DIM_ENT"] == None:
                    ent_group["DIM_ENT"] = []
                    ent_group["DIM_ENT"].append(entity)
                else:
                    ent_group["DIM_ENT"].append(entity)

    elif layout_kind == "M":
        ent_blo_name = ["LAD-HOLE", "U107", "LAD-HBTN-HOLE", "LAD-CP", "LAD-GOV", "LAD-TM", "LAD-HATCH", "DIM_ENT", "hst_cent",
                        "hst_hole_cent", "door_cdnt"]
        ent_group = dict.fromkeys(ent_blo_name)
        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbBlockReference':
                if "LAD-HOISTWAY" in entity.EffectiveName and entity.EffectiveName[-1] == "H":
                    ent_group.update({"hst_hole_cent": entity.InsertionPoint})
                    entity.explode()
                elif "LAD-HOISTWAY" in entity.EffectiveName:
                    ent_group.update({"hst_cent": entity.InsertionPoint})
                    entity.explode()
                elif entity.EffectiveName == "A$C4bb02043":
                    entity.explode()
                else:
                    for ent_name in ent_blo_name:
                        if ent_name in entity.EffectiveName:
                            if ent_group[ent_name] == None:
                                ent_group[ent_name] = []
                                ent_group[ent_name].append(entity)
                            else:
                                ent_group[ent_name].append(entity)

        for entity in doc.ModelSpace:
            if entity.EntityName == "AcDbRotatedDimension" and entity.TextOverride != "":
                if ent_group["DIM_ENT"] == None:
                    ent_group["DIM_ENT"] = []
                    ent_group["DIM_ENT"].append(entity)
                else:
                    ent_group["DIM_ENT"].append(entity)
            elif entity.EntityName == "AcDbText" and "기계실 출입문" in entity.TextString:
                ent_group.update({"door_cdnt": entity.InsertionPoint})
            elif ent_group["LAD-CP"] == None and entity.EntityName == 'AcDbBlockReference' and "LAD-CP" in entity.EffectiveName:
                ent_group.update({"LAD-CP": [entity]})

        if ent_group["hst_cent"] == None:
            if ent_group["LAD-TM"] != None:
                hst_cent = ent_group["LAD-TM"][0].InsertionPoint
                ent_group.update({"hst_cent": hst_cent})
            else:
                print("hoistway 중심점을 확인할 수 없습니다.")  # 이 부분은 나중에 직접 입력 또는 CAD에서 직접 선택할 수 있도록 한다.

        if ent_group["hst_hole_cent"] == None:
            if ent_group["LAD-HOLE"] != None:
                hst_hole_cent = ent_group["LAD-HOLE"][0].InsertionPoint
                ent_group.update({"hst_hole_cent": hst_hole_cent})
            elif ent_group["U107"] != None:
                hst_hole_cent = ent_group["U107"][0].InsertionPoint
                ent_group.update({"hst_hole_cent": hst_hole_cent})
            else:
                print("hoistway hole 중심점을 확인할 수 없습니다.")  # 이 부분은 나중에 직접 입력 또는 CAD에서 직접 선택할 수 있도록 한다.

    elif layout_kind == "S":
        ent_group = {"Text":[], "Polyline":[]}
        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbBlockReference':
                if entity.EffectiveName == "LAD-FORM-A3-SIMPLE":
                    palette_area = (238 * 388) * entity.XEffectiveScaleFactor
                    ent_group.update({"palette_area" : palette_area})
                elif entity.EffectiveName == "LAD-TABLE-FLOOR-HEIGHT":
                    ent_group.update({"floor_table_y_cdnt" : entity.InsertionPoint[1]})
                    entity.Explode()
                elif entity.EffectiveName == "LAD-TABLE-FIRE-DOOR":
                    ent_group.update({"fdoor_table_y_cdnt" : entity.InsertionPoint[1]})
                    entity.Explode()
                elif "LAD-HOISTWAY" in entity.EffectiveName:
                    ent_group.update({"hoistway_info":entity})

        for entity in doc.ModelSpace:
            if entity.EntityName == 'AcDbText':
                ent_group["Text"].append(entity)
            elif entity.EntityName == 'AcDbPolyline':
                ent_group["Polyline"].append(entity)

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

    try:
        if layout_kind == "H" or layout_kind == "M":
            doc.Activate
            doc.SendCommand('setxdata ')
            time.sleep(1)
    except:
        time.sleep(2)

    return ent_group


def get_floor_plan_data():
    spec_ent = ent_group["spec_data"]
    spec_data = {}
    tag_name = {"@GOVERNOR": "EL_ECGV", "@CAR_SAFETY": "EL_ECSF", "@TM_TYPE": "EL_ETM"}  # 특성코드와 dic형태로 매칭해주는 것도 생각해볼
    trs_tag_name = {"@BALANCE": "EL_ECBA", "@NO": "EL_ECN", "@V_SPEC": ["EL_AVOLT", "EL_ALI", "EL_AHZ"],
                    "@DRIVE_TYPE": "EL_ADRV", "@DRIVE": "EL_ATYP", "@SPEED": "EL_ASPD",
                    "@CAPA": ["EL_AMAN", "EL_ACAPA"], "@USE": "EL_AUSE", "@DOOR_DRIVE": "EL_AOPEN",
                    "@MOTOR_CAPA": "EL_ETMM",
                    "@ROPE_SPEC": ["EL_ERPD", "EL_ERPW", "EL_ERPR"], "@DOOR_SIZE": ["EL_ECJJ", "EL_ECHH"],
                    "@CAR_SIZE": ["EL_ECCA", "EL_ECCB", "EL_ECCH"], "@CB_TYPE": "EL_DURTB"}  # 변환이 필요한 코드
    for spec_att in spec_ent.GetAttributes():
        if spec_att.TagString in tag_name.keys():
            spec_data.update({tag_name[spec_att.TagString]: spec_att.TextString})
        elif spec_att.TagString in trs_tag_name.keys():
            if spec_att.TagString == "@BALANCE":
                att_value = re.findall("\d+", spec_att.TextString)[0]
                spec_data.update({trs_tag_name[spec_att.TagString]: att_value})
            elif spec_att.TagString == "@NO":
                att_value = re.findall("\d+", spec_att.TextString)
                if len(att_value) == 1:
                    spec_data.update({trs_tag_name[spec_att.TagString]: att_value[0]})
                else:
                    spec_data.update({trs_tag_name[spec_att.TagString]: att_value})
            elif spec_att.TagString == "@V_SPEC":
                att_value = spec_att.TextString.lower().replace(" ", "")
                att_value_list = re.findall("\d+(?=v)|\d+(?=hz)", att_value)
                for idx in range(len(att_value_list)):
                    spec_data.update({trs_tag_name[spec_att.TagString][idx]: att_value_list[idx]})
            elif spec_att.TagString == "@DRIVE_TYPE":
                car_oper_type = re.findall("\d+", spec_att.TextString)
                att_value = car_oper_type[0] + "C" + car_oper_type[1] + "BC"
                spec_data.update({trs_tag_name[spec_att.TagString]: att_value})
            elif spec_att.TagString == "@DRIVE":
                if "WBSS" in spec_att.TextString:
                    att_value = "WBSS2_(SSVF)"
                elif "LXVF" in spec_att.TextString or "WBLX" in spec_att.TextString:
                    att_value = "WBLX1_(LXVF)"
                else:
                    att_value = spec_att.TextString + "은 사양 추가 요청바랍니다."
                spec_data.update({trs_tag_name[spec_att.TagString]: att_value})
            elif spec_att.TagString == "@SPEED":
                att_value = re.search("^\d+", spec_att.TextString).group()
                spec_data.update({trs_tag_name[spec_att.TagString]: att_value})
            elif spec_att.TagString == "@CAPA":
                att_value_list = re.findall("\d+", spec_att.TextString)
                for idx in range(len(att_value_list)):
                    spec_data.update({trs_tag_name[spec_att.TagString][idx]: att_value_list[idx]})
            elif spec_att.TagString == "@USE":
                if "(" in spec_att.TextString or "[" in spec_att.TextString:
                    att_value_idx = re.search("(.+)\(|(.+)\[", spec_att.TextString).end()
                else:
                    att_value_idx = len(spec_att.TagString)
                att_value_list = re.findall("\w+", spec_att.TextString[:att_value_idx])
                if len(att_value_list) == 1:
                    use_cvt = {"인승": "PS", "장애": "HC", "비상": "EP", "병원": "BD", "전망": "OB", "누드": "ND", "인화": "PF",
                               "화물": "FT", "자동차": "AM"}
                    cvt_value = att_value_list[0][:2]
                    use_value = use_cvt[cvt_value]
                else:
                    use_cvt = {"비상": "E", "병원": "B", "전망": "O", "누드": "N", "인화": "F", "장애": "H"}
                    for be_data, af_data in use_cvt.items():
                        for cvt_value in att_value_list:
                            if be_data == cvt_value[:2] and "use_value" not in locals():
                                use_value = af_data
                            elif be_data == cvt_value and "use_value" in locals():
                                use_value = att_value + af_data
                spec_data.update({trs_tag_name[spec_att.TagString]: use_value})
            elif spec_att.TagString == "@DOOR_DRIVE":
                pdm_drive = ["1SCO", "2SSO", "2SL", "2SR", "2SLR", "3SSO", "3SL", "3SR", "3SLR", "2SCO", "2UP", "2UL",
                             "2UR",
                             "2ULR", "3UP", "3UL", "3UR", "3ULR", "1SSO", "1SL", "1SR", "1SLR"]
                for drive in pdm_drive:
                    if drive in spec_att.TextString:
                        drive_value = drive
                if "drive_value" not in locals():
                    if "CENTER" in spec_att.TextString:
                        if re.search('\d', spec_att.TextString).group() == "2":
                            drive_value = "1SCO"
                        else:
                            drive_value = "Door open" + spec_att.TextString + "에 대한 정의가 핑요합니다."
                    elif "SIDE" in spec_att.TextString:
                        if re.search('\d', spec_att.TextString) == "2":
                            drive_value = "2SSO"
                        else:
                            drive_value = "Door open" + spec_att.TextString + "에 대한 정의가 핑요합니다."
                spec_data.update({trs_tag_name[spec_att.TagString]: drive_value})
            elif spec_att.TagString == "@MOTOR_CAPA":
                att_value = re.findall('(\d+\.?\d?)', spec_att.TextString)[0]
                spec_data.update({trs_tag_name[spec_att.TagString]: att_value})
            elif spec_att.TagString == "@CB_TYPE":
                if spec_att.TextString == "URETHAN":
                    u_bfr = "Y"
                else:
                    u_bfr = "N"
                spec_data.update({trs_tag_name[spec_att.TagString]: u_bfr})
            elif spec_att.TagString == "@ROPE_SPEC":
                cvt_value = spec_att.TextString.replace(" ", "")
                for under_name in trs_tag_name["@ROPE_SPEC"]:
                    if under_name == "EL_ERPD":
                        att_value = re.findall("ø(\d+)", cvt_value)[0]
                    elif under_name == "EL_ERPW":
                        att_value = re.findall("X(\d+)", cvt_value)[0]
                    elif under_name == "EL_ERPR":
                        att_value = re.findall("\((\d+:\d+)\)", cvt_value)[0]
                    spec_data.update({under_name: att_value})
            elif spec_att.TagString == "@DOOR_SIZE":
                for under_name in trs_tag_name["@DOOR_SIZE"]:
                    if under_name == "EL_ECJJ":
                        att_value = re.findall("JJ\D+(\d+)", spec_att.TextString)[0]
                    elif under_name == "EL_ECHH":
                        att_value = re.findall("HH\D+(\d+)", spec_att.TextString)[0]
                    spec_data.update({under_name: int(att_value)})
            elif spec_att.TagString == "@CAR_SIZE":
                for under_name in trs_tag_name["@CAR_SIZE"]:
                    if under_name == "EL_ECCA":
                        att_value = re.findall("CA\D+(\d+)", spec_att.TextString)[0]
                    elif under_name == "EL_ECCB":
                        att_value = re.findall("CB\D+(\d+)", spec_att.TextString)[0]
                    elif under_name == "EL_ECCH":
                        att_value = re.findall("CH\D+(\d+)", spec_att.TextString)[0]
                    spec_data.update({under_name: int(att_value)})

    floor_plan_data = {}
    if ent_group["hoistway_m"] != None:
        floor_plan_data.update({"EL_EHM": ent_group["hoistway_m"]})  # 승강로 재질
    dim_name = {"균형추레일간의거리(세로)": "EL_ECWBG", "승강로내부(세로)": "EL_EHV", "카바닥(세로)": "EL_ECBB", "카내부(세로)": "EL_ECCB",
                "카바닥(가로)": "EL_ECAA", "출입구유효폭(가로)": "EL_ECJJ", "카레일간의거리(가로)": "EL_ECBG", "승강로내부(가로)": "EL_EHH",
                "카내부(가로)": "EL_ECCA"}
    for dim_ent in ent_group["DIM_ENT"]:
        del_s = dim_ent.TextOverride.replace(" ", "")
        size_name = re.findall("[가-힣]+", del_s)[0]
        Xdata = dim_ent.GetXData("", "Type", "Data")
        pt1 = Xdata[1][-2]
        pt2 = Xdata[1][-1]
        if int(pt1[0]) == int(pt2[0]):
            size_name = size_name + "(세로)"
        elif int(pt1[1]) == int(pt2[1]):
            size_name = size_name + "(가로)"
        else:
            gaps = {}
            gaps.update({abs(int(pt1[0]) - int(pt2[0])): "(가로)"})
            gaps.update({abs(int(pt1[1]) - int(pt2[1])): "(세로)"})
            size_name = size_name + gaps[int(dim_ent.Measurement)]

        size = round(dim_ent.Measurement)

        if size_name == "승강로내부(가로)":
            hoist_lft_x = min(int(pt1[0]), int(pt2[0]))
            hoist_rgt_x = max(int(pt1[0]), int(pt2[0]))
            car_cen_h = abs(hoist_lft_x - int(ent_group["car_center"][0]))
            floor_plan_data.update({"EL_ECHOR": car_cen_h})  # 카중심:가로
        elif size_name == "승강로내부(세로)":
            hoist_fro_y = min(int(pt1[1]), int(pt2[1]))
            hoist_rear_y = max(int(pt1[1]), int(pt2[1]))
            car_cen_v = abs(hoist_fro_y - int(ent_group["car_center"][1]))
            floor_plan_data.update({"EL_ECVER": car_cen_v})  # 카중심:세로
        elif size_name == "카바닥(세로)":
            car_fro_y = min(int(pt1[1]), int(pt2[1]))
            car_rear_y = max(int(pt1[1]), int(pt2[1]))
            car_ee = int(ent_group["car_center"][1]) - car_fro_y
            floor_plan_data.update({"EL_ECEE": car_ee})  # CAR;EE
            ent_group.update({"car_rear_y": car_rear_y})

        if size_name in dim_name.keys():
            floor_plan_data.update({dim_name[size_name]: size})

    if len(ent_group["LAD-OPB"]) > 1:
        for opb_ent in ent_group["LAD-OPB"]:
            opb_x_cdnt = int(opb_ent.InsertionPoint[0])
            if opb_x_cdnt > hoist_rgt_x:  # 승강로 외부에 있다면 삭제
                ent_group["LAD-OPB"].remove(opb_ent)

    if len(ent_group["LAD-OPB"]) > 1:
        dis_opb_cnt = 0
        dis_opb_ents = []
        for opb_ent in ent_group["LAD-OPB"]:
            if opb_ent.EffectiveName == "LAD-OPB-DISABLED":
                dis_opb_ents.append(opb_ent)
                if dis_opb_cnt == 0:
                    dis_opb_cnt = 1
                else:
                    dis_opb_cnt = dis_opb_cnt + 1
        if dis_opb_ents:
            floor_plan_data.update({"EL_BHOPBQ": dis_opb_cnt})
            for dis_opb_ent in dis_opb_ents:
                ent_group["LAD-OPB"].remove(dis_opb_ent)

    opbs = {}
    for opb_ent in ent_group["LAD-OPB"]:
        opb_rotate = opb_ent.Rotation
        opb_x_cdnt = opb_ent.InsertionPoint[0]
        opb_y_cdnt = opb_ent.InsertionPoint[1]
        if opb_y_cdnt < ent_group["car_center"][1]:  # 카중심보다 밑에 있을 떄
            if opb_rotate == 0:
                if opb_x_cdnt < ent_group["car_center"][0]:
                    opb_pst = "R"  # RIGHT
                    opb_open = "CO"
                elif opb_x_cdnt > ent_group["car_center"][0]:
                    opb_pst = "L"  # LEFT
                    opb_open = "SOR"
            elif opb_rotate > 0:
                if opb_x_cdnt < ent_group["car_center"][0]:
                    opb_pst = "SR"  # RIGHT(측벽)
                    opb_open = "SOR"
                elif opb_x_cdnt > ent_group["car_center"][0]:
                    opb_pst = "SL"  # LEFT(측벽)
                    opb_open = "CO"
        elif opb_y_cdnt == ent_group["car_center"][1]:
            if opb_x_cdnt < ent_group["car_center"][0]:
                opb_pst = "SR"  # RIGHT(측벽)
                opb_open = "CO"
            elif opb_x_cdnt > ent_group["car_center"][0]:
                opb_pst = "SL"  # LEFT(측벽)
                opb_open = "CO"
        if len(ent_group["LAD-OPB"]) == 1:
            floor_plan_data.update({"EL_EOPBP": opb_pst, "EL_BMOPBO": opb_open})  # OPB 위치, MAIN OPB OPEN
        elif len(ent_group["LAD-OPB"]) == 2:
            if len(opbs) < 2:
                opbs.update({opb_ent.InsertionPoint[0]: [opb_pst, opb_open]})
            elif len(opbs) == 2:
                opbs = sorted(opbs.items())
                if "S" not in opbs[0][1][0] and "S" not in opbs[1][1][0]:
                    floor_plan_data.update({"EL_EOPBP": "A", "EL_BMOPBO": opbs[0][1][1], "EL_BSOPBO": opbs[1][1][1]})
                elif "S" in opbs[0][1][0] and "S" in opbs[1][1][0]:
                    floor_plan_data.update({"EL_EOPBP": "SA", "EL_BMOPBO": opbs[0][1][1], "EL_BSOPBO": opbs[1][1][1]})
                else:
                    floor_plan_data.update({"EL_EOPBP": "OPB 위치 확인이 필요합니다. MAIN OPB 위치 : " + opbs[0][1][
                        0] + ", SUB OPB 위치 : " + opbs[1][1][0], "EL_BMOPBO": opbs[0][1][1], "EL_BSOPBO": opbs[1][1][1]})

    for cwt_ent in ent_group["LAD-CWT"]:
        if "BRAKET" not in cwt_ent.EffectiveName:
            cwt_x_cdnt = round(cwt_ent.InsertionPoint[0])
            cwt_y_cdnt = round(cwt_ent.InsertionPoint[1])
            rope_x = abs(int(ent_group["car_center"][0] - cwt_x_cdnt))
            rope_y = abs(int(ent_group["car_center"][1] - cwt_y_cdnt))
            floor_plan_data.update({"EL_EPPX":rope_x, "EL_EPPY":rope_y})
            if rope_x < rope_y:  # 후락
                for cwt_prt in cwt_ent.GetDynamicBlockProperties():
                    if cwt_prt.propertyname == "@HEIGHT-T":
                        weight_t = cwt_prt.value  # subweight 상단폭
                    elif cwt_prt.propertyname == "@HEIGHT-B":
                        weight_b = cwt_prt.value  # subweight 하단폭
                    if "weight_t" in locals() and "weight_b" in locals():
                        break
                weight_w = int(weight_t + weight_b)
                floor_plan_data.update({"EL_ECWTP": "R"})  # CWT 위치 : REAR
            elif rope_x > rope_y:  # 횡락
                for cwt_prt in cwt_ent.GetDynamicBlockProperties():
                    if cwt_prt.propertyname == "@WIDTH-L":
                        weight_l = cwt_prt.value  # subweight 좌측폭
                    elif cwt_prt.propertyname == "@WIDTH-R":
                        weight_r = cwt_prt.value # subweight 우측폭
                    if "weight_l" in locals() and "weight_r" in locals():
                        break
                weight_w = int(weight_l + weight_r)
                if cwt_x_cdnt < int(ent_group["car_center"][0]):
                    cwt_pst = "R/L"  # FRONT, REAR 구분 필요
                elif cwt_x_cdnt > int(ent_group["car_center"][0]):
                    cwt_pst = "R/R"  # FRONT, REAR 구분 필요
                floor_plan_data.update({"EL_ECWTP": cwt_pst})
            floor_plan_data.update({"EL_ECWW": weight_w})  # CWT;WEIGHT폭

    for rail_ent in ent_group["LAD-RAIL"]:
        rail_x_cdnt = int(rail_ent.InsertionPoint[0])
        rail_y_cdnt = int(rail_ent.InsertionPoint[1])
        if rail_y_cdnt == int(ent_group["car_center"][1]):
            car_rail_spec = re.findall("(\d+)K", rail_ent.EffectiveName)[0]
            floor_plan_data.update({"EL_ECRL": car_rail_spec})  # CAR;RAIL(K)
            if rail_x_cdnt < ent_group["car_center"][0]: # right rail
                for rail_prt in rail_ent.GetDynamicBlockProperties():
                    if rail_prt.PropertyName == "@P1 Y":
                        rail_size = abs(rail_prt.Value)
                        rail_h1 = int((rail_x_cdnt - rail_size - 3) - hoist_lft_x)
                        floor_plan_data.update({"EL_ERBH1":rail_h1})
                        pass
            elif rail_x_cdnt > ent_group["car_center"][0]: # left rail
                for rail_prt in rail_ent.GetDynamicBlockProperties():
                    if rail_prt.PropertyName == "@P1 Y":
                        rail_size = abs(rail_prt.Value)
                        rail_h2 = int(hoist_rgt_x - (rail_x_cdnt+rail_size+3))
                        floor_plan_data.update({"EL_ERBH2":rail_h2})
                        pass
        else:
            cwt_rail_spec = re.findall("(\d+)K", rail_ent.EffectiveName)[0]
            floor_plan_data.update({"EL_ECWRL": cwt_rail_spec})  # CWT;RAIL(K)
            if rail_y_cdnt > car_rear_y:#후락
                rail_h3 = int(hoist_rear_y - rail_y_cdnt)
            else:
                floor_plan_data.update({"EL_ERBAG":floor_plan_data["EL_EPPY"]})
                if rail_x_cdnt < ent_group["car_center"][0]: # 좌락
                    rail_h3 = int(rail_x_cdnt - hoist_lft_x)
                elif rail_x_cdnt < ent_group["car_center"][0]: # 우락
                    rail_h3 = int(hoist_rgt_x - rail_x_cdnt)
            floor_plan_data.update({"EL_ERBH3":rail_h3})


    gov_ent = ent_group["LAD-GOV"][0]
    gov_x_cdnt = int(gov_ent.InsertionPoint[0])
    gov_y_cdnt = int(gov_ent.InsertionPoint[1])
    gov_y_gap = int(ent_group["car_center"][1]) - gov_y_cdnt
    car_cc = abs(gov_y_gap)
    floor_plan_data.update({"EL_ECCC": car_cc})  # CAR;CC
    if gov_y_gap < 0:
        if gov_x_cdnt < int(ent_group["car_center"][1]):
            floor_plan_data.update({"EL_ECGP": "R/L"})  # REAR & LEFT
        else:
            floor_plan_data.update({"EL_ECGP": "R/R"})  # REAR & RIGHT
    else:
        if gov_x_cdnt < int(ent_group["car_center"][1]):
            floor_plan_data.update({"EL_ECGP": "F/L"})  # FRONT & LEFT
        else:
            floor_plan_data.update({"EL_ECGP": "F/R"})  # FROTN & RIGHT

    if ent_group["LAD-CP"] == None:
        pass
    else:
        cp_ent = ent_group["LAD-CP"][0]
        if cp_ent.EffectiveName == "LAD-CP" or cp_ent.EffectiveName == "LAD-CP-DOOR":  # 승강장 jamb 취부형 제어반
            for cp_prt in cp_ent.GetDynamicBlockProperties():
                if cp_prt.propertyname == "@CASE-L":
                    case_l = cp_prt.value
                elif cp_prt.propertyname == "@CASE-R":
                    case_r = cp_prt.value
                if "case_l" in locals() and "case_r" in locals():
                    break
            sj = int(case_l + case_r)
            floor_plan_data.update({"EL_EMRLCJW": sj})  # MRL;CP JAMB 폭(SJ)
            if cp_ent.EffectiveName == "LAD-CP":
                cp_type = "J"
            elif cp_ent.EffectiveName == "LAD-CP-DOOR":
                cp_type = "C"
            if cp_ent.InsertionPoint[0] < ent_group["platform_cp"][0]:
                cp_pst = "L"
            else:
                cp_pst = "R"
            floor_plan_data.update({"EL_EMRLCJ": cp_type + cp_pst})  # MRL;CP JAMB TYPE
        elif cp_ent.EffectiveName != "LAD-CP-AC":  # 승강로 제어반
            floor_plan_data.update({"EL_EMRLHSCP": "Y"})
            cp_x_cdnt = cp_ent.InsertionPoint[0]
            cp_y_cdnt = cp_ent.InsertionPoint[1]
            if cp_y_cdnt > ent_group["car_rear_y"]:
                floor_plan_data.update({"승강로 제어반 위치": "REAR"})
            elif cp_y_cdnt > ent_group["car_center"][1]:
                if cp_x_cdnt < ent_group["car_center"][0]:
                    floor_plan_data.update({"승강로 제어반 위치": "R/R"})
                else:
                    floor_plan_data.update({"승강로 제어반 위치": "R/L"})
            elif cp_y_cdnt < ent_group["car_center"][1]:
                if cp_x_cdnt < ent_group["car_center"][0]:
                    floor_plan_data.update({"승강로 제어반 위치": "F/R"})
                else:
                    floor_plan_data.update({"승강로 제어반 위치": "F/L"})
    chk_spec = {}
    for spec_code, val in spec_data.items():
        if spec_code in floor_plan_data.keys():
            if floor_plan_data[spec_code] != val:
                chk_spec.update({spec_code:val})
        else:
            floor_plan_data.update({spec_code:val})

    return floor_plan_data, chk_spec


def get_mr_data():
    mr_data = {}
    for dim_ent in ent_group["DIM_ENT"]:
        del_s = dim_ent.TextOverride.replace(" ", "")
        size_name = re.findall("[가-힣]+", del_s)[0]
        Xdata = dim_ent.GetXData("", "Type", "Data")
        pt1 = Xdata[1][-2]
        pt2 = Xdata[1][-1]
        if int(pt1[0]) == int(pt2[0]):
            size_name = size_name + "(세로)"
        elif int(pt1[1]) == int(pt2[1]):
            size_name = size_name + "(가로)"
        else:
            gaps = {}
            gaps.update({abs(int(pt1[0]) - int(pt2[0])): "(가로)"})
            gaps.update({abs(int(pt1[1]) - int(pt2[1])): "(세로)"})
            size_name = size_name + gaps[round(dim_ent.Measurement)]

        size = round(dim_ent.Measurement)
        if size_name == "승강로내부(가로)":
            hoist_lft_x = min(int(pt1[0]), int(pt2[0]))
            if int(ent_group["hst_cent"][0]) < hoist_lft_x:
                size_name = "EL_EHH_CHK"
            else:
                size_name = "EL_EHH"
        elif size_name == "승강로내부(세로)":
            ver_dim_x = int(pt1[0])
            if ver_dim_x < int(ent_group["hst_cent"][0]):
                size_name = "EL_EHV"
            elif ver_dim_x > int(ent_group["hst_hole_cent"][0]):
                size_name = "EL_EHV_CHK"
            else:
                gap_hoistway = abs(ver_dim_x - int(ent_group["hst_cent"][0]))
                gap_hoistway_hole = abs(ver_dim_x - int(ent_group["hst_hole_cent"][0]))
                if min(gap_hoistway, gap_hoistway_hole) == gap_hoistway:
                    size_name = "EL_EHV"
                else:
                    size_name = "EL_EHV_CHK"
        mr_data.update({size_name: int(size)})

    if mr_data["EL_EHH"] == mr_data["EL_EHH_CHK"]:
        del mr_data["EL_EHH_CHK"]
    if mr_data["EL_EHV"] == mr_data["EL_EHV_CHK"]:
        del mr_data["EL_EHV_CHK"]

    if ent_group["LAD-HATCH"] != None:
        mr_data.update({"EL_DMRCP":"Y"})

    if ent_group["LAD-CP"] != None:
        cp_ent = ent_group["LAD-CP"][0]
        for cp_att in cp_ent.GetAttributes():
            if cp_att.TagString == "@TEXT":
                cp_cdnt = cp_att.TextAlignmentPoint

        if ent_group["LAD-HBTN-HOLE"] != None:
            duct_ent = ent_group["LAD-HBTN-HOLE"][0]
            hole_ent_x_gap = int(ent_group["hst_cent"][0] - ent_group["hst_hole_cent"][0])
            duct_x = int(duct_ent.InsertionPoint[0]) + hole_ent_x_gap
            cp_duct_x = abs(int(cp_cdnt[0]) - duct_x)
            cp_duct_y = abs(int(cp_cdnt[1] - duct_ent.InsertionPoint[1]))
            cp_to_duct = round(cp_duct_x + cp_duct_y, -3) + 1250
            mr_data.update({"EL_EDTA": int(cp_to_duct)})

        if ent_group["LAD-TM"] != None:
            tm_ent = ent_group["LAD-TM"][0]
            for tm_prt in tm_ent.GetDynamicBlockProperties():
                if tm_prt.propertyname == "@PP":
                    mr_data.update({"EL_EPPY": int(tm_prt.value)})
                    break
            cp_tm_x = abs(int(cp_cdnt[0] - tm_ent.InsertionPoint[0]))
            cp_tm_y = abs(int(cp_cdnt[1] - tm_ent.InsertionPoint[1]))
            cp_to_tm = round(cp_tm_x + cp_tm_y + 1500, -3) + 1000
            mr_data.update({"EL_EDTB": int(cp_to_tm)})

        if ent_group["door_cdnt"] != None:
            cp_door_x = abs(int(cp_cdnt[0] - ent_group["door_cdnt"][0]))
            cp_door_y = abs(int(cp_cdnt[1] - ent_group["door_cdnt"][1]))
            cp_to_pwr = round(cp_door_x + cp_door_y + 1650, -3) + 1000
            mr_data.update({"EL_EDTC": int(cp_to_pwr)})

        if ent_group["LAD-GOV"] != None:
            for gov_ent in ent_group["LAD-GOV"]:
                if gov_ent.EffectiveName[-1] == "H":
                    for gov_h_prt in gov_ent.GetDynamicBlockProperties():
                        if gov_h_prt.propertyname == "@DIST":
                            gov_name = "EL_ECGV_CHK"
                            gov_spec = "DG" + str(int(gov_h_prt.value))
                            break
                else:
                    gov_y_cdnt = gov_ent.InsertionPoint[1]
                    for gov_prt in gov_ent.GetDynamicBlockProperties():
                        if gov_prt.propertyname == "@DIST":
                            gov_name = "EL_ECGV"
                            gov_spec = "DG" + str(int(gov_prt.value))
                            break
                    cp_gov_x = abs(int(cp_cdnt[0] - gov_ent.InsertionPoint[0]))
                    cp_gov_y = abs(int(cp_cdnt[1] - (gov_y_cdnt + gov_prt.value)))
                    gov_cc = abs(int(ent_group["hst_cent"][1] - gov_ent.InsertionPoint[1]))
                    cp_to_gov = round(cp_gov_x + cp_gov_y + 150, -3) + 1000
                    mr_data.update({"EL_EDTE": cp_to_gov, "EL_ECCC": int(gov_cc)})
                mr_data.update({gov_name: gov_spec})

        if mr_data["EL_ECGV"] == mr_data["EL_ECGV_CHK"]:
            del mr_data["EL_ECGV_CHK"]

    return mr_data


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



def get_table_cdnt(table_name):

    palette_area = ent_group["palette_area"]

    if table_name == "floor":
        floor_table_y_cdnt = ent_group["floor_table_y_cdnt"]
        if floor_table_y_cdnt != None:
            for poly_ent in ent_group["Polyline"]:  # 층 테이블 좌표 구하기
                if floor_table_y_cdnt in poly_ent.Coordinates:
                    fl_st_x_cdnt = poly_ent.Coordinates[0]
                    fl_ed_x_cdnt = poly_ent.Coordinates[2]
                    fl_st_y_cdnt = poly_ent.Coordinates[1]
                    fl_ed_y_cdnt = poly_ent.Coordinates[5]
                    table_cdnt = [fl_st_x_cdnt, fl_ed_x_cdnt, fl_st_y_cdnt, fl_ed_y_cdnt]

        elif floor_table_y_cdnt == None:
            text_insert_cdnt = {}
            for text_ent in ent_group["Text"]:  # 층 테이블의 데이터 좌표 구하기
                if text_ent.TextString == "FL / ST":
                    text_base_y_cdnt = text_ent.TextAlignmentPoint[1]
                elif text_ent.TextString == "층":
                    text_insert_cdnt.update({text_ent.TextAlignmentPoint[1]: text_ent.TextAlignmentPoint[0]})  # text Y:X(Y값은 상이함)
            text_base_x_cdnt = text_insert_cdnt[text_base_y_cdnt] # "FL/ST" text와 같은 Y좌표를 가진 "층" text의 X좌표 구하기

            for poly_ent in ent_group["Polyline"]:  # 데이터와 가까운 Line 좌표 구하기
                if poly_ent.Layer == "LAD-OUTLINE":
                    x_gap = abs(text_base_x_cdnt - poly_ent.Coordinates[0])
                    y_gap = abs(text_base_y_cdnt - poly_ent.Coordinates[1])
                    gap = x_gap + y_gap
                    if palette_area > gap:  # 방화도어 TABLE과 하고 겹치지 않도록 gap 비교
                        palette_area = gap
                        fl_st_x_cdnt = poly_ent.Coordinates[0]
                        fl_ed_x_cdnt = poly_ent.Coordinates[2]
                        fl_st_y_cdnt = poly_ent.Coordinates[1]
                        fl_ed_y_cdnt = poly_ent.Coordinates[5]
                        table_cdnt = [fl_st_x_cdnt, fl_ed_x_cdnt, fl_st_y_cdnt, fl_ed_y_cdnt]

    if table_name == "fire_door":
        fdoor_table_y_cdnt = ent_group["fdoor_table_y_cdnt"]
        if fdoor_table_y_cdnt != None:
            for poly_ent in ent_group["Polyline"]:  # 방화도어 TABLE 좌표 구하기
                if fdoor_table_y_cdnt in poly_ent.Coordinates:
                    fd_st_x_cdnt = poly_ent.Coordinates[2]
                    fd_ed_x_cdnt = poly_ent.Coordinates[0]
                    fd_st_y_cdnt = poly_ent.Coordinates[5]
                    fd_ed_y_cdnt = poly_ent.Coordinates[1]
                    table_cdnt = [fd_st_x_cdnt, fd_ed_x_cdnt, fd_st_y_cdnt, fd_ed_y_cdnt]


        elif fdoor_table_y_cdnt == None:
            for text_ent in ent_group["Text"]:  # 방화도어 TABLE의 데이터 좌표 구하기
                if text_ent.TextString == "방화도어 유무":
                    text_base_x_cdnt = text_ent.TextAlignmentPoint[0]
                    text_base_y_cdnt = text_ent.TextAlignmentPoint[1]

            for poly_ent in ent_group["Polyline"]:  # 데이터와 가까운 Line 좌표 구하기
                if poly_ent.Layer == "LAD-OUTLINE":
                    x_gap = abs(text_base_x_cdnt - poly_ent.Coordinates[2])
                    y_gap = abs(text_base_y_cdnt - poly_ent.Coordinates[1])
                    gap = x_gap + y_gap
                    if palette_area > gap:  # 방화도어 table 하고 겹치지 않도록 gap 비교
                        palette_area = gap
                        fd_st_x_cdnt = poly_ent.Coordinates[2]
                        fd_ed_x_cdnt = poly_ent.Coordinates[0]
                        fd_st_y_cdnt = poly_ent.Coordinates[5]
                        fd_ed_y_cdnt = poly_ent.Coordinates[1]
                        table_cdnt = [fd_st_x_cdnt, fd_ed_x_cdnt, fd_st_y_cdnt, fd_ed_y_cdnt]

    return table_cdnt


def get_floor_height_data():

    s_x_cdnt, e_x_cdnt, s_y_cdnt, e_y_cdnt = get_table_cdnt("floor")

    table_data = {}
    x_cdnt_list = []
    floor_height_data = {}
    for text_ent in ent_group["Text"]:
        x_cdnt = text_ent.InsertionPoint[0]
        y_cdnt = text_ent.InsertionPoint[1]
        if x_cdnt > s_x_cdnt and x_cdnt < e_x_cdnt and y_cdnt < s_y_cdnt and y_cdnt > e_y_cdnt:
            table_data.update({text_ent.TextAlignmentPoint: text_ent.TextString})  # 좌표안에 있는 테이블에 있는 모든 TEXT get
            if text_ent.TextString == "층":  # 윗행(층)과 아래행(층고) 나누는 기준
                floor_y_cdnt = text_ent.TextAlignmentPoint[1]
            elif text_ent.TextString == "층고":  # 윗행(층)과 아래행(층고) 나누는 기준
                floor_hei_y_cdnt = text_ent.TextAlignmentPoint[1]
            elif text_ent.TextString == "FL / ST":  # 층수 구하기
                flst_x_cdnt = text_ent.TextAlignmentPoint[0]
            else:
                x_cdnt_list.append(text_ent.TextAlignmentPoint[0])

    x_cdnt_list = list(set(x_cdnt_list))  # 중복 좌표 삭제
    x_cdnt_list.remove(flst_x_cdnt)
    x_cdnt_list.sort()  # x좌표 순서대로 정리

    for x in x_cdnt_list:
        floor_text = table_data[(x, floor_y_cdnt, 0.0)] # 층표기 구하기
        floor_mark_list = special_str_split(floor_text)
        floor_height = table_data[(x, floor_hei_y_cdnt, 0.0)] # 층고 구하기
        for floor_mark in floor_mark_list:
            floor_height_data.update({floor_mark:floor_height})

    floor_data = {}
    fl_st_data = table_data[(flst_x_cdnt, floor_hei_y_cdnt, 0.0)]
    floor_data.update({"EL_AFQ" : re.findall("(\d+)/", fl_st_data)[0]}) # 층수
    floor_data.update({"EL_ASTQ" : re.findall("/(\d+)", fl_st_data)[0]}) # 정지층수
    floor_data.update({"EL_EFHB" : list(floor_height_data.values())[0]}) # BOT 층고
    floor_data.update({"EL_EFHT" : list(floor_height_data.values())[-2]}) # TOP-1 층고
    floor_data.update({"EL_EFHMAX" : max(table_data.values())}) # 최대 층고
    floor_data.update({"EL_EFHMIN" : min(table_data.values())}) # 최소 층고


    if ent_group["hoistway_info"]:
        hstw_ent = ent_group["hoistway_info"]
        hstw_att_name = {"OH":"EL_EHO", "HH":"EL_EHTH", "TH":"EL_EHTRH", "PIT":"EL_EHP"}
        for hstw_att in hstw_ent.GetAttributes():
            if hstw_att.TagString in hstw_att_name:
                floor_data.update({hstw_att_name[hstw_att.TagString]:hstw_att.TextString})
            elif hstw_att.TagString == "@BRAKET":
                bracket_q = re.findall("(\d+)EA", hstw_att.TextString)[0]
                floor_data.update({"EL_ERBQ":int(bracket_q)+4})

    return floor_height_data, floor_data


def get_fire_door():

    s_x_cdnt, e_x_cdnt, s_y_cdnt, e_y_cdnt = get_table_cdnt("fire_door")

    table_data = {}
    x_cdnt_list = []
    fire_door_floor = {}
    for text_ent in ent_group["Text"]:
        x_cdnt = text_ent.InsertionPoint[0]
        y_cdnt = text_ent.InsertionPoint[1]
        if x_cdnt > s_x_cdnt and x_cdnt < e_x_cdnt and y_cdnt < s_y_cdnt and y_cdnt > e_y_cdnt:
            table_data.update({text_ent.TextAlignmentPoint: text_ent.TextString}) # 좌표안에 있는 테이블에 있는 모든 TEXT get
            if text_ent.TextString == "층":  # 윗행(층)과 아래행(층고) 나누는 기준
                floor_y_cdnt = text_ent.TextAlignmentPoint[1]
            elif "방화도어" in text_ent.TextString:  # 윗행(층)과 아래행(층고) 나누는 기준
                fire_door_y_cdnt = text_ent.TextAlignmentPoint[1]
            else:
                x_cdnt_list.append(text_ent.TextAlignmentPoint[0])

    x_cdnt_list = list(set(x_cdnt_list)) # 중복 좌표 삭제
    x_cdnt_list.sort() # x좌표 순서대로 정리


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
            fire_door_floor.update({floor_mark: fire_door})

    return fire_door_floor



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
            for item, spec in hall_items.items():
                floor_data.loc[app_floor, item] = spec
                floor_data.loc[app_floor, "위치"] = "기준층"
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
                for item, spec in hall_items.items():
                    floor_data.loc[app_floor, item] = spec

    return floor_data


def get_hall_items(hall_odr, jamb_spec):
    cable_x_cdnt = {}
    for ent_name, ents in ent_group.items():
        name = ent_name.replace("-", "_")
        if ents == None:  # 객체가 없는 데이터 제외
            globals()[name] = None
        elif ent_name == "LAD-OPEN-HOLE":
            for cable_hole_cdnt, ent in ents.items():
                cable_x = ent.InsertionPoint[0]
                for att in ent.GetDynamicBlockProperties():
                    if att.propertyname == "@OFFSET-Y" and att.value < 1400:  # hole 높이
                        cable_x_cdnt.update({cable_x: "HBTN"})
                    elif att.propertyname == "@OFFSET-Y" and att.value >= 1400:
                        cable_x_cdnt.update({cable_x: "OTHER"})
        elif ent_name == "LAD-EMCY-SWITCH":
            firesw_x_cdnt = list(ents.keys())[0]
            jamb_cdnt_list = list(ent_group["LAD-OPEN-AC"].keys())
            jamb_cdnt_list.sort()
            firesw_app_jamb_cdnt = min(jamb_cdnt_list, key=lambda x: abs(x - firesw_x_cdnt))
            firesw_app_jamb_ord = jamb_cdnt_list.index(firesw_app_jamb_cdnt)
            globals()[name] = ents.values()
        else:
            sort_ents = []
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
        app_jamb = "JAMB(" + str(hall_odr + 1) + ");"
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
            lant_app_jamb_cdnt = min(jamb_cdnt_list, key=lambda x: abs(x - lant_cdnt))
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

            if firesw_app_jamb_cdnt - firesw_x_cdnt < 0:
                firesw_pst = "RIGHT"
            else:
                firesw_pst = "LEFT"

            if firesw_pst in box_hole:
                firesw = "BOX"
            elif firesw_x_cdnt in cable_x_cdnt.keys() and "OTHER" in cable_x_cdnt.values():
                firesw = "BOXLESS"
            else:
                firesw = "소방스위치 type을 확인할 수 없습니다."
            floor_items.update({"소방스위치": firesw})
        else:
            pass

    if hall_odr == len(LAD_DOOR_JAMB) - 1 and LAD_REMOTE_CP != None:  # 마지막 jamb(=최상층 jamb)일 떄
        remote_cp = "Y"
        floor_items.update({"분리형 보조제어반": remote_cp})

    return floor_items


def get_proj_data(proj_no):
    df_spec = get_pdm_spec(proj_no)

    doc = layout_find(proj_no[:6], "H")
    get_entity("H")
    floor_plan_data, chk_data = get_floor_plan_data()
    df_spec.loc["평면도"]= floor_plan_data
    if len(chk_data):
        df_spec.loc["CHK_평면도"] = chk_data


    doc = layout_find(proj_no[:6], "M")
    get_entity("M")
    mr_data = get_mr_data()  # mr일 때 실행하기, mrl은 기계담당자 확인 후 로직 작성 요망
    df_spec.loc["기계실_배치도"] = mr_data


    doc = layout_find(proj_no[:6], "S")
    get_entity("S")

    floor_height_data, floor_data = get_floor_height_data()

    df_spec.loc["승강로_단면도"] = floor_data

    df_floor = pd.DataFrame(floor_height_data.values(), floor_height_data.keys(), columns=["층고"]) # 층고 table 작성

    fire_door_data = get_fire_door() # 층고 table에 방화도어 정보 추가
    df_floor["방화도어"] = fire_door_data

    doc = layout_find(proj_no[:6], "E")
    get_entity("E")
    df_floor = get_hall_data(df_floor)

    df_spec.fillna("", inplace=True)
    df_floor.fillna("", inplace=True)

    return df_spec, df_floor


proj_data, floor_data = get_proj_data("188899L01")
print(proj_data)
print("걸린 시간 : ", time.time() - start)