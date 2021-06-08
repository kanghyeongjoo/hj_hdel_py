import pandas as pd
import win32com.client
import re
import time
import requests
from bs4 import BeautifulSoup

start = time.time()

# get ouid
def get_ouid(proj_no):
    ouid_url = "http://plm.hdel.co.kr/jsp/help/ouidList.jsp?md%24number={}".format(proj_no)
    ouid_data = requests.get(ouid_url)
    ouid_address = BeautifulSoup(ouid_data.content, "html.parser")
    ouid_get = ouid_address.find("form").next_sibling.strip()
    return ouid_get #get spec을 진행할 때 sepc url을 여기서 return하는 것도 고려해볼 것.

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
            spec_get.update({code:val})
    return spec_get

def get_pdm_spec(proj_no):
    ouid = get_ouid(proj_no)
    pdm_spec = get_spec(ouid)
    df_pdm_spec = pd.DataFrame([pdm_spec])
    df_pdm_spec = df_pdm_spec.transpose()
    df_pdm_spec.columns = ["PDM_DATA"]
    return df_pdm_spec


def get_entity(layout_kind):
    acad = win32com.client.Dispatch("AutoCAD.Application")
    global doc, ent_group
    doc = acad.ActiveDocument
    if layout_kind == "S":
        ent_blo_name = ["LAD-RAIL", "LAD-OPB", "LAD-CWT", "LAD-GOV", "DIM_ENT", "LAD-CWT", "LAD-CP"]
        ent_group = dict.fromkeys(ent_blo_name)
        for entity in doc.ModelSpace:
            if entity.EntityName == "AcDbBlockReference" and entity.EffectiveName =="LAD-HOISTWAY-HP-SC":
                ent_group.update({"hoistway_m":"CEMEN"})
                entity.explode()
            elif entity.EntityName == "AcDbBlockReference" and entity.EffectiveName =="LAD-HOISTWAY-HP-SS":
                ent_group.update({"hoistway_m": "ST"})
                entity.explode()
            elif entity.EntityName == "AcDbBlockReference" and entity.EffectiveName =="LAD-CAR-1SCO":
                ent_group.update({"car_center":entity.InsertionPoint})
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

    return ent_group


def get_floor_plan_data(ent_group):

    spec_ent = ent_group["spec_data"]
    spec_data = {}
    tag_name = {"@GOVERNOR":"EL_ECGV", "@CAR_SAFETY":"EL_ECSF", "@TM_TYPE":"EL_ETM"}  # 특성코드와 dic형태로 매칭해주는 것도 생각해볼
    trs_tag_name = {"@BALANCE":"EL_ECBA", "@NO":"EL_ECN", "@V_SPEC":["EL_AVOLT","EL_ALI","EL_AHZ"] , "@DRIVE_TYPE":"EL_ADRV", "@DRIVE":"EL_ATYP", "@SPEED":"EL_ASPD", "@CAPA":["EL_AMAN", "EL_ACAPA"], "@USE":"EL_AUSE", "@DOOR_DRIVE":"EL_AOPEN", "@MOTOR_CAPA":"EL_ETMM",
                    "@ROPE_SPEC":["EL_ERPD", "EL_ERPW", "EL_ERPR"], "@DOOR_SIZE":["EL_ECJJ", "EL_ECHH"], "@CAR_SIZE":["EL_ECCA", "EL_ECCB", "EL_ECCH"],"@CB_TYPE":"EL_DURTB"} # 변환이 필요한 코드
    for spec_att in spec_ent.GetAttributes():
        if spec_att.TagString in tag_name.keys():
            spec_data.update({tag_name[spec_att.TagString]:spec_att.TextString})
        elif spec_att.TagString in trs_tag_name.keys():
            if spec_att.TagString == "@BALANCE":
                att_value = re.findall("\d+", spec_att.TextString)[0]
                spec_data.update({trs_tag_name[spec_att.TagString]:att_value})
            elif spec_att.TagString == "@NO":
                att_value = re.findall("\d+", spec_att.TextString)
                if len(att_value) == 1:
                    spec_data.update({trs_tag_name[spec_att.TagString]: att_value[0]})
                else:
                    spec_data.update({trs_tag_name[spec_att.TagString]:att_value})
            elif spec_att.TagString == "@V_SPEC":
                att_value = spec_att.TextString.lower().replace(" ", "")
                att_value_list = re.findall("\d+(?=v)|\d+(?=hz)", att_value)
                for idx in range(len(att_value_list)):
                    spec_data.update({trs_tag_name[spec_att.TagString][idx]:att_value_list[idx]})
            elif spec_att.TagString == "@DRIVE_TYPE":
                car_oper_type = re.findall("\d+", spec_att.TextString)
                att_value = car_oper_type[0]+"C"+car_oper_type[1]+"BC"
                spec_data.update({trs_tag_name[spec_att.TagString]: att_value})
            elif spec_att.TagString == "@DRIVE":
                if "WBSS" in spec_att.TextString:
                    att_value = "WBSS2_(SSVF)"
                elif "LXVF" in spec_att.TagString or "WBLX" in spec_att.TagString:
                    att_value = "WBLX1"
                else:
                    att_value = spec_att.TextString + "은 사양 추가 요청바랍니다."
                spec_data.update({trs_tag_name[spec_att.TagString]:att_value})
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
                pdm_drive = ["1SCO", "2SSO", "2SL", "2SR", "2SLR", "3SSO", "3SL", "3SR", "3SLR", "2SCO", "2UP", "2UL", "2UR",
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
                spec_data.update({trs_tag_name[spec_att.TagString]:u_bfr})
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
    hoistway_m = ent_group["hoistway_m"]
    floor_plan_data.update({"EL_EHM":hoistway_m})
    doc.SendCommand('setxdata ')
    dim_name = {"균형추레일간의거리(세로)":"EL_ECWBG", "승강로내부(세로)":"EL_EHV", "카바닥(세로)":"EL_ECBB", "카내부(세로)":"EL_ECCB", "카바닥(가로)":"EL_ECAA", "출입구유효폭(가로)":"EL_ECJJ", "카레일간의거리(가로)":"EL_ECBG", "승강로내부(가로)":"EL_EHH", "카내부(가로)":"EL_ECCA"}
    for dim_ent in ent_group["DIM_ENT"]:
        del_s = dim_ent.TextOverride.replace(" ", "")
        size_name = re.findall("[가-힣]+", del_s)[0]
        Xdata = dim_ent.GetXData("", "Type", "Data")
        pt1 = Xdata[1][-2]
        pt2 = Xdata[1][-1]
        if int(pt1[0]) == int(pt2[0]):
            size_name = size_name+"(세로)"
        elif int(pt1[1]) == int(pt2[1]):
            size_name = size_name + "(가로)"
        else:
            gaps = {}
            gaps.update({abs(int(pt1[0]) - int(pt2[0])) : "(가로)"})
            gaps.update({abs(int(pt1[1]) - int(pt2[1])) : "(세로)"})
            size_name = size_name + gaps[int(dim_ent.Measurement)]

        if size_name == "승강로내부(가로)":
            hoist_lft_x = min(int(pt1[0]), int(pt2[0]))
            hoist_rgt_x = max(int(pt1[0]), int(pt2[0]))
            car_cen_h = abs(hoist_lft_x - int(ent_group["car_center"][0]))
            floor_plan_data.update({"EL_ECHOR": car_cen_h}) #카중심:가로
        elif size_name == "승강로내부(세로)":
            hoist_fro_y = min(int(pt1[1]), int(pt2[1]))
            car_cen_v = abs(hoist_fro_y - int(ent_group["car_center"][1]))
            floor_plan_data.update({"EL_ECVER": car_cen_v}) #카중심:세로
        elif size_name == "카바닥(세로)":
            car_fro_y = min(int(pt1[1]), int(pt2[1]))
            car_rear_y = max(int(pt1[1]), int(pt2[1]))
            car_ee = int(ent_group["car_center"][1]) - car_fro_y
            floor_plan_data.update({"EL_ECEE": car_ee}) #CAR;EE
            ent_group.update({"car_rear_y":car_rear_y})

        if size_name in dim_name.keys():
            floor_plan_data.update({dim_name[size_name]: int(dim_ent.Measurement)})


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
        if opb_y_cdnt < ent_group["car_center"][1]: #카중심보다 밑에 있을 떄
            if opb_rotate == 0:
                if opb_x_cdnt < ent_group["car_center"][0]:
                    opb_pst = "R" # RIGHT
                    opb_open = "CO"
                elif opb_x_cdnt > ent_group["car_center"][0]:
                    opb_pst = "L" # LEFT
                    opb_open = "SOR"
            elif opb_rotate > 0:
                if opb_x_cdnt < ent_group["car_center"][0]:
                    opb_pst = "SR" # RIGHT(측벽)
                    opb_open = "SOR"
                elif opb_x_cdnt > ent_group["car_center"][0]:
                    opb_pst = "SL" # LEFT(측벽)
                    opb_open = "CO"
        elif opb_y_cdnt == ent_group["car_center"][1]:
            if opb_x_cdnt < ent_group["car_center"][0]:
                opb_pst = "SR" # RIGHT(측벽)
                opb_open = "CO"
            elif opb_x_cdnt > ent_group["car_center"][0]:
                opb_pst = "SL" # LEFT(측벽)
                opb_open = "CO"
        if len(ent_group["LAD-OPB"]) == 1:
            floor_plan_data.update({"EL_EOPBP": opb_pst, "EL_BMOPBO": opb_open}) # OPB 위치, MAIN OPB OPEN
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
                    floor_plan_data.update({"EL_EOPBP": "OPB 위치 확인이 필요합니다. MAIN OPB 위치 : " + opbs[0][1][0] + ", SUB OPB 위치 : " + opbs[1][1][0], "EL_BMOPBO": opbs[0][1][1], "EL_BSOPBO": opbs[1][1][1]})

    for cwt_ent in ent_group["LAD-CWT"]:
        if "BRAKET" not in cwt_ent.EffectiveName:
            cwt_x_cdnt = int(cwt_ent.InsertionPoint[0])
            cwt_y_cdnt = int(cwt_ent.InsertionPoint[1])
            if abs(cwt_x_cdnt - int(ent_group["car_center"][0])) < abs(cwt_y_cdnt - int(ent_group["car_center"][1])): #후락
                for cwt_prt in cwt_ent.GetDynamicBlockProperties():
                    if cwt_prt.propertyname == "@HEIGHT-T":
                        weight_t = cwt_prt.value # subweight 상단폭
                    elif cwt_prt.propertyname == "@HEIGHT-B":
                         weight_b = cwt_prt.value # subweight 하단폭
                weight_w = int(weight_t + weight_b)
                floor_plan_data.update({"EL_ECWTP": "R"}) #CWT 위치 : REAR
            elif abs(cwt_x_cdnt - int(ent_group["car_center"][0])) > abs(cwt_y_cdnt - int(ent_group["car_center"][1])): # 횡락
                for cwt_prt in cwt_ent.GetDynamicBlockProperties():
                    if cwt_prt.propertyname == "@WIDTH-L":
                        weight_l = cwt_prt.value # subweight 좌측폭
                    elif cwt_prt.propertyname == "@WIDTH-R":
                         weight_r = cwt_prt.value
                weight_w = int(weight_l + weight_r) # subweight 우측폭
                if cwt_x_cdnt < int(ent_group["car_center"][0]):
                    cwt_pst = "R/L" # FRONT, REAR 구분 필요
                elif cwt_x_cdnt > int(ent_group["car_center"][0]):
                    cwt_pst = "R/R" # FRONT, REAR 구분 필요
                floor_plan_data.update({"EL_ECWTP": cwt_pst})
            floor_plan_data.update({"EL_ECWW":weight_w}) # CWT;WEIGHT폭


    for rail_ent in ent_group["LAD-RAIL"]:
        rail_cdnt = rail_ent.InsertionPoint[1]
        if rail_cdnt == ent_group["car_center"][1]:
            car_rail_spec = re.findall("(\d+)K", rail_ent.EffectiveName)[0]
            floor_plan_data.update({"EL_ECRL":car_rail_spec}) # CAR;RAIL(K)
        else:
            cwt_rail_spec = re.findall("(\d+)K", rail_ent.EffectiveName)[0]
            floor_plan_data.update({"EL_ECWRL": cwt_rail_spec}) # CWT;RAIL(K)

    gov_ent = ent_group["LAD-GOV"][0]
    gov_x_cdnt = int(gov_ent.InsertionPoint[0])
    gov_y_cdnt = int(gov_ent.InsertionPoint[1])
    gov_y_gap = int(ent_group["car_center"][1]) - gov_y_cdnt
    car_cc = abs(gov_y_gap)
    floor_plan_data.update({"EL_ECCC": car_cc}) # CAR;CC
    if gov_y_gap < 0:
        if gov_x_cdnt < int(ent_group["car_center"][1]):
            floor_plan_data.update({"EL_ECGP" : "R/L"}) # REAR & LEFT
        else:
            floor_plan_data.update({"EL_ECGP" : "R/R"}) # REAR & RIGHT
    else:
        if gov_x_cdnt < int(ent_group["car_center"][1]):
            floor_plan_data.update({"EL_ECGP" : "F/L"}) # FRONT & LEFT
        else:
            floor_plan_data.update({"EL_ECGP" : "F/R"}) # FROTN & RIGHT


    cp_ent = ent_group["LAD-CP"][0]
    if cp_ent == None: #spec에서 얻은 기종으로 구분하여 pass할 것.
        pass
    elif cp_ent.EffectiveName == "LAD-CP" or cp_ent.EffectiveName == "LAD-CP-DOOR": # 승강장 jamb 취부형 제어반
        for cp_prt in cp_ent.GetDynamicBlockProperties():
            if cp_prt.propertyname == "@CASE-L":
                case_l = cp_prt.value
            elif cp_prt.propertyname == "@CASE-R":
                case_r = cp_prt.value
        sj = int(case_l + case_r)
        floor_plan_data.update({"EL_EMRLCJW":sj}) # MRL;CP JAMB 폭(SJ)
        if cp_ent.EffectiveName == "LAD-CP":
            cp_type = "J"
        elif cp_ent.EffectiveName == "LAD-CP-DOOR":
            cp_type = "C"
        if cp_ent.InsertionPoint[0] < ent_group["platform_cp"][0]:
            cp_pst = "L"
        else:
            cp_pst = "R"
        floor_plan_data.update({"EL_EMRLCJ": cp_type + cp_pst}) # MRL;CP JAMB TYPE
    elif cp_ent.EffectiveName != "LAD-CP-AC" :#승강로 제어반
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

    return spec_data, floor_plan_data


df_spec = get_pdm_spec("188402L01")
entity = get_entity("S")
floor_plan_spec, floor_plan_data = get_floor_plan_data(entity)

for code, val in floor_plan_spec.items():
    df_spec.loc[code, "승강로_평면도"] = val
df_spec = df_spec.fillna({"승강로_평면도":""})

for code, val in floor_plan_data.items():
    if df_spec.loc[code, "승강로_평면도"] == "":
        df_spec.loc[code, "승강로_평면도"] == val
    elif df_spec.loc[code, "승강로_평면도"] != val:
        df_spec.loc[code, "승강로_평면도(SPEC)"] = val

print(df_spec)

print("걸린 시간 : ", time.time() - start)
