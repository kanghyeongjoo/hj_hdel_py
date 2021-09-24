import win32com.client
import re
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time


class get_pdm_data:

    def __init__(self, projNo):
        self.projNo =  projNo
        self.el_ouid = self.get_ouid()
        self.pdm_data = self.get_data()

    def get_ouid(self):
        ouid = {}
        ouid_url = "http://plm.hdel.co.kr/jsp/help/projectouidList.jsp?md%24number={}".format(self.projNo)
        ouid_page = requests.get(ouid_url)
        chk_proj = ouid_page.text
        if "현장번호가 없음" in chk_proj:
            return "현장번호 없음"
        else:
            ouid_address = BeautifulSoup(ouid_page.content, "html.parser")
            get_ouids = ouid_address.findAll("br")
            for ouid_data in get_ouids:
                ouid_str = ouid_data.previousSibling.strip()
                if ouid_str[:2] == "el":
                    return ouid_str

    def get_data(self):
        spec_url = "http://plm.hdel.co.kr/jsp/plmetc/elvInfo/elvinfomation.jsp?cOuid=860c9bb8&iOuid={}".format(self.el_ouid)
        spec_data = requests.get(spec_url)
        spec_address = BeautifulSoup(spec_data.content, "html.parser")
        spec_list = spec_address.find_all("tr", "01-cell")
        pdm_spec = {}
        el_code_list = ["EL_A", "EL_B", "EL_C", "EL_D", "EL_E"]
        chk_name = ["삭제", "교체", "재사용", "전기", "기계"]
        floor_code_list = ["EL_AFF", "EL_ATF", "EL_CJM1F", "EL_CJM2F", "EL_CJM3F"]
        for spec in spec_list:
            code = spec.find_all("td")[3].get_text()
            if "TEXT" not in code and any(el_code in code for el_code in el_code_list):
                get_name = spec.find_all("td")[2].get_text()
                name = re.findall("\W+(.+)\r", get_name)
                if len(name) > 0 and all(chk_str not in name[0] for chk_str in chk_name):
                    val = spec.find_all("td")[4].get_text()
                    if val == "" or val == "0":
                        pdm_spec.update({code:[name[0],None]})
                    else:
                        if code in floor_code_list:
                            val = ",".join(convert_special_str(val).convert_str)
                        elif code == "EL_ECN" and val.isdigit():
                            val = int(val)
                        pdm_spec.update({code:[name[0],str(val)]})

        return pdm_spec


class convert_special_str:

    def __init__(self, floor_mark):
        self.floor_mark = floor_mark
        self.convert_str = self.special_str_split()

    def special_str_split(self):
        floor_mark_list = []
        if "," not in self.floor_mark and "." not in self.floor_mark:
            comma_split_floor = [self.floor_mark]
        elif "," in self.floor_mark or "." in self.floor_mark:
            comma_split_floor = re.split("[,.]", self.floor_mark)

        for split_floor in comma_split_floor:
            if "~" not in split_floor and "-" not in split_floor:
                floor_mark_list.append(split_floor)
            elif "~" in split_floor or "-" in split_floor:
                str_floor = re.findall("(\w+)\W", split_floor)[0]
                str_text = re.findall("(\D+)\d+", str_floor)
                end_floor = re.findall("\W(\w+)", split_floor)[0]
                end_text = re.findall("(\D+)\d+", end_floor)
                if str_floor == "F":
                    str_floor = "4"
                    str_chg_f = "Y"
                if end_floor == "F":
                    end_floor = "4"
                    end_chg_f = "Y"
                if not len(str_text):  # start 층표기에 B2~3과 같은 문자가 있는지 확인
                    st_no = int(str_floor)
                    end_no = int(end_floor) + 1
                    for floor in range(st_no, end_no):
                        floor_mark_list.append(str(floor))
                elif not len(end_text):  # start 층표기에는 문자가 있고, end 층표기에는 문자가 없을 때
                    text = str_text[0]
                    st_no = re.findall("\d+", str_floor)[0]
                    st_no = int(st_no)
                    end_no = int(end_floor) + 1
                    for floor in range(st_no, 0, -1):
                        floor_mark_list.append(text + str(floor))
                    for floor in range(1, end_no):
                        floor_mark_list.append(str(floor))
                elif len(end_text) > 0:  # start, end 모두 층표기에 문자가 있을 때
                    text = str_text[0]
                    st_no = re.findall("\d+", str_floor)[0]
                    st_no = int(st_no)
                    end_no = re.findall("\d+", end_floor)[0]
                    end_no = int(end_no) - 1
                    for floor in range(st_no, end_no, -1):
                        floor_mark_list.append(text + str(floor))
                if "str_chg_f" in locals() or "end_chg_f" in locals():
                    f_idx = floor_mark_list.index("4")
                    floor_mark_list.insert(f_idx, "F")
                    floor_mark_list.remove("4")

        return floor_mark_list


class get_layout_data():

    def __init__(self, projNo):
        self.acad = win32com.client.Dispatch("AutoCAD.Application")
        self.projNo = projNo[:6]
        self.document = self.find_layout()
        self.pdm_data = get_pdm_data(projNo).pdm_data
        self.floor_plan_data, self.floor_plan_chk = self.get_floor_plan_data()
        if self.machinroomType == "MRL":
            self.machinroom_data, self.machinroom_chk = self.get_mrl_data()
        elif self.machinroomType == "MR":
            self.machinroom_data = self.get_mr_data()
        self.section_data, self.floor_height_data, self.fire_door_data = self.get_section_data()
        self.hall_data, self.hall_item = self.get_hall_data()

    def organize_data(self):
        data_table = pd.DataFrame(self.pdm_data, index=["특성값", "PDM DATA"])
        add_data = {"승강로 평면도":self.floor_plan_data, "기계실 배치도":self.machinroom_data, "출입구 의장도":self.hall_data, "승강로 단면도":self.section_data}
        self.all_table = data_table.transpose()
        for kind, data in add_data.items():
            self.all_table[kind] = self.all_table.index.map(data)
        self.all_table.rename_axis("특성 코드", inplace=True)
        self.all_table.reset_index(inplace=True)
        nunique_idx = self.all_table.index[self.all_table.nunique(axis=1) > 3] # 3인 이유는 특성코드, 특성값 2개의 값을 제외한 유일한 값을 찾기 위함
        self.all_table["비교 결과"] = "True"
        self.all_table.loc[nunique_idx, "비교 결과"] = "False"
        self.fill_table = self.all_table.copy()
        self.fill_table.dropna(subset=["PDM DATA", "승강로 평면도", "기계실 배치도", "출입구 의장도", "승강로 단면도"], thresh=1, inplace=True)
        self.all_table.fillna("", inplace=True)
        self.fill_table.fillna("", inplace=True)
        self.false_table = self.fill_table[self.fill_table["비교 결과"] == "False"]
        self.true_table = self.fill_table[self.fill_table["비교 결과"] == "True"]

        self.floor_table = pd.DataFrame(self.floor_height_data.values(), self.floor_height_data.keys(), columns=["층고"])
        self.floor_table["방화 도어"] = self.floor_table.index.map(self.fire_door_data)
        for floor, data in self.hall_item.items():
            for col, val in data.items():
                self.floor_table.loc[floor, col] = val
        self.floor_table.rename_axis("층표기", inplace=True)
        self.floor_table.reset_index(inplace=True)
        self.floor_table.fillna("", inplace=True)
        # collec_data = {"All":self.all_table, "Fill":self.fill_table, "False":self.false_table, "True":self.true_table, "Floor":self.floor_table}
        collec_data = [self.all_table, self.fill_table, self.true_table, self.false_table, self.floor_table]
        # print("All", self.all_table)
        # print("Fill", self.fill_table)
        # print("False", self.false_table)
        # print("True", self.true_table)
        # print("Floor", self.floor_table)
        return collec_data

    def find_layout(self):
        document = {}
        chk_kind = {"H":"승강로 평면도", "M":"기계실 배치도", "E":"출입구 의장도", "S":"승강로 단면도"}
        for doc in self.acad.Documents:
            layout_projNo = doc.Name[:6]
            if layout_projNo == self.projNo:
                layout_kind = re.findall("(\w)\.dwg", doc.Name)[0]
                if layout_kind in list(chk_kind.keys()):
                    document.update({layout_kind:doc})
        open_layout_kind = list(document.keys())
        for open_chk in open_layout_kind:
            del chk_kind[open_chk]
        if len(chk_kind) > 0:
            non_layout = list(chk_kind.values())
            raise Exception("Not enough", non_layout)

        return document

    def get_entity(self, layout_kind):
        if layout_kind == "H" and "T" not in self.document.keys():
            layout = self.document[layout_kind]
            ent_blo_name = ["LAD-RAIL", "LAD-OPB", "LAD-CWT", "LAD-GOV", "DIM_ENT", "LAD-CP", "LAD-NOTE-FIXED-BEAM"]
            ent_group = dict.fromkeys(ent_blo_name)
            for entity in layout.ModelSpace:
                if entity.EntityName == "AcDbBlockReference" and entity.EffectiveName == "LAD-HOISTWAY-HP-SC":
                    ent_group.update({"hoistway_m": "CEMEN"})
                    entity.explode()
                elif entity.EntityName == "AcDbBlockReference" and entity.EffectiveName == "LAD-HOISTWAY-HP-SS":
                    ent_group.update({"hoistway_m": "ST"})
                    entity.explode()
                elif entity.EntityName == "AcDbBlockReference" and entity.EffectiveName in ["LAD-CAR-1SCO", "LAD-CAR-2SSO"]:
                    ent_group.update({"car_center": entity.InsertionPoint})
                    entity.explode()
                elif entity.EntityName == "AcDbBlockReference" and entity.EffectiveName == "LAD-CAR-1SCO-CP":
                    ent_group.update({"platform_cp": entity.InsertionPoint})
                elif entity.EntityName == 'AcDbBlockReference' and entity.Name == "LAD-FORM-A3-DETAIL":
                    ent_group.update({"spec_data": entity})

            for entity in layout.ModelSpace:
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

        elif layout_kind == "M" and self.machinroomType == "MR":
            layout = self.document[layout_kind]
            ent_blo_name = ["LAD-HOLE", "U107", "LAD-HBTN-HOLE", "LAD-CP", "LAD-ELD", "LAD-GOV", "LAD-TM", "LAD-HATCH",
                            "DIM_ENT", "hst_cent",
                            "hst_hole_cent", "door_cdnt"]
            ent_group = dict.fromkeys(ent_blo_name)
            for entity in layout.ModelSpace:
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

            for entity in layout.ModelSpace:
                if entity.EntityName == "AcDbRotatedDimension" and entity.TextOverride != "":
                    if ent_group["DIM_ENT"] == None:
                        ent_group["DIM_ENT"] = []
                        ent_group["DIM_ENT"].append(entity)
                    else:
                        ent_group["DIM_ENT"].append(entity)
                elif entity.EntityName == "AcDbText" and "기계실 출입문" in entity.TextString:
                    ent_group.update({"door_cdnt": entity.InsertionPoint})
                elif ent_group[
                    "LAD-CP"] == None and entity.EntityName == 'AcDbBlockReference' and "LAD-CP" in entity.EffectiveName:
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

        elif layout_kind == "M" and self.machinroomType == "MRL":
            layout = self.document[layout_kind]
            ent_group = {}
            for entity in layout.ModelSpace:
                if entity.EntityName == 'AcDbBlockReference':
                    if "LAD-HOISTWAY" in entity.EffectiveName:
                        ent_group.update({entity.EffectiveName: entity})
                    elif entity.EffectiveName == "LAD-CAR-TP-INV":
                        ent_group.update({entity.EffectiveName: entity})

        elif layout_kind == "S":
            layout = self.document[layout_kind]
            ent_group = {"Text": [], "Polyline": []}
            for entity in layout.ModelSpace:
                if entity.EntityName == 'AcDbBlockReference':
                    if entity.EffectiveName == "LAD-FORM-A3-SIMPLE":
                        palette_area = (238 * 388) * entity.XEffectiveScaleFactor
                        ent_group.update({"palette_area": palette_area})
                    elif entity.EffectiveName == "LAD-TABLE-FLOOR-HEIGHT":
                        ent_group.update({"floor_table_y_cdnt": entity.InsertionPoint[1]})
                        entity.Explode()
                    elif entity.EffectiveName == "LAD-TABLE-FIRE-DOOR":
                        ent_group.update({"fdoor_table_y_cdnt": entity.InsertionPoint[1]})
                        entity.Explode()
                    elif "LAD-HOISTWAY" in entity.EffectiveName:
                        ent_group.update({"hoistway_info": entity})

            for entity in layout.ModelSpace:
                if entity.EntityName == 'AcDbText':
                    ent_group["Text"].append(entity)
                elif entity.EntityName == 'AcDbPolyline':
                    ent_group["Polyline"].append(entity)

        elif layout_kind == "E":
            layout = self.document[layout_kind]
            ent_blo_name = ["LAD-TITLE", "LAD-DOOR-JAMB", "LAD-OPEN-HOLE", "LAD-HBTN", "LAD-OPEN-AC", "LAD-HALL-LANTERN", "LAD-EMCY-SWITCH","LAD-REMOTE-CP"]
            ent_group = dict.fromkeys(ent_blo_name)
            ent_and_cdnt = {}
            for entity in layout.ModelSpace:
                if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-OPEN-HOLE":  # x좌표가 중복될 수 있으므로 별도로 update함.
                    if ent_group["LAD-OPEN-HOLE"] == None:
                        ent_group["LAD-OPEN-HOLE"] = [entity]
                    else:
                        ent_group["LAD-OPEN-HOLE"].append(entity)
                else:
                    for ent_name in ent_blo_name:
                        if entity.EntityName == 'AcDbBlockReference' and ent_name in entity.EffectiveName:
                            if ent_name not in ent_and_cdnt.keys():
                                ent_and_cdnt.update({ent_name: {entity.InsertionPoint[0]: entity}})
                            else:
                                ent_and_cdnt[ent_name].update({entity.InsertionPoint[0]: entity})
                            break

            for ent_name, x_cdnt_and_ents in ent_and_cdnt.items():
                if len(x_cdnt_and_ents) == 1:
                    ent_group[ent_name] = list(x_cdnt_and_ents.values())
                elif len(x_cdnt_and_ents) > 1:
                    sort_ents = []
                    sort_cdnts = sorted(x_cdnt_and_ents.keys())
                    if ent_name == "LAD-OPEN-AC":
                        jamb_x_cdnt_list = sort_cdnts
                    for cdnt in sort_cdnts:
                        sort_ents.append(x_cdnt_and_ents[cdnt])
                    ent_group[ent_name] = sort_ents

            if ent_group["LAD-EMCY-SWITCH"] != None:
                firesw_ent = ent_group["LAD-EMCY-SWITCH"][0]
                firesw_x = firesw_ent.InsertionPoint[0]
                firesw_jamb_x = min(jamb_x_cdnt_list, key=lambda x: abs(x - firesw_x))
                firesw_jamb_ord = jamb_x_cdnt_list.index(firesw_jamb_x)
                if firesw_jamb_x < firesw_x:
                    firesw_pst = "RIGHT"
                elif firesw_jamb_x > firesw_x:
                    firesw_pst = "LEFT"
                ent_group.update({"FIRESW_X": firesw_x, "FIRESW_JAMB_ORD": firesw_jamb_ord, "FIRESW_PST": firesw_pst})

            if ent_group["LAD-OPEN-HOLE"] != None:
                ent_group.update({"HBTN_HOLE": [], "OTHER_HOLE": []})
                for cable_hole in ent_group["LAD-OPEN-HOLE"]:
                    for cable_hole_prt in cable_hole.GetDynamicBlockProperties():
                        if cable_hole_prt.propertyname == "@OFFSET-Y" and cable_hole_prt.value < 1400:  # hole 높이
                            ent_group["HBTN_HOLE"].append(cable_hole.InsertionPoint[0])
                            break
                        elif cable_hole_prt.propertyname == "@OFFSET-Y" and cable_hole_prt.value >= 1400:
                            ent_group["OTHER_HOLE"].append(cable_hole.InsertionPoint[0])
                            break
                del ent_group["LAD-OPEN-HOLE"]

            if ent_group["LAD-HALL-LANTERN"] != None:
                for lntn_ent in ent_group["LAD-HALL-LANTERN"]:
                    lntn_x = lntn_ent.InsertionPoint[0]
                    lntn_jamb_x = min(jamb_x_cdnt_list, key=lambda x: abs(x - lntn_x))
                    lntn_jamb_ord = jamb_x_cdnt_list.index(lntn_jamb_x)
                    if lntn_jamb_x < lntn_x:
                        lant_pst = "RIGHT"
                    elif lntn_jamb_x > lntn_x:
                        lant_pst = "LEFT"
                ent_group.updata({"LNTN" + str(lntn_jamb_ord): {"LNTN_X": lntn_x, "LNTN_PST": lant_pst}})

        if layout_kind == "H" or layout_kind == "MR":
            layout = self.document[layout_kind]
            layout.Activate
            layout.SendCommand('setxdata ')

        return ent_group

    def get_floor_plan_data(self):
        floor_plan_entity = self.get_entity("H")
        spec_ent = floor_plan_entity["spec_data"]
        spec_data = {}
        tag_name = {"@GOVERNOR": "EL_ECGV", "@CAR_SAFETY": "EL_ECSF", "@TM_TYPE": "EL_ETM"}  # 특성코드와 dic형태로 매칭해주는 것도 생각해볼
        trs_tag_name = {"@BALANCE": "EL_ECBA", "@NUMBER": "EL_ACD1", "@NO": "EL_ECN", "@V_SPEC": ["EL_AVOLT", "EL_ALI", "EL_AHZ"],"@DRIVE_TYPE": "EL_ADRV", "@DRIVE": "EL_ATYP",
                        "@SPEED": "EL_ASPD", "@CAPA": ["EL_AMAN", "EL_ACAPA"], "@USE": "EL_AUSE", "@DOOR_DRIVE": "EL_AOPEN","@MOTOR_CAPA": "EL_ETMM",
                        "@ROPE_SPEC": ["EL_ERPD", "EL_ERPW", "EL_ERPR"], "@DOOR_SIZE": ["EL_ECJJ", "EL_ECHH"],"@CAR_SIZE": ["EL_ECCA", "EL_ECCB", "EL_ECCH"], "@CB_TYPE": "EL_DURTB"}  # 변환이 필요한 코드

        for spec_att in spec_ent.GetAttributes():
            if spec_att.TagString in tag_name.keys():
                spec_data.update({tag_name[spec_att.TagString]: spec_att.TextString})
            elif spec_att.TagString in trs_tag_name.keys():
                if spec_att.TagString == "@BALANCE":
                    att_value = re.findall("\d+", spec_att.TextString)[0]
                    spec_data.update({trs_tag_name[spec_att.TagString]: att_value})
                elif spec_att.TagString == "@NUMBER":
                    if spec_att.TextString[0] == "1":
                        spec_data.update({trs_tag_name[spec_att.TagString]: "D"})
                elif spec_att.TagString == "@NO":
                    att_value = re.findall("\d+", spec_att.TextString)
                    if len(att_value) == 1:
                        spec_data.update({trs_tag_name[spec_att.TagString]: att_value[0]})
                        spec_data.update({"EL_ABANK": "1C"})  # 승강로 당 카 수량
                    else:
                        spec_data.update({trs_tag_name[spec_att.TagString]: att_value})
                        spec_data.update({"EL_ABANK": str(len(att_value)) + "C"})
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
                        self.machinroomType = "MRL"
                    elif "LXVF" in spec_att.TextString or "WBLX" in spec_att.TextString:
                        att_value = "WBLX1_(LXVF)"
                        self.machinroomType = "MR"
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
                        att_value_idx = len(spec_att.TextString)
                    att_value_list = re.findall("\w+", spec_att.TextString[:att_value_idx])
                    if len(att_value_list) == 1:
                        use_cvt = {"인승": "PS", "장애": "HC", "비상": "EP", "병원": "BD", "전망": "OB", "누드": "ND", "인화": "PF",
                                   "화물": "FT", "자동차": "AM"}
                        cvt_value = att_value_list[0][:2]
                        use_value = use_cvt[cvt_value]
                    else:
                        use_cvt = {"비상": "E", "병원": "B", "전망": "O", "누드": "N", "인화": "F", "장애": "H"}
                        for be_data, af_data in use_cvt.items():
                            if be_data in att_value_list:
                                if "use_value" not in locals():
                                    use_value = af_data
                                elif "use_value" in locals():
                                    use_value = use_value + af_data
                    spec_data.update({trs_tag_name[spec_att.TagString]: use_value})
                elif spec_att.TagString == "@DOOR_DRIVE":
                    pdm_drive = ["1SCO", "2SSO", "2SL", "2SR", "2SLR", "3SSO", "3SL", "3SR", "3SLR", "2SCO", "2UP",
                                 "2UL","2UR", "2ULR", "3UP", "3UL", "3UR", "3ULR", "1SSO", "1SL", "1SR", "1SLR"]
                    for drive in pdm_drive:
                        if drive in spec_att.TextString:
                            drive_value = drive
                    if "drive_value" not in locals():
                        if "CENTER" in spec_att.TextString:
                            if re.search('\d', spec_att.TextString).group() == "1" or "2PANEL" in spec_att.TextString:
                                drive_value = "1SCO"
                            else:
                                drive_value = "Door open" + spec_att.TextString + "에 대한 정의가 핑요합니다."
                        elif "SIDE" in spec_att.TextString:
                            if re.search('\d', spec_att.TextString).group() == "2":
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
                        spec_data.update({under_name: str(att_value)})
                elif spec_att.TagString == "@CAR_SIZE":
                    for under_name in trs_tag_name["@CAR_SIZE"]:
                        if under_name == "EL_ECCA":
                            att_value = re.findall("CA\D+(\d+)", spec_att.TextString)[0]
                        elif under_name == "EL_ECCB":
                            att_value = re.findall("CB\D+(\d+)", spec_att.TextString)[0]
                        elif under_name == "EL_ECCH":
                            att_value = re.findall("CH\D+(\d+)", spec_att.TextString)[0]
                        spec_data.update({under_name: str(att_value)})

        floor_plan_data = {}
        if floor_plan_entity["hoistway_m"] != None:
            floor_plan_data.update({"EL_EHM": floor_plan_entity["hoistway_m"]})  # 승강로 재질

        dim_name = {"균형추레일간의거리(세로)": "EL_ECWBG", "균형추레일간의거리(가로)": "EL_ECWBG", "승강로내부(세로)": "EL_EHV","카바닥(세로)": "EL_ECBB","카내부(세로)": "EL_ECCB",
                    "카바닥(가로)": "EL_ECAA", "출입구유효폭(가로)": "EL_ECJJ", "카레일간의거리(가로)": "EL_ECBG", "승강로내부(가로)": "EL_EHH", "카내부(가로)": "EL_ECCA"}

        for dim_ent in floor_plan_entity["DIM_ENT"]:
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
                hoist_rgt_x = max(int(pt1[0]), int(pt2[0]))
                car_cen_h = abs(hoist_lft_x - int(floor_plan_entity["car_center"][0]))
                floor_plan_data.update({"EL_ECHOR": str(car_cen_h)})  # 카중심:가로
            elif size_name == "승강로내부(세로)":
                hoist_fro_y = min(int(pt1[1]), int(pt2[1]))
                hoist_rear_y = max(int(pt1[1]), int(pt2[1]))
                car_cen_v = abs(hoist_fro_y - int(floor_plan_entity["car_center"][1]))
                floor_plan_data.update({"EL_ECVER": str(car_cen_v)})  # 카중심:세로
            elif size_name == "카바닥(세로)":
                car_fro_y = min(int(pt1[1]), int(pt2[1]))
                car_rear_y = max(int(pt1[1]), int(pt2[1]))
                car_ee = int(floor_plan_entity["car_center"][1]) - car_fro_y
                floor_plan_data.update({"EL_ECEE": str(car_ee)})  # CAR;EE
                floor_plan_entity.update({"car_rear_y": car_rear_y})

            if size_name in dim_name.keys():
                floor_plan_data.update({dim_name[size_name]: str(size)})

        if len(floor_plan_entity["LAD-OPB"]) > 1:
            for opb_ent in floor_plan_entity["LAD-OPB"]:
                opb_x_cdnt = int(opb_ent.InsertionPoint[0])
                if opb_x_cdnt > hoist_rgt_x:  # 승강로 외부에 있다면 삭제
                    floor_plan_entity["LAD-OPB"].remove(opb_ent)

        if len(floor_plan_entity["LAD-OPB"]) > 1:
            dis_opb_cnt = 0
            dis_opb_ents = []
            for opb_ent in floor_plan_entity["LAD-OPB"]:
                if opb_ent.EffectiveName == "LAD-OPB-DISABLED":
                    dis_opb_ents.append(opb_ent)
                    if dis_opb_cnt == 0:
                        dis_opb_cnt = 1
                    else:
                        dis_opb_cnt = dis_opb_cnt + 1
            if dis_opb_ents:
                floor_plan_data.update({"EL_BHOPBQ": dis_opb_cnt})
                for dis_opb_ent in dis_opb_ents:
                    floor_plan_entity["LAD-OPB"].remove(dis_opb_ent)

        opbs = {}
        for opb_ent in floor_plan_entity["LAD-OPB"]:
            opb_rotate = opb_ent.Rotation
            opb_x_cdnt = opb_ent.InsertionPoint[0]
            opb_y_cdnt = opb_ent.InsertionPoint[1]
            if opb_y_cdnt < floor_plan_entity["car_center"][1]:  # 카중심보다 밑에 있을 떄
                if opb_rotate == 0:
                    if opb_x_cdnt < floor_plan_entity["car_center"][0]:
                        opb_pst = "R"  # RIGHT
                        opb_open = "CO"
                    elif opb_x_cdnt > floor_plan_entity["car_center"][0]:
                        opb_pst = "L"  # LEFT
                        opb_open = "SOR"
                elif opb_rotate > 0:
                    if opb_x_cdnt < floor_plan_entity["car_center"][0]:
                        opb_pst = "SR"  # RIGHT(측벽)
                        opb_open = "SOR"
                    elif opb_x_cdnt > floor_plan_entity["car_center"][0]:
                        opb_pst = "SL"  # LEFT(측벽)
                        opb_open = "CO"
            elif opb_y_cdnt == floor_plan_entity["car_center"][1]:
                if opb_x_cdnt < floor_plan_entity["car_center"][0]:
                    opb_pst = "SR"  # RIGHT(측벽)
                    opb_open = "CO"
                elif opb_x_cdnt > floor_plan_entity["car_center"][0]:
                    opb_pst = "SL"  # LEFT(측벽)
                    opb_open = "CO"
            if len(floor_plan_entity["LAD-OPB"]) == 1:
                floor_plan_data.update({"EL_EOPBP": opb_pst, "EL_BMOPBO": opb_open})  # OPB 위치, MAIN OPB OPEN
            elif len(floor_plan_entity["LAD-OPB"]) == 2:
                if len(opbs) < 2:
                    opbs.update({opb_ent.InsertionPoint[0]: [opb_pst, opb_open]})
                elif len(opbs) == 2:
                    opbs = sorted(opbs.items())
                    if "S" not in opbs[0][1][0] and "S" not in opbs[1][1][0]:
                        floor_plan_data.update({"EL_EOPBP": "A", "EL_BMOPBO": opbs[0][1][1], "EL_BSOPBO": opbs[1][1][1]})
                    elif "S" in opbs[0][1][0] and "S" in opbs[1][1][0]:
                        floor_plan_data.update({"EL_EOPBP": "SA", "EL_BMOPBO": opbs[0][1][1], "EL_BSOPBO": opbs[1][1][1]})
                    else:
                        floor_plan_data.update({"EL_EOPBP": "OPB 위치 확인이 필요합니다. MAIN OPB 위치 : " + opbs[0][1][0] + ", SUB OPB 위치 : " + opbs[1][1][0], "EL_BMOPBO": opbs[0][1][1],"EL_BSOPBO": opbs[1][1][1]})

        for cwt_ent in floor_plan_entity["LAD-CWT"]:
            if "ARROW" in cwt_ent.EffectiveName:
                for arr_prt in cwt_ent.GetDynamicBlockProperties():
                    if arr_prt.PropertyName == "@MIRROR" and arr_prt.value == 1:
                        tm_drt = "L"
                        break
                    elif arr_prt.PropertyName == "@MIRROR" and arr_prt.value == 0:
                        tm_drt = "R"
                        break
                floor_plan_data.update({"EL_ETMD": tm_drt})
            elif "BRAKET" not in cwt_ent.EffectiveName and "ARROW" not in cwt_ent.EffectiveName:
                cwt_x_cdnt = round(cwt_ent.InsertionPoint[0])
                cwt_y_cdnt = round(cwt_ent.InsertionPoint[1])
                rope_x = abs(round(floor_plan_entity["car_center"][0] - cwt_x_cdnt))
                rope_y = abs(round(floor_plan_entity["car_center"][1] - cwt_y_cdnt))
                floor_plan_data.update({"EL_EPPX": str(rope_x), "EL_EPPY": str(rope_y)})
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
                            weight_r = cwt_prt.value  # subweight 우측폭
                        if "weight_l" in locals() and "weight_r" in locals():
                            break
                    weight_w = int(weight_l + weight_r)
                    if cwt_x_cdnt < int(floor_plan_entity["car_center"][0]):
                        cwt_pst = "R/L"  # FRONT, REAR 구분 필요
                    elif cwt_x_cdnt > int(floor_plan_entity["car_center"][0]):
                        cwt_pst = "R/R"  # FRONT, REAR 구분 필요
                    floor_plan_data.update({"EL_ECWTP": cwt_pst})
                floor_plan_data.update({"EL_ECWW": str(weight_w)})  # CWT;WEIGHT폭

        for rail_ent in floor_plan_entity["LAD-RAIL"]:
            rail_x_cdnt = int(rail_ent.InsertionPoint[0])
            rail_y_cdnt = int(rail_ent.InsertionPoint[1])
            if rail_y_cdnt == int(floor_plan_entity["car_center"][1]):  # CAR RAIL
                car_rail_spec = re.findall("(\d+)K", rail_ent.EffectiveName)[0]
                floor_plan_data.update({"EL_ECRL": car_rail_spec})  # CAR;RAIL(K)
                if rail_x_cdnt < floor_plan_entity["car_center"][0]:  # right rail
                    for rail_prt in rail_ent.GetDynamicBlockProperties():
                        if rail_prt.PropertyName == "@P1 Y":
                            rail_size = abs(rail_prt.Value)
                            rail_h1 = int((rail_x_cdnt - rail_size - 3) - hoist_lft_x)
                            floor_plan_data.update({"EL_ERBH1": str(rail_h1)})
                            break
                elif rail_x_cdnt > floor_plan_entity["car_center"][0]:  # left rail
                    for rail_prt in rail_ent.GetDynamicBlockProperties():
                        if rail_prt.PropertyName == "@P1 Y":
                            rail_size = abs(rail_prt.Value)
                            rail_h2 = int(hoist_rgt_x - (rail_x_cdnt + rail_size + 3))
                            floor_plan_data.update({"EL_ERBH2": str(rail_h2)})
                            break
            else:  # CWT RAIL
                cwt_rail_spec = re.findall("(\d+)K", rail_ent.EffectiveName)[0]
                floor_plan_data.update({"EL_ECWRL": cwt_rail_spec})  # CWT;RAIL(K)
                if rail_y_cdnt > car_rear_y:  # 후락
                    rail_h3 = int(hoist_rear_y - rail_y_cdnt)
                else:
                    floor_plan_data.update({"EL_ERBAG": floor_plan_data["EL_EPPY"]})
                    if rail_x_cdnt < floor_plan_entity["car_center"][0]:  # 좌락
                        rail_h3 = int(rail_x_cdnt - hoist_lft_x)
                    elif rail_x_cdnt > floor_plan_entity["car_center"][0]:  # 우락
                        rail_h3 = int(hoist_rgt_x - rail_x_cdnt)
                floor_plan_data.update({"EL_ERBH3": str(rail_h3)})

        gov_ent = floor_plan_entity["LAD-GOV"][0]
        gov_x_cdnt = int(gov_ent.InsertionPoint[0])
        gov_y_cdnt = int(gov_ent.InsertionPoint[1])
        gov_y_gap = int(floor_plan_entity["car_center"][1]) - gov_y_cdnt
        car_cc = abs(gov_y_gap)
        floor_plan_data.update({"EL_ECCC": str(car_cc)})  # CAR;CC
        if gov_y_gap < 0:  # REAR
            if gov_x_cdnt < int(floor_plan_entity["car_center"][1]):
                floor_plan_data.update({"EL_ECGP": "R/L"})  # REAR & LEFT
            else:
                floor_plan_data.update({"EL_ECGP": "R/R"})  # REAR & RIGHT
        elif gov_y_gap > 0:  # FRONT
            if gov_x_cdnt < int(floor_plan_entity["car_center"][1]):
                floor_plan_data.update({"EL_ECGP": "F/L"})  # FRONT & LEFT
            else:
                floor_plan_data.update({"EL_ECGP": "F/R"})  # FROTN & RIGHT

        if floor_plan_entity["LAD-CP"] != None:
            cp_ent = floor_plan_entity["LAD-CP"][0]
            if cp_ent.EffectiveName == "LAD-CP" or cp_ent.EffectiveName == "LAD-CP-DOOR":  # 승강장 jamb 취부형 제어반
                for cp_prt in cp_ent.GetDynamicBlockProperties():
                    if cp_prt.propertyname == "@CASE-L":
                        case_l = cp_prt.value
                    elif cp_prt.propertyname == "@CASE-R":
                        case_r = cp_prt.value
                    if "case_l" in locals() and "case_r" in locals():
                        break
                sj = int(case_l + case_r)
                floor_plan_data.update({"EL_EMRLCJW": str(sj)})  # MRL;CP JAMB 폭(SJ)
                if cp_ent.EffectiveName == "LAD-CP":
                    cp_type = "J"
                elif cp_ent.EffectiveName == "LAD-CP-DOOR":
                    cp_type = "C"
                if cp_ent.InsertionPoint[0] < floor_plan_entity["platform_cp"][0]:
                    cp_pst = "L"
                else:
                    cp_pst = "R"
                floor_plan_data.update({"EL_EMRLCJ": cp_type + cp_pst})  # MRL;CP JAMB TYPE
            elif cp_ent.EffectiveName != "LAD-CP-AC":  # 승강로 제어반
                floor_plan_data.update({"EL_EMRLHSCP": "Y"})
                cp_x_cdnt = cp_ent.InsertionPoint[0]
                cp_y_cdnt = cp_ent.InsertionPoint[1]
                if cp_y_cdnt > floor_plan_entity["car_rear_y"]:
                    floor_plan_data.update({"승강로 제어반 위치": "REAR"})
                elif cp_y_cdnt > floor_plan_entity["car_center"][1]:
                    if cp_x_cdnt < floor_plan_entity["car_center"][0]:
                        floor_plan_data.update({"승강로 제어반 위치": "R/R"})
                    else:
                        floor_plan_data.update({"승강로 제어반 위치": "R/L"})
                elif cp_y_cdnt < floor_plan_entity["car_center"][1]:
                    if cp_x_cdnt < floor_plan_entity["car_center"][0]:
                        floor_plan_data.update({"승강로 제어반 위치": "F/R"})
                    else:
                        floor_plan_data.update({"승강로 제어반 위치": "F/L"})

            if floor_plan_entity["LAD-NOTE-FIXED-BEAM"] != None:
                fix_bm_ent = floor_plan_entity["LAD-NOTE-FIXED-BEAM"][0]
                floor_plan_data.updata({"EL_ESPB": "Y"})
                for fix_bm_att in fix_bm_ent.GetAttributes():
                    if fix_bm_att.TagString == "@BEAM-C":
                        if "100" in fix_bm_att.TextSting and "50" in fix_bm_att:
                            floor_plan_data.update({"EL_ESPBS": "E100"})
        chk_spec = {}
        for spec_code, val in spec_data.items():
            if spec_code in floor_plan_data.keys():
                if floor_plan_data[spec_code] != val:
                    chk_spec.update({spec_code: val})
            else:
                floor_plan_data.update({spec_code: val})

        return floor_plan_data, chk_spec

    def get_mr_data(self):
        mr_entity = self.get_entity("M")
        mr_data = {}
        for dim_ent in mr_entity["DIM_ENT"]:
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
                if int(mr_entity["hst_cent"][0]) < hoist_lft_x:
                    size_name = "EL_EHH_CHK"
                else:
                    size_name = "EL_EHH"
            elif size_name == "승강로내부(세로)":
                ver_dim_x = int(pt1[0])
                if ver_dim_x < int(mr_entity["hst_cent"][0]):
                    size_name = "EL_EHV"
                elif ver_dim_x > int(mr_entity["hst_hole_cent"][0]):
                    size_name = "EL_EHV_CHK"
                else:
                    gap_hoistway = abs(ver_dim_x - int(mr_entity["hst_cent"][0]))
                    gap_hoistway_hole = abs(ver_dim_x - int(mr_entity["hst_hole_cent"][0]))
                    if min(gap_hoistway, gap_hoistway_hole) == gap_hoistway:
                        size_name = "EL_EHV"
                    else:
                        size_name = "EL_EHV_CHK"
            mr_data.update({size_name: str(size)})

        if mr_data["EL_EHH"] == mr_data["EL_EHH_CHK"]:
            del mr_data["EL_EHH_CHK"]
        if mr_data["EL_EHV"] == mr_data["EL_EHV_CHK"]:
            del mr_data["EL_EHV_CHK"]

        if mr_entity["LAD-HATCH"] != None:
            mr_data.update({"EL_DMRCP": "Y"})

        if mr_entity["LAD-CP"] != None:
            cp_ent = mr_entity["LAD-CP"][0]
            for cp_att in cp_ent.GetAttributes():
                if cp_att.TagString == "@TEXT":
                    cp_cdnt = cp_att.TextAlignmentPoint

            if mr_entity["LAD-HBTN-HOLE"] != None:
                duct_ent = mr_entity["LAD-HBTN-HOLE"][0]
                hole_ent_x_gap = int(mr_entity["hst_cent"][0] - mr_entity["hst_hole_cent"][0])
                duct_x = int(duct_ent.InsertionPoint[0]) + hole_ent_x_gap
                cp_duct_x = abs(int(cp_cdnt[0]) - duct_x)
                cp_duct_y = abs(int(cp_cdnt[1] - duct_ent.InsertionPoint[1]))
                cp_to_duct = round(cp_duct_x + cp_duct_y, -3) + 1250
                mr_data.update({"EL_EDTA": str(cp_to_duct)})

            if mr_entity["LAD-TM"] != None:
                tm_ent = mr_entity["LAD-TM"][0]
                for tm_prt in tm_ent.GetDynamicBlockProperties():
                    if tm_prt.propertyname == "@PP":
                        mr_data.update({"EL_EPPY": str(round(tm_prt.value, 0))})
                    elif tm_prt.PropertyName == "@P2 Y":
                        tm_y_ang = int(tm_prt.value)
                    elif tm_prt.PropertyName == "@P2 X":
                        tm_x_ang = int(tm_prt.value)
                    elif "EL_EPPY" in mr_data.keys() and "tm_y_ang" in locals() and "tm_x_ang" in locals():
                        break

                if tm_y_ang == 0:
                    mr_data.update({"EL_EMFD": "90", "EL_EMCBD": "R"})
                elif tm_x_ang == 0:
                    mr_data.update({"EL_EMFD": "90", "EL_EMCBD": "SH"})
                elif tm_x_ang != 0 and tm_y_ang > 0:
                    mr_data.update({"EL_EMFD": "SR", "EL_EMCBD": "SR"})
                elif tm_x_ang != 0 and tm_y_ang < 0:
                    mr_data.update({"EL_EMFD": "SL", "EL_EMCBD": "SL"})

                cp_tm_x = abs(int(cp_cdnt[0] - tm_ent.InsertionPoint[0]))
                cp_tm_y = abs(int(cp_cdnt[1] - tm_ent.InsertionPoint[1]))
                cp_to_tm = round(cp_tm_x + cp_tm_y + 1500, -3) + 1000
                mr_data.update({"EL_EDTB": str(cp_to_tm)})

            if mr_entity["door_cdnt"] != None:
                cp_door_x = abs(int(cp_cdnt[0] - mr_entity["door_cdnt"][0]))
                cp_door_y = abs(int(cp_cdnt[1] - mr_entity["door_cdnt"][1]))
                cp_to_pwr = round(cp_door_x + cp_door_y + 1650, -3) + 1000
                mr_data.update({"EL_EDTC": str(cp_to_pwr)})

            if mr_entity["LAD-GOV"] != None:
                for gov_ent in mr_entity["LAD-GOV"]:
                    if gov_ent.EffectiveName[-1] == "H":  # GOV HOLE
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
                        gov_cc = abs(int(mr_entity["hst_cent"][1] - gov_ent.InsertionPoint[1]))
                        cp_to_gov = round(cp_gov_x + cp_gov_y + 150, -3) + 1000
                        mr_data.update({"EL_EDTE": str(cp_to_gov), "EL_ECCC": str(gov_cc)})
                    mr_data.update({gov_name: gov_spec})

            if mr_entity["LAD-ELD"] != None:
                eld_ent = mr_entity["LAD-ELD"][0]
                for eld_prt in eld_ent.GetDynamicBlockProperties():
                    if eld_prt.PropertyName == "@DIST-X":
                        eld_x_dst = eld_prt.value
                    elif eld_prt.PropertyName == "@DIST-Y":
                        eld_y_dst = eld_prt.value
                    elif eld_prt.PropertyName == "@MIRROR":
                        eld_drt = eld_prt.value
                    elif "eld_x" in locals() and "eld_y" in locals() and eld_drt in locals():
                        break

                eld_y = eld_ent.InsertionPoint[1] - eld_y_dst
                if eld_drt == 0:
                    eld_x = eld_ent.InsertionPoint[0] + eld_x_dst
                elif eld_drt == 1:
                    eld_x = eld_ent.InsertionPoint[0] - eld_x_dst

                cp_eld_x = abs(int(cp_cdnt[0] - eld_x))
                cp_eld_y = abs(int(cp_cdnt[1] - eld_y))
                cp_eld_dst = round(cp_eld_x + cp_eld_y + 1275, -3)
                cp_to_eld = max(4000, cp_eld_dst)
                mr_data.update({"EL_DELD": "Y", "EL_EDTG": str(cp_to_eld)})

            if mr_data["EL_ECGV"] == mr_data["EL_ECGV_CHK"]:
                del mr_data["EL_ECGV_CHK"]

        return mr_data

    def get_mrl_data(self):
        mrl_entity = self.get_entity("M")
        mrl_data, chk_data = {}, {}
        mrl_data.update({"EL_EMRLTMP": "TT"})
        for ent_name, m_ent in mrl_entity.items():
            if "LAD-HOISTWAY-TP-WB" in ent_name:
                if ent_name[-2:] == "SC":
                    mrl_data.update({"EL_EHM": "CEMEN"})
                    for hoist_prt in m_ent.GetDynamicBlockProperties():
                        if hoist_prt.PropertyName == "@HA-L":
                            hoist_l = round(hoist_prt.value)
                            mrl_data.update({"EL_ECHOR": str(hoist_l)})
                        elif hoist_prt.PropertyName == "@HA-R":
                            hoist_r = round(hoist_prt.value)
                        elif hoist_prt.PropertyName == "@OH":
                            mrl_data.update({"EL_EHO": str(round(hoist_prt.value))})
                            mrl_data.update({"EL_EHO": str(round(hoist_prt.value))})
                        elif hoist_prt.PropertyName == "@VISIBLE":
                            if hoist_prt.value == "HB-OX":
                                mrl_data.update({"EL_DHK": "SS400"})
                            elif hoist_prt.value == "HB-XO":
                                mrl_data.update({"EL_DHK": "ANCHOR"})
                    mrl_data.update({"EL_EHH": str(hoist_l + hoist_r)})
                elif ent_name[-2:] == "SS":
                    mrl_data.update({"EL_EHM": "STWL"})
                    for hoist_prt in m_ent.GetDynamicBlockProperties():
                        if hoist_prt.PropertyName == "@HA-L":
                            hoist_l = round(hoist_prt.value)
                            mrl_data.update({"EL_ECHOR": str(hoist_l)})
                        elif hoist_prt.PropertyName == "@HA-R":
                            hoist_r = round(hoist_prt.value)
                        elif hoist_prt.PropertyName == "@OH":
                            mrl_data.update({"EL_EHO": str(hoist_prt.value)})
                    mrl_data.update({"EL_EHH": str(hoist_l + hoist_r)})

            elif ent_name == "LAD-HOISTWAY-TP-P-SC":
                for hoist_p_prt in m_ent.GetDynamicBlockProperties():
                    if hoist_p_prt.PropertyName == "@HA-L":
                        chk_hoist_l = round(hoist_p_prt.value)
                        chk_data.update({"CHK_EL_ECHOR": str(chk_hoist_l)})
                    elif hoist_p_prt.PropertyName == "@HA-R":
                        chk_hoist_r = round(hoist_p_prt.value)
                    elif hoist_p_prt.PropertyName == "@HB-B":
                        hoist_b = round(hoist_p_prt.value)
                        mrl_data.update({"EL_ECVER": str(hoist_b)})
                    elif hoist_p_prt.PropertyName == "@HB-T":
                        hoist_t = round(hoist_p_prt.value)
                chk_data.update({"CHK_EL_EHH": str(chk_hoist_l + chk_hoist_r)})
                mrl_data.update({"EL_EHV": str(hoist_b + hoist_t)})

            elif ent_name == "LAD-CAR-TP-INV":
                for car_prt in m_ent.GetDynamicBlockProperties():
                    if car_prt.PropertyName == "@CAR-H":
                        mrl_data.update({"EL_ECCH": str(round(car_prt.value))})
                    elif car_prt.PropertyName == "@DOOR-H":
                        mrl_data.update({"EL_ECHH": str(round(car_prt.value))})
                    elif car_prt.PropertyName == "@JJ-L":
                        jj_l = round(car_prt.value)
                    elif car_prt.PropertyName == "@JJ-R":
                        jj_r = round(car_prt.value)
                mrl_data.update({"EL_ECJJ": str(jj_l + jj_r)})

        chk_spec = {}
        if len(chk_data):
            for chk_code, chk_val in chk_data.items():
                if mrl_data[chk_code[4:]] != chk_val:
                    chk_spec.update({chk_code[4:]: chk_val})

        return mrl_data, chk_spec

    def get_section_table_cdnt(self, table_name):
        palette_area = self.section_entity["palette_area"]
        if table_name == "floor":
            if "floor_table_y_cdnt" in self.section_entity.keys():
                floor_table_y_cdnt = self.section_entity["floor_table_y_cdnt"]
                for poly_ent in self.section_entity["Polyline"]:  # 층 테이블 좌표 구하기
                    if floor_table_y_cdnt in poly_ent.Coordinates:
                        fl_st_x_cdnt = poly_ent.Coordinates[0]
                        fl_ed_x_cdnt = poly_ent.Coordinates[2]
                        fl_st_y_cdnt = poly_ent.Coordinates[1]
                        fl_ed_y_cdnt = poly_ent.Coordinates[5]
                        table_cdnt = [fl_st_x_cdnt, fl_ed_x_cdnt, fl_st_y_cdnt, fl_ed_y_cdnt]

            elif "floor_table_y_cdnt" not in self.section_entity.keys():
                text_insert_cdnt = {}
                for text_ent in self.section_entity["Text"]:  # 층 테이블의 데이터 좌표 구하기
                    if text_ent.TextString == "FL / ST":
                        text_base_y_cdnt = text_ent.TextAlignmentPoint[1]
                    elif text_ent.TextString == "층":
                        text_insert_cdnt.update(
                            {text_ent.TextAlignmentPoint[1]: text_ent.TextAlignmentPoint[0]})  # text Y:X(Y값은 상이함)
                text_base_x_cdnt = text_insert_cdnt[text_base_y_cdnt]  # "FL/ST" text와 같은 Y좌표를 가진 "층" text의 X좌표 구하기

                for poly_ent in self.section_entity["Polyline"]:  # 데이터와 가까운 Line 좌표 구하기
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
            if "fdoor_table_y_cdnt" in self.section_entity.keys():
                fdoor_table_y_cdnt = self.section_entity["fdoor_table_y_cdnt"]
                for poly_ent in self.section_entity["Polyline"]:  # 방화도어 TABLE 좌표 구하기
                    if fdoor_table_y_cdnt in poly_ent.Coordinates:
                        fd_st_x_cdnt = poly_ent.Coordinates[2]
                        fd_ed_x_cdnt = poly_ent.Coordinates[0]
                        fd_st_y_cdnt = poly_ent.Coordinates[5]
                        fd_ed_y_cdnt = poly_ent.Coordinates[1]
                        table_cdnt = [fd_st_x_cdnt, fd_ed_x_cdnt, fd_st_y_cdnt, fd_ed_y_cdnt]


            elif "fdoor_table_y_cdnt" not in self.section_entity.keys():
                for text_ent in self.section_entity["Text"]:  # 방화도어 TABLE의 데이터 좌표 구하기
                    if text_ent.TextString == "방화도어 유무":
                        text_base_x_cdnt = text_ent.TextAlignmentPoint[0]
                        text_base_y_cdnt = text_ent.TextAlignmentPoint[1]

                for poly_ent in self.section_entity["Polyline"]:  # 데이터와 가까운 Line 좌표 구하기
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

    def get_fire_door_data(self):
        s_x_cdnt, e_x_cdnt, s_y_cdnt, e_y_cdnt = self.get_section_table_cdnt("fire_door")
        table_data = {}
        x_cdnt_list = []
        fire_door_floor = {}
        for text_ent in self.section_entity["Text"]:
            x_cdnt = text_ent.InsertionPoint[0]
            y_cdnt = text_ent.InsertionPoint[1]
            if x_cdnt > s_x_cdnt and x_cdnt < e_x_cdnt and y_cdnt < s_y_cdnt and y_cdnt > e_y_cdnt:
                table_data.update({text_ent.TextAlignmentPoint: text_ent.TextString})  # 좌표안에 있는 테이블에 있는 모든 TEXT get
                if text_ent.TextString == "층":  # 윗행(층)과 아래행(층고) 나누는 기준
                    floor_y_cdnt = text_ent.TextAlignmentPoint[1]
                elif "방화도어" in text_ent.TextString:  # 윗행(층)과 아래행(층고) 나누는 기준
                    fire_door_y_cdnt = text_ent.TextAlignmentPoint[1]
                else:
                    x_cdnt_list.append(text_ent.TextAlignmentPoint[0])

        x_cdnt_list = list(set(x_cdnt_list))  # 중복 좌표 삭제
        x_cdnt_list.sort()  # x좌표 순서대로 정리

        for x in x_cdnt_list:
            floor_text = table_data.get((x, floor_y_cdnt, 0.0))
            floor_mark_list = convert_special_str(floor_text).convert_str

            fire_door = table_data.get((x, fire_door_y_cdnt, 0.0))
            fire_door = fire_door.upper()
            if fire_door == "O" or fire_door == "YES":
                fire_door = re.sub("O|YES", "Y", fire_door)[0]
            elif fire_door == "X" or fire_door == "NO":
                fire_door = re.sub("X|NO", "N", fire_door)[0]

            for floor_mark in floor_mark_list:
                fire_door_floor.update({floor_mark: fire_door})

        return fire_door_floor

    def get_section_data(self):
        self.section_entity = self.get_entity("S")
        s_x_cdnt, e_x_cdnt, s_y_cdnt, e_y_cdnt = self.get_section_table_cdnt("floor")
        table_data, floor_height_data = {}, {}
        x_cdnt_list = []
        for text_ent in self.section_entity["Text"]:
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
            floor_text = table_data[(x, floor_y_cdnt, 0.0)]  # 층표기 구하기
            floor_mark_list = convert_special_str(floor_text).convert_str
            floor_height = table_data[(x, floor_hei_y_cdnt, 0.0)]  # 층고 구하기
            for floor_mark in floor_mark_list:
                floor_height_data.update({floor_mark: floor_height})

        section_data = {}
        fl_st_data = table_data[(flst_x_cdnt, floor_hei_y_cdnt, 0.0)]
        section_data.update({"EL_ATF": ",".join(floor_height_data.keys())})  # total 층표기
        section_data.update({"EL_AFQ": re.findall("(\d+)/", fl_st_data)[0]})  # 층수
        section_data.update({"EL_ASTQ": re.findall("/(\d+)", fl_st_data)[0]})  # 정지층수
        section_data.update({"EL_EFHB": list(floor_height_data.values())[0]})  # BOT 층고
        section_data.update({"EL_EFHT": list(floor_height_data.values())[-2]})  # TOP-1 층고
        section_data.update({"EL_EFHMAX": max(floor_height_data.values())})  # 최대 층고
        section_data.update({"EL_EFHMIN": min(floor_height_data.values())})  # 최소 층고

        if section_data["EL_AFQ"] == section_data["EL_ASTQ"]:
            section_data.update({"EL_AFF": section_data["EL_ATF"]})
            section_data.update({"EL_AFQ": section_data["EL_AFQ"]})

        if self.section_entity["hoistway_info"]:
            hstw_ent = self.section_entity["hoistway_info"]
            hstw_att_name = {"@OH": "EL_EHO", "@HH": "EL_EHTH", "@TH": "EL_EHTRH", "@PIT": "EL_EHP"}
            for hstw_att in hstw_ent.GetAttributes():
                if hstw_att.TagString in hstw_att_name.keys():
                    section_data.update({hstw_att_name[hstw_att.TagString]: str(hstw_att.TextString)})
                elif hstw_att.TagString == "@BRAKET":
                    bracket_q = re.findall("(\d+)EA", hstw_att.TextString)[0]
                    section_data.update({"EL_ERBQ": str(int(bracket_q) + 3)})

        for floor, height in floor_height_data.items():
            if floor.isdigit():
                if int(floor) < 10:
                    section_data.update({"EL_EFH0"+floor:height})
                elif int(floor) < 40:
                    section_data.update({"EL_EFH"+floor:height})
            elif floor in ["B1", "B2", "B3","B4", "B5"]:
                section_data.update({"EL_EFH"+floor:height})
        fire_door_data = self.get_fire_door_data()

        return section_data, floor_height_data, fire_door_data

    def get_hall_items(self, hall_odr, jamb_spec):
        jamb_hole_ent = self.hall_entity["LAD-OPEN-AC"][hall_odr]
        jamb_hole_x_cdnt = jamb_hole_ent.InsertionPoint[0]
        hole_dic = {"@EMSW-H": "LEFT", "@HBTN-H": "HBTN", "@HPI-H": "HPI", "@LTRN-H": "RIGHT"}
        box_hole = []
        for hole_att in jamb_hole_ent.GetDynamicBlockProperties():
            if hole_att.propertyname in hole_dic.keys():
                if int(hole_att.value) > 0:
                    box_hole.append(hole_dic[hole_att.propertyname])
                del hole_dic[hole_att.propertyname]
            if not len(hole_dic):
                break

        jamb_ent = self.hall_entity["LAD-DOOR-JAMB"][hall_odr]
        for jamb_prt in jamb_ent.GetDynamicBlockProperties():
            if jamb_prt.PropertyName == "@HH":
                hh = str(int(jamb_prt.value))
            elif jamb_prt.PropertyName == "@JJ-L":
                jj_l = int(jamb_prt.value)
            elif jamb_prt.PropertyName == "@JJ-R":
                jj_r = int(jamb_prt.value)
            if "hh" in locals() and "jj_l" in locals() and "jj_r" in locals():
                break
        jj = str(jj_l + jj_r)

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
                hpi = ""
                break

        btn_ent = self.hall_entity["LAD-HBTN"][hall_odr]
        btn_x_cdnt = btn_ent.InsertionPoint[0]
        if "SMALL" in btn_ent.EffectiveName.upper():
            btn_spec = "HPB"
        elif "LARGE" in btn_ent.EffectiveName.upper():
            btn_spec = "HIP"

        if "HBTN" in box_hole:
            btn_type = "BOX"
        elif btn_x_cdnt in self.hall_entity["HBTN_HOLE"]:
            btn_type = "BOXLESS"
        else:
            btn_type = "확인할 수 없습니다."

        floor_items = {"JAMB": jamb_type, "JAMB_ORD": app_jamb, "HH": hh, "JJ": jj, "홀버튼": btn_spec, "홀버튼_취부": btn_type, "HPI": hpi}

        if hpi == "Y":
            if "2" in jamb_spec:
                hpi_type = "JAMB 취부"
            else:
                if "HPI" in box_hole:
                    hpi_type = "BOX"
                elif "HPI" not in box_hole:
                    hpi_type = "BOXLESS"
            floor_items.update({"HPI_취부": hpi_type})

        if self.hall_entity["LAD-HALL-LANTERN"] != None:
            lntn_ord = "LNTN" + str(hall_odr)
            if self.hall_entity[lntn_ord]["LNTN_PST"] in box_hole:
                lantern = "BOX"
            elif self.hall_entity[lntn_ord]["LNTN_X"] in self.hall_entity["OTHER_HOLE"]:
                lantern = "BOXLESS"
            else:
                lantern = "홀랜턴 type을 확인할 수 없습니다."
            floor_items.update({"홀랜턴": lantern})

        if self.hall_entity["LAD-EMCY-SWITCH"] != None and self.hall_entity["FIRESW_JAMB_ORD"] == hall_odr:
            if self.hall_entity["FIRESW_PST"] in box_hole:
                firesw = "BOX"
            if self.hall_entity["FIRESW_X"] in self.hall_entity["OTHER_HOLE"]:
                firesw = "BOXLESS"
            else:
                firesw = "소방스위치 type을 확인할 수 없습니다."
            floor_items.update({"소방스위치": firesw})

        if hall_odr == len(self.hall_entity["LAD-DOOR-JAMB"]) - 1 and self.hall_entity["LAD-REMOTE-CP"] != None:  # 마지막 jamb(=최상층 jamb)일 떄
            remote_cp = "Y"
            floor_items.update({"분리형 보조제어반": remote_cp})

        return floor_items

    def get_hall_data(self):
        self.hall_entity = self.get_entity("E")
        hall_data = {}
        hall_ord = 0
        floor_data = list(self.floor_height_data.keys())
        hall_item = dict.fromkeys(floor_data, {"위치":"기타층"})
        hall_item.update({floor_data[-1]:{"위치":"최상층"}})
        for tit_ent in self.hall_entity["LAD-TITLE"]:
            for att in tit_ent.GetAttributes():
                if att.tagstring == "@TITLE-T":
                    jamb_spec = re.findall("\d+", att.textstring)[0]
                elif att.tagstring == "@TITLE-B":
                    app_floor_info = att.textstring
                    hall_items = self.get_hall_items(hall_ord, jamb_spec)
            if "소방스위치" in hall_items.keys():
                fireman_sw = hall_items["소방스위치"]
                if fireman_sw == "BOXLESS":
                    firesw_type = "BL"
                elif fireman_sw == "BOX":
                    firesw_type = "B"
                hall_data.update({"EL_CFRSW": firesw_type})
            if "분리형 보조제어반" in hall_items.keys():
                remote_dv = hall_items["분리형 보조제어반"]
                hall_data.update({"EL_EMRLHDRD": remote_dv})
            hall_ord = hall_ord + 1
            if "기준층" in app_floor_info:
                floor_info = re.sub("층|FL|\s+", "", app_floor_info)
                main_txt_idx = floor_info.index("기준")
                main_txt_split = [floor_info[:main_txt_idx]]
                main_txt_split.append(floor_info[main_txt_idx + 2:])
                for split_txt in main_txt_split:
                    if split_txt == "":
                        pass
                    else:
                        brk_split_txt = re.split("\(|\)", split_txt) # "(" 또는 ")"로 split
                        for del_brk_txt in brk_split_txt:
                            if del_brk_txt.isalnum() and not del_brk_txt.isalpha():
                                main_floor = del_brk_txt
                            elif del_brk_txt != "" and not del_brk_txt.isalpha():
                                main_app_floor = convert_special_str(del_brk_txt).convert_str
                if "F" in floor_data and "4" in main_app_floor:
                    f_idx = main_app_floor.index("4")
                    main_app_floor.insert(f_idx, "F")
                    main_app_floor.remove("4")
                if "main_app_floor" not in locals():
                    main_app_floor = list(main_floor)
                elif main_floor not in main_app_floor:
                    main_app_floor.append(main_floor)
                hall_data.update({"EL_EMF": str(main_floor)})
                hall_item.update({main_floor:{"위치":"기준층"}})
                for app_floor in main_app_floor:
                    set_floor = hall_item[app_floor]
                    set_floor.update(hall_items)
            elif "기타층" in app_floor_info:
                for floor, location in hall_item.items():
                    if "기타층" in location.values():
                        set_floor = hall_item[floor]
                        set_floor.update(hall_items)
            elif "최상층" in app_floor_info:
                set_floor = hall_item[floor_data[-1]]
                set_floor.update(hall_items)
            else:
                app_floor_info = re.sub("층|FL|\s+", "", app_floor_info)
                floor_list = convert_special_str(app_floor_info).convert_str
                for app_floor in floor_list:
                    set_floor = hall_item[app_floor]
                    set_floor.update(hall_items)

        chk_datas = {"HH":set(), "JJ":set()}
        jamb_type_list = {"JP201": "JP201U", "JP200": "JP200U", "JP110": "JP110", "JP100": "JP100", "JP50": "JP50"}
        for floor, floor_spec in hall_item.items():
            chk_datas["HH"].add(floor_spec["HH"])
            chk_datas["JJ"].add(floor_spec["JJ"])
            jamb_ord = floor_spec["JAMB_ORD"][-3].replace("P", "4")
            if floor_spec["JAMB"] in jamb_type_list:
                jamb_type = jamb_type_list[floor_spec["JAMB"]]
            else:
                jamb_type = floor_spec["JAMB"]
            if "EL_CJM" + jamb_ord not in hall_data.keys():
                hall_data.update({"EL_CJM" + jamb_ord: jamb_type})
                hall_data.update({"EL_CJM" + jamb_ord + "F":floor})
                hall_data.update({"EL_CJM" + jamb_ord + "Q":"1"})
            else:
                app_floor = hall_data["EL_CJM" + jamb_ord + "F"]
                app_Q = hall_data["EL_CJM" + jamb_ord + "Q"]
                hall_data.update({"EL_CJM" + jamb_ord + "F":app_floor+","+floor})
                hall_data.update({"EL_CJM" + jamb_ord + "Q":int(app_Q)+1})

            if floor in self.fire_door_data.keys(): # jamb ord는 같은데 방화도어 적용여부가 다르다면 jabm ord가 추가되도록 할 것.(작업 필요)
                hall_data.update({"EL_CJM" + jamb_ord+ "FR":self.fire_door_data[floor]})

            if floor_spec["HPI"] == "Y":
                hall_data.update({"EL_CHPI" + jamb_ord : floor_spec["HPI_취부"]})

            if floor == floor_data[0]:
                hall_data.update({"EL_C"+floor_spec["홀버튼"]+"B":floor_spec["홀버튼_취부"]})
                hall_data.update({"EL_C"+floor_spec["홀버튼"]+"BF":floor})
                hall_data.update({"EL_C"+floor_spec["홀버튼"]+"BQ":"1"})
            elif floor_spec["위치"] == "최상층":
                hall_data.update({"EL_C"+floor_spec["홀버튼"]+"T":floor_spec["홀버튼_취부"]})
                hall_data.update({"EL_C"+floor_spec["홀버튼"]+"TF":floor})
                hall_data.update({"EL_C"+floor_spec["홀버튼"]+"TQ":"1"})
            elif floor_spec["위치"] == "기타층":
                if "EL_C"+floor_spec["홀버튼"]+"M1" not in hall_data.keys():
                    hall_data.update({"EL_C"+floor_spec["홀버튼"]+"M1":floor_spec["홀버튼_취부"]})
                    hall_data.update({"EL_C"+floor_spec["홀버튼"]+"M1F":floor})
                    hall_data.update({"EL_C"+floor_spec["홀버튼"]+"M1Q":"1"})
                else:
                    app_floor = hall_data["EL_C"+floor_spec["홀버튼"]+"M1F"]
                    hall_data.update({"EL_C"+floor_spec["홀버튼"]+"M1F":app_floor+","+floor})
                    app_Q = hall_data["EL_C"+floor_spec["홀버튼"]+"M1Q"]
                    hall_data.update({"EL_C"+floor_spec["홀버튼"]+"M1Q": int(app_Q)+1})
            elif floor_spec["위치"] == "기준층":
                hall_data.update({"EL_C" + floor_spec["홀버튼"] + "M2": floor_spec["홀버튼_취부"]})
                hall_data.update({"EL_C" + floor_spec["홀버튼"] + "M2F": floor})
                hall_data.update({"EL_C" + floor_spec["홀버튼"] + "M2Q": "1"})

        if len(chk_datas["HH"]) == 1:
            hall_data.update({"EL_ECHH": chk_datas["HH"].pop()})
        else:
            hall_data.update({"EL_ECHH": "층별 HH가 상이합니다."})

        if len(chk_datas["JJ"]) == 1:
            hall_data.update({"EL_ECJJ": chk_datas["JJ"].pop()})
        else:
            hall_data.update({"EL_ECJJ": "층별 JJ가 상이합니다."})

        return hall_data, hall_item
