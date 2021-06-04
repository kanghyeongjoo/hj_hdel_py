import pandas as pd
import win32com.client
import re


acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

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



# 함수 구분
car_size={}

hoistway_m = ent_group["hoistway_m"]
car_size.update({"승강로;재질":hoistway_m})

doc.SendCommand('setxdata ')
for dim_ent in ent_group["DIM_ENT"]:
        del_s = dim_ent.TextOverride.replace(" ", "")
        size_name = re.findall("[가-힣]+", del_s)[0]
        Xdata = dim_ent.GetXData("", "Type", "Data")
        pt1 = Xdata[1][-2]
        pt2 = Xdata[1][-1]
        if int(pt1[0]) == int(pt2[0]):
            size_name = size_name+"(세로)"
            car_size.update({size_name: int(dim_ent.Measurement)})
        elif int(pt1[1]) == int(pt2[1]):
            size_name = size_name + "(가로)"
            car_size.update({size_name: int(dim_ent.Measurement)})
        else:
            gaps = {}
            gaps.update({abs(int(pt1[0]) - int(pt2[0])) : "(가로)"})
            gaps.update({abs(int(pt1[1]) - int(pt2[1])) : "(세로)"})
            size_name = size_name + gaps[int(dim_ent.Measurement)]
            car_size.update({size_name:int(dim_ent.Measurement)})


        if size_name == "승강로내부(가로)":
            hoist_lft_x = min(int(pt1[0]), int(pt2[0]))
            car_cen_h = abs(hoist_lft_x - int(ent_group["car_center"][0]))
            car_size.update({"카중심:가로": car_cen_h})
        elif size_name == "승강로내부(세로)":
            hoist_fro_y = min(int(pt1[1]), int(pt2[1]))
            car_cen_v = abs(hoist_fro_y - int(ent_group["car_center"][1]))
            car_size.update({"카중심:세로": car_cen_v})
        elif size_name == "카바닥(세로)":
            car_fro_y = min(int(pt1[1]), int(pt2[1]))
            car_rear_y = max(int(pt1[1]), int(pt2[1]))
            car_ee = int(ent_group["car_center"][1]) - car_fro_y
            car_size.update({"CAR;EE": car_ee})
            ent_group.update({"car_rear_y":car_rear_y})


if len(ent_group["LAD-OPB"]) > 1:
    for opb_ent in ent_group["LAD-OPB"]:
        if opb_ent.EffectiveName == "LAD-OPB-DISABLED":
            car_size.update({"장애자_OPB": "Y"})
            ent_group["LAD-OPB"].remove(opb_ent)

opbs = {}
for opb_ent in ent_group["LAD-OPB"]:
    opb_rotate = opb_ent.Rotation
    opb_x_cdnt = opb_ent.InsertionPoint[0]
    opb_y_cdnt = opb_ent.InsertionPoint[1]
    if opb_y_cdnt < ent_group["car_center"][1]: #카중심보다 밑에 있을 떄
        if opb_rotate == 0:
            if opb_x_cdnt < ent_group["car_center"][0]:
                opb_pst = "RIGHT"
                opb_open = "CO"
            elif opb_x_cdnt > ent_group["car_center"][0]:
                opb_pst = "LEFT"
                opb_open = "SOR"
        elif opb_rotate > 0:
            if opb_x_cdnt < ent_group["car_center"][0]:
                opb_pst = "RIGHT(측벽)"
                opb_open = "SOR"
            elif opb_x_cdnt > ent_group["car_center"][0]:
                opb_pst = "LEFT(측벽)"
                opb_open = "CO"
    elif opb_y_cdnt == ent_group["car_center"][1]:
        if opb_x_cdnt < ent_group["car_center"][0]:
            opb_pst = "RIGHT(측벽)"
            opb_open = "CO"
        elif opb_x_cdnt > ent_group["car_center"][0]:
            opb_pst = "LEFT(측벽)"
            opb_open = "CO"
    if len(ent_group["LAD-OPB"]) == 1:
        car_size.update({"MAIN_OPB_위치": opb_pst, "MAIN_OPB_OPEN": opb_open})
    elif len(ent_group["LAD-OPB"]) == 2:
        if len(opbs) < 2:
            opbs.update({opb_ent.InsertionPoint[0]: [opb_pst, opb_open]})
        elif len(opbs) == 2:
            opbs = sorted(opbs.items())
            car_size.update({"MAIN_OPB_위치": opbs[0][1][0], "MAIN_OPB_OPEN": opbs[0][1][1]})
            car_size.update({"SUB_OPB_위치": opbs[1][1][0], "SUB_OPB_OPEN": opbs[1][1][1]})

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
            car_size.update({"cwt위치": "REAR"})
        elif abs(cwt_x_cdnt - int(ent_group["car_center"][0])) > abs(cwt_y_cdnt - int(ent_group["car_center"][1])): # 횡락
            for cwt_prt in cwt_ent.GetDynamicBlockProperties():
                if cwt_prt.propertyname == "@WIDTH-L":
                    weight_l = cwt_prt.value # subweight 좌측폭
                elif cwt_prt.propertyname == "@WIDTH-R":
                     weight_r = cwt_prt.value
            weight_w = int(weight_l + weight_r) # subweight 우측폭
            if cwt_x_cdnt < int(ent_group["car_center"][0]):
                cwt_pst = "LEFT"
            elif cwt_x_cdnt > int(ent_group["car_center"][0]):
                cwt_pst = "RIGHT"
            car_size.update({"cwt위치": cwt_pst})
        car_size.update({"WEIGHT폭":weight_w})


for rail_ent in ent_group["LAD-RAIL"]:
    rail_cdnt = rail_ent.InsertionPoint[1]
    if rail_cdnt == ent_group["car_center"][1]:
        car_rail_spec = re.findall("\d+K", rail_ent.EffectiveName)[0]
        car_size.update({"car_rail":car_rail_spec})
    else:
        cwt_rail_spec = re.findall("\d+K", rail_ent.EffectiveName)[0]
        car_size.update({"cwt_rail": cwt_rail_spec})

gov_ent = ent_group["LAD-GOV"][0]
gov_x_cdnt = int(gov_ent.InsertionPoint[0])
gov_y_cdnt = int(gov_ent.InsertionPoint[1])
gov_y_gap = int(ent_group["car_center"][1]) - gov_y_cdnt
car_cc = abs(gov_y_gap)
car_size.update({"CAR;CC": car_cc})
if gov_y_gap < 0:
    if gov_x_cdnt < int(ent_group["car_center"][1]):
        car_size.update({"car_gov_위치" : "R/L"}) # REAR & LEFT
    else:
        car_size.update({"car_gov_위치" : "R/R"}) # REAR & RIGHT
else:
    if gov_x_cdnt < int(ent_group["car_center"][1]):
        car_size.update({"car_gov_위치" : "F/L"}) # FRONT & LEFT
    else:
        car_size.update({"car_gov_위치" : "F/R"}) # FROTN & RIGHT


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
    car_size.update({"CP JAMB 폭(SJ)":sj})
    if cp_ent.EffectiveName == "LAD-CP":
        cp_type = "J"
    elif cp_ent.EffectiveName == "LAD-CP-DOOR":
        cp_type = "C"
    if cp_ent.InsertionPoint[0] < ent_group["platform_cp"][0]:
        cp_pst = "L"
    else:
        cp_pst = "R"
    car_size.update({"MRL;CP JAMB TYPE":cp_type + cp_pst})
elif cp_ent.EffectiveName != "LAD-CP-AC" :#승강로 제어반
    car_size.update({"승강로 CP":"Y"})
    cp_x_cdnt = cp_ent.InsertionPoint[0]
    cp_y_cdnt = cp_ent.InsertionPoint[1]
    if cp_y_cdnt > ent_group["car_rear_y"]:
        car_size.update({"제어반 위치":"REAR"})
    elif cp_y_cdnt > ent_group["car_center"][1]:
        if cp_x_cdnt < ent_group["car_center"][0]:
            car_size.update({"제어반 위치":"R/R"})
        else:
            car_size.update({"제어반 위치": "R/L"})
    elif cp_y_cdnt < ent_group["car_center"][1]:
        if cp_x_cdnt < ent_group["car_center"][0]:
            car_size.update({"제어반 위치":"F/R"})
        else:
            car_size.update({"제어반 위치": "F/L"})


print(car_size)

df_el = pd.DataFrame([car_size])
print(df_el)









