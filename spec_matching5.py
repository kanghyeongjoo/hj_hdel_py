def transe_property(tagstring, textstring):

    if tagstring == "@BALANCE":
        trs_tagstring = [tagstring]
        trs_textstring = [textstring.rstrip("%")]
    elif tagstring == "@NO":
        trs_tagstring = []
        trs_textstring = re.findall("\d+", textstring)
        for car_no in range(1,len(trs_textstring)+1):
            trs_tagstring.append("@NO"+str(car_no))
    elif tagstring == "@V_SPEC":
        trs_tagstring=["@동력전원", "@조명전원", "@주파수"]
        textstring = textstring.upper().replace(" ", "")
        trs_textstring = re.findall("\d+(?=V)|\d+(?=HZ)",textstring)
    elif tagstring == "@DRIVE_TYPE":
        trs_tagstring = [tagstring]
        car_btn = re.findall("\d+", textstring)
        trs_textstring = [''.join([car_btn[0], "C", car_btn[1], "BC"])]
    elif tagstring == "@DRIVE":
        trs_tagstring = [tagstring]
        if "WBSS" in textstring:
            trs_textstring = ["WBSS2"]
        elif "LXVF" in textstring or "WBLX" in textstring:
            trs_textstring = ["WBLX1"]
    elif tagstring == "@SPEED":
        trs_tagstring = [tagstring]
        trs_textstring = re.findall("\d+", textstring)
    elif tagstring == "@CAPA":
        trs_tagstring = ["@PERSON", "@CAPA"]
        trs_textstring = re.findall("\d+", textstring)
    elif tagstring == "@USE":
        trs_tagstring = [tagstring]
        be_textstring = []
        be_use = re.sub("\s|\(|\)", "", textstring)
        spc_chr = re.findall("\w\W", be_use)
        while re.search("\W", be_use) != None:
            spc_st = re.search("\W", be_use).start()
            spc_ed = re.search("\W", be_use).end()
            be_textstring.append(be_use[:spc_st][:2])
            be_use = be_use.lstrip(be_use[:spc_ed])
        be_textstring.append(be_use[:2])
        if len(be_textstring) == 1:
            pdm_use_list = {"인승": "PS", "장애": "HC", "비상": "EP", "병원": "BD", "전망": "OB", "누드": "ND", "인화": "PF",
                            "화물": "FT", "자동차": "AM"}
            for layout_data, pdm_data in pdm_use_list.items():
                if layout_data in be_textstring:
                    trs_textstring = [pdm_data]
        else:
            pdm_use_list = {"비상": "E", "병원": "B", "전망": "O", "누드": "N", "인화": "F", "장애": "H"}
            text_list = []
            for layout_data, pdm_data in pdm_use_list.items():
                if layout_data in be_textstring:
                    text_list.append(pdm_data)
            trs_textstring=["".join(text_list)]
    elif tagstring == "@MOTOR_CAPA":
        trs_tagstring = [tagstring]
        trs_textstring = re.findall('(\d+\.?\d?)', textstring)
    elif tagstring == "@ROPE_SPEC":
        trs_tagstring = ["@ROPE_mm", "@ROPE_Q", "@ROPING"]
        textstring = textstring.replace(" ", "")
        trs_textstring = re.findall('ø(\d+)', textstring)
        trs_textstring.append(re.findall("X(\d+)", textstring)[0])
        trs_textstring.append(re.findall("\((\d+:\d+)\)", textstring)[0])
    elif tagstring == "@DOOR_SIZE":
        trs_tagstring = ["@DOOR_JJ", "@DOOR_HH"]
        textstring = textstring
        trs_textstring = re.findall("JJ\D+(\d+)", textstring)
        trs_textstring.append(re.findall("HH\D+(\d+)", textstring)[0])
    elif tagstring == "@CAR_SIZE":
        trs_tagstring = ["@CAR_CA", "@CAR_CB", "@CAR_CH"]
        trs_textstring = re.findall("CA\D+(\d+)", textstring)
        trs_textstring.append(re.findall("CB\D+(\d+)", textstring)[0])
        trs_textstring.append(re.findall("CH\D+(\d+)", textstring)[0])

    return trs_tagstring, trs_textstring

import win32com.client
import math
import string
import fnmatch
import pandas as pd
import re

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

def get_property():
    ext_list=["@DOOR_DRIVE","@GOVERNOR","@CAR_SAFETY","@TM_TYPE","@CB_TYPE"] #특성코드와 dic형태로 매칭해주는 것도 생각해볼 것
    trs_list=["@BALANCE","@NO","@V_SPEC","@DRIVE_TYPE","@DRIVE","@SPEED","@CAPA","@USE","@MOTOR_CAPA","@ROPE_SPEC","@DOOR_SIZE","@CAR_SIZE"] # 변환이 필요한 코드
    att_dict = {}
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.Name == "LAD-FORM-A3-DETAIL":
            for att in entity.GetAttributes():
                tagstring = att.tagstring
                textstring = att.textstring
                if tagstring in ext_list:
                    att_dict.update({tagstring:textstring})
                elif tagstring in trs_list:
                    trs_tagstring, trs_textstring = transe_property(tagstring,textstring) # return된 2개의 값을 받아야함
                    print(trs_tagstring, trs_textstring)
                    for odr_trs in range(0,len(trs_tagstring)):
                        att_dict.update({trs_tagstring[odr_trs]: trs_textstring[odr_trs]})
            att_list = [att_dict]
    return att_list


el_spec = get_property()
#el_spec_df = pd.DataFrame(el_spec)

print(el_spec)



