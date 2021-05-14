import win32com.client
import math
import string
import fnmatch
import pandas as pd
import re

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

ext_list=["@DOOR_DRIVE","@GOVERNOR","@CAR_SAFETY","@TM_TYPE"] #특성코드와 dic형태로 매칭해주는 것도 생각해볼 것

trs_list=["@BALANCE","@NO","@V_SPEC","@DRIVE_TYPE","@DRIVE","@SPEED","@CAPA","@USE","@CAR_NO"
         "@MOTOR_CAPA","@CB_TYPE","@ROPE_SPEC","@DOOR_SIZE","@CAR_SIZE"] # 변환이 필요한 코드

def get_property():
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
        textstring = textstring.upper()
        trs_textstring = re.findall("\d\d\dV|\d\dHZ",textstring)
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
        trs_textstring = []
        be_use = re.sub("\s|\(|\)", textstring)
        spc_chr_iter = re.finditer("\w\W")
        if len(spc_chr_iter) == 0:
            trs_textstring.append(be_use[:2])
        elif len(spc_chr_iter) >= 1:
            spc_chr_idx = []
            for spc_chr in spc_chr_iter:
                spc_chr_idx.append(spc_chr.end()-1)
            while len(spc_chr_idx) > 0:
                af_use = textstring[:spc_chr_idx.pop(0)]
                trs_textstring.append


    else:
        trs_tagstring=["ttt"]
        trs_textstring=["sdsf"]

    return trs_tagstring, trs_textstring

el_spec = get_property()
#el_spec_df = pd.DataFrame(el_spec)

print(el_spec)



