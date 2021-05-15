import win32com.client
import math
import string
import fnmatch

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

def get_property():
    att_list = []
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.Name == "LAD-FORM-A3-DETAIL":
            for att in entity.GetAttributes():
                tag_v = att.textstring
                if tag_v != "":
                    tag = {att.tagstring:att.textstring}
                    att_list.append(tag)
    return att_list

def get_layout_dim():
    dim_list = []
    for entity in doc.ModelSpace:
        if entity.EntityName == "AcDbRotatedDimension":
            dim_name_all = entity.TextOverride
            dim_name = dim_name_all.strip(string.punctuation2)
            if dim_name != "":
                dim = int(entity.Measurement)
                dim_dic = {dim_name:dim}
                dim_list.append(dim_dic)
    return dim_list

def find_cp_pst():
    cpdoor_JJ_gap_li=[]
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and fnmatch.fnmatch(entity.EffectiveName, "LAD-CP*"):
            cpdoor_xcdnt = int(entity.InsertionPoint[0])

    for entity in doc.ModelSpace:
        if entity.EntityName == "AcDbRotatedDimension" and entity.TextOverride.strip(string.punctuation2) == "출입구 유효폭":
            JJ_xcdnt = (int(entity.TextPosition[0]))
            cpdoor_JJ_gap_li.append(cpdoor_xcdnt-JJ_xcdnt)

    for gap in cpdoor_JJ_gap_li:
        min_gap = 10000
        if min_gap > abs(gap):
            min_gap = abs(gap)
            if gap < 0:
                cp_door_pst = {"제어반 위치":"LEFT"}
            else:
                cp_door_pst = {"제어반 위치":"RIGHT"}

    return cp_door_pst


el_spec=get_property()
print(el_spec)
print(get_layout_dim())
print(find_cp_pst())
