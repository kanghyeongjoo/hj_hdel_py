import win32com.client
import math
import string
import fnmatch

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument
#doc.Utility.Prompt("aa\n")

def get_PIT_dim():
    for entity in doc.ModelSpace:
        if entity.EntityName == "AcDbRotatedDimension":
            dim_name_all = entity.TextOverride
            dim_name = dim_name_all.strip(string.punctuation2)
            if dim_name == "PIT":
                dim = int(entity.Measurement)
                PIT_dim = {dim_name:dim}
    return PIT_dim

def get_railbrkt_itv():
    dim_XY_list = []
    dim_Y_list = []
    for entity in doc.ModelSpace:
        if entity.EntityName == "AcDbRotatedDimension":
            dim_name = entity.TextOverride
            if dim_name == "":
                dim = int(entity.Measurement)
                dim_xcdnt = entity.TextPosition[0]
                dim_ycdnt = entity.TextPosition[1]
                dim_xycdnt = {dim_xcdnt:dim_ycdnt}
                dim_XY_list.append(dim_xycdnt)
    return dim_XY_list

print(get_PIT_dim())
print(get_railbrkt_itv())

