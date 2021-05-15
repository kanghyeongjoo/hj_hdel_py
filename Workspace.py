import win32com.client
import math
import string
import re

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

def get_layout_dim():
    dim_dic = {}
    for entity in doc.ModelSpace:
        if entity.EntityName == "AcDbRotatedDimension":
            dim_name = entity.TextOverride
            if dim_name != "":

                dim = int(entity.Measurement)
                # trs_layout_dim(dim_name, dim)
                dim_dic.update({dim_name:dim})
    dim_list=[dim_dic]
    return dim_list


def trs_dim_name(dimname):

    dim_name = " ".join(re.findall("[가-힣]+", dimname))

    return trs_dim_name

dim = get_layout_dim()
print(dim)