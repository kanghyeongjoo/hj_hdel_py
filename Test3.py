import win32com.client
import math
import string
import re

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument


for entity in doc.ModelSpace:
    if entity.EntityName == "AcDbRotatedDimension":
        dim_name_all = entity.TextOverride
        if dim_name_all != "":
            din = dim_name_all.replace(" ","")
            dind = re.findall("[가-힣]+\s?", dim_name_all)
            dindd = re.findall("[^\W+]", dim_name_all)
            dindddd = " ".join(re.findall("[가-힣]+", dim_name_all))
            dinddd = re.findall("^\w+", din)
            print(dindddd)

use = ["인화"]
uses = ["비상", "장애"]

print(re.sub(use, uses, "인화"))