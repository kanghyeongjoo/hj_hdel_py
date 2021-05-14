
import pyautocad
from pyautocad import Autocad, APoint

cad = Autocad()
cad.prompt("hello")

for text in cad.iter_objects("AcDbText"):
    print(text.TextString, text.InsertionPoint)

dim_list = []
for dim_obj in cad.iter_objects("AcDbRotatedDimension"):
    dim_name_all = dim_obj.TextOverride
    dim_name_split = dim_name_all.split(" ")
    dim_name = dim_name_split[1].rstrip(":}<>")
    dim = int(dim_obj.Measurement)
    dim_dic = {dim_name:dim}
    dim_list.append(dim_dic)

print(dim_list)