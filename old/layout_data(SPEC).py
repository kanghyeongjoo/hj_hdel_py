import win32com.client
import math
import string

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

dim_list = []
for entity in doc.ModelSpace:
    if entity.EntityName == "AcDbRotatedDimension":
        dim_name_all = entity.TextOverride
        dim_name = dim_name_all.strip(string.punctuation2)
        dim = int(entity.Measurement)
        dim_dic = {dim_name:dim}
        dim_list.append(dim_dic)

#att_list = []
#for entity in doc.ModelSpace:
 #   if entity.EntityName == 'AcDbBlockReference' and entity.Name == "LAD-FORM-A3-DETAIL":
  #      for att in entity.GetAttributes():
   #         tag = {att.tagstring:att.textstring}
    #        att_list.append(tag)

print(dim_list)
#print(att_list)
