import win32com.client
import math
import string
import fnmatch

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

for entity in doc.ModelSpace:
    if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-TABLE-FLOOR-HEIGHT":
        table_blo_y_cdnt = entity.InsertionPoint[1]
        entity.Explode()

print(table_blo_y_cdnt)

for entity in doc.ModelSpace:
    if entity.EntityName == 'AcDbPolyline' and entity.Coordinates[1] == table_blo_y_cdnt:
        start_x_cdnt = entity.Coordinates[0]
        end_x_cdnt = entity.Coordinates[2]
        start_y_cdnt = entity.Coordinates[1]
        end_y_cdnt = entity.Coordinates[5]
print(start_x_cdnt,end_x_cdnt,start_y_cdnt,end_y_cdnt)

for entity in doc.ModelSpace:
    if entity.EntityName == 'AcDbText' and entity.TextString == "층" and entity.InsertionPoint[0] > start_x_cdnt and \
            entity.InsertionPoint[0] < end_x_cdnt and entity.InsertionPoint[1] < start_y_cdnt and \
            entity.InsertionPoint[1] > end_y_cdnt:
        floor_y_pst = entity.InsertionPoint[1]

floor_list = []
for entity in doc.ModelSpace:
    if entity.EntityName == 'AcDbText' and entity.TextString != "층" and entity.TextString != "FL / ST" and entity.InsertionPoint[1] == floor_y_pst:
        floor_x_cdnt = entity.InsertionPoint[0]
        floor_name = entity.TextString
        floor_list.append(entity.TextString)

print(floor_list)