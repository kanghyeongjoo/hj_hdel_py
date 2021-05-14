import win32com.client
import math

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

def getcp1no():
    i = 1000
    for entity in doc.ModelSpace:
        name = entity.EntityName
        if name == 'AcDbBlockReference':
            for attrib in entity.GetAttributes():
                cpno = attrib.tagstring
                if cpno == "호기번호":
                    no = int(attrib.textstring)
                    if no < i:
                        i = no
                        return i

def cptocp(cp1no):
    for entity in doc.ModelSpace:
        name = entity.EntityName
        if name == 'AcDbBlockReference':
            for attrib in entity.GetAttributes():
                cptag = attrib.tagstring
                if cptag == "호기번호":
                    cpno = int(attrib.textstring)
                    if cpno == cp1no:
                        x1, y1, z1 = map (int, (entity.insertionpoint))
                    elif cpno == cp1no + 1:
                        x2, y2, z2 = map(int, (entity.insertionpoint))

    x = abs(x1 - x2)
    y = abs(y1 - y2)
    cptocpdis = (math.ceil(( x + y)/1000)*1000)

    return cptocpdis

cptocpno = getcp1no()
print("제어반 번호 : ", cptocpno)
cptocpdistance = cptocp(cptocpno)
print("CP to CP : ", cptocpdistance)

