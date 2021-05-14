# pyautocad 가져오기

import pyautocad

# AutoCAD instance 생성
cad = pyautocad.Autocad()

print(cad.doc.Name)
