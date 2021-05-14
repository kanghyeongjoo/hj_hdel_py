import pyautocad
from pyautocad import Autocad, APoint

acad = Autocad(create_if_not_exists=True, visible=True)
adoc = acad.app.Documents
opb_path ="D:\도면\전기설계 라이브러리\동적 BLOCK\OPB-DA21A(28000671)"

for doc in adoc: # 도면 파일명 get
    print(doc.Name)

adoc.open(opb_path) # 도면 open



#print(dim_list)