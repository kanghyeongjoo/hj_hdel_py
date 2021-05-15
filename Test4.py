import re

tex = "AC 10kW"

zzzz = re.findall("(\d+[.]?\d+?)kW", tex)

print(zzzz)

dd ="ø6 X 12 WIRE (2:1)"
dd = dd.replace(" ","")

print(re.findall(":",dd))


textstg = "3ø 4선 380V / 1ø 220V 60HZ"
trs = re.findall("(\d\d\d)V|(\d\d)HZ", textstg)
print(trs)

pdm_use = {"인승": "PS", "장애": "HC", "비상": "EP", "병원": "BD", "전망": "OB", "누드": "ND",
           "인화": "PF", "화물": "FT", "자동차": "AM"}
uses ={"비상":"E", "병원":"B", "전망":"O", "누드":"N", "인화":"F", "장애":"H"}# for 순서대로 조건문을 줘서 IN이면, OUT해라

du = ["인화"]

for ke, vl in pdm_use.items():
    if ke in du:
        print(vl)

# aa = "Dfdfd"&"dfdfd"
# print(aa)

du = ["인화","장애"]

pdm_use_list = {"비상": "E", "병원": "B", "전망": "O", "누드": "N", "인화": "F", "장애": "H"}
for layout_data, pdm_data in pdm_use_list.items():
    if layout_data in du:
        text_list = "".join(pdm_data)

print(text_list)
