import re

tex = "AC 10kW"

zz = re.findall("(\d+[.]?\d+?)kW", tex)

print(zz)