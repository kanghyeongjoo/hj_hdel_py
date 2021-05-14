import re

tex = "AC 10kW"

zzz = re.findall("(\d+[.]?\d+?)kW", tex)

print(zzz)