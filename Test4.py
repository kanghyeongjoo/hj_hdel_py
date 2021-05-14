import re

tex = "AC 10kW"

zzzz = re.findall("(\d+[.]?\d+?)kW", tex)

print(zzzz)