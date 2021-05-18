import win32com.client
import re

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

def get_property():
    ext_list = ["@OH","@HH","@TH","@PIT"]
    trs_list = ["@BRAKET"]
    att_dict={}
    for entity in doc.ModelSpace:
        if entity.EntityName == 'AcDbBlockReference' and entity.EffectiveName == "LAD-HOISTWAY-HS-INV-AC":
            for att in entity.GetAttributes():
                tagstring = att.tagstring
                textstring = att.textstring
                if tagstring in ext_list:
                    att_dict.update({tagstring:textstring})
                elif tagstring in trs_list:
                    textstring = str(int(re.findall("(\d+)(?=EA)", textstring)[0])+4)
                    att_dict.update({tagstring:textstring})
            att_list=[att_dict]
    return att_list

print(get_property())