import requests
from bs4 import BeautifulSoup
import pyautocad
from pyautocad import Autocad, APoint

# get ouid
def get_ouid(proj_no):
    ouid_url = "http://plm.hdel.co.kr/jsp/help/ouidList.jsp?md%24number={}".format(proj_no)
    ouid_data = requests.get(ouid_url)
    ouid_address = BeautifulSoup(ouid_data.content, "html.parser")
    ouid_get = ouid_address.find("form").next_sibling.strip()
    return ouid_get #get spec을 진행할 때 sepc url을 여기서 return하는 것도 고려해볼 것.

# 코드별 값 확인 URL에서 현장 정보 가져오기
def get_spec(ouid):
    spec_url = "http://plm.hdel.co.kr/jsp/plmetc/elvInfo/elvinfomation.jsp?cOuid=860c9bb8&iOuid={}".format(ouid)
    spec_data = requests.get(spec_url)
    spec_address = BeautifulSoup(spec_data.content, "html.parser")
    spec_list = spec_address.find_all("tr", "01-cell")
    code_qty = len(spec_list)
    spec_get = []
    for ord_no in range(0, code_qty):
        spec = spec_list[ord_no]
        code = spec.find_all("td")[3].get_text()
        val = spec.find_all("td")[4].get_text()
        code_val = {code:val}
        spec_get.append(code_val)
    return spec_get

def get_proj_spec(proj_no):
    ouid = get_ouid(proj_no)
    proj_spec = get_spec(ouid)
    return proj_spec

project_spec = get_proj_spec("186975L01")
print(project_spec)