import requests
from bs4 import BeautifulSoup

def find_ouid(proj_no):
    ouid_url = "http://plm.hdel.co.kr/jsp/help/ouidList.jsp?md%24number={}".format(proj_no)
    res = requests.get(ouid_url)
    ouid_address = BeautifulSoup(res.content, "html.parser")
    get_ouid = ouid_address.find("form").next_sibling.strip()#get spec을 진행할 때 sepc url을 여기서 return하는 것도 고려해볼 것.
    return get_ouid

ouid = find_ouid("185258L01")
print(ouid)

project_spec_url = "http://plm.hdel.co.kr/jsp/plmetc/elvInfo/elvinfomation.jsp?cOuid=860c9bb8&iOuid={}".format(ouid)
print(project_spec_url)