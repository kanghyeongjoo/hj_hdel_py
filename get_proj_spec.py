import requests
from bs4 import BeautifulSoup


def get_proj_spec(proj_spec_url):
    res = requests.get(proj_spec_url)
    proj_spec_address = BeautifulSoup(res.content, "html.parser")
    info_list = proj_spec_address.find_all("tr", "01-cell")
    proj_spec = []
    for info in info_list:
        code = info.find_all("td")[3].get_text()
        val = info.find_all("td")[4].get_text()
        code_val = {code:val}
        proj_spec.append(code_val)
    return proj_spec, len(proj_spec)

project_spec_url = "http://plm.hdel.co.kr/jsp/plmetc/elvInfo/elvinfomation.jsp?cOuid=860c9bb8&iOuid=elv_info$vf@947B4DDB"

project_spec = get_proj_spec(project_spec_url)
print(project_spec)

