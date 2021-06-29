import re
from bs4 import BeautifulSoup
from requests_html import HTMLSession

user = {"userid": "2020203", "pwd": "S92462010*"}
url = "http://plm.hdel.co.kr/UIGenerate.do?cmd=center&gbn=info&cOuid=860c9851&iOuid=constructiondrawing$vf@9507d3ef&tabOuid=FT"
headers = {
    'Connection': "keep-alive",
    'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36",
    'Content-Type': "application/x-www-form-urlencoded",
    'Accept': "*/*",
    'Origin': "http://plm.hdel.co.kr",
    'Referer': "http://plm.hdel.co.kr/etc/JsCopyright.jsp",
    'Accept-Language': "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    'Cookie': "SaveStateCookie=undefined; GWP_RENDER_STYLE=%5Bobject%20Object%5D; Notice=done; .SmartPortalAuthentication=49D607A2F8D71A6A5372811D767B16218E3A6CE3658E9D9168763E79A3A04BC90215A78A6429E5F0755A56A84C8D7FC3F93E337F071194B968FB7E06D54189FDF98E391D61D770BE85AB56BFA9151F3A5CF891FB691710D928287AF2A9889A26CD8ABE79; MSP_TENANT_ID=HEL; MSP_LAST_TNTID_helco{}=HEL; MSP_PORTAL_ON=helco{}; GWP_COMPANY_ID=HEL; GWP_LANGUAGE_CODE=ko; lcid=1042; JSESSIONID=RNt0IAXet91jEKFvYgHk632mqalXdapcMsygcEIIQayqTFRZJIjaG9Lbvbc6S1eA.amV1c19kb21haW4vTVMx".format(user["userid"],user["userid"])
    }

with HTMLSession() as ss:
    download_page_data = ss.get("http://plm.hdel.co.kr/UIGenerate.do?cmd=center&gbn=info&cOuid=860c9851&iOuid=constructiondrawing$vf@9507d3ef&tabOuid=FT", headers=headers)
    print(download_page_data.text)
