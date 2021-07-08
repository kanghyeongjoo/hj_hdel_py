import requests
from selenium import webdriver

options = webdriver.ChromeOptions()
options.add_argument("headless")
options.add_argument('window-size=1200,1000')
options.add_argument("disable-gpu")
options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36") #Headless 탐지 막기(User-Agent)

driver = webdriver.Chrome("D:\Python\chromedriver.exe", options=options)
driver.get("http://plm.hdel.co.kr/jsp/login/JsLogin.jsp")

driver.find_element_by_name("userid").send_keys("2020203")
driver.find_element_by_name("pwd").send_keys("S92462010*")
driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[3]/td[3]/table/tbody/tr[6]/td/input").click()
driver.get("http://plm.hdel.co.kr/jsp/help/migrationInput.jsp")
zz = driver.get_cookies()
cookeis = zz[0]["value"]

headers = {
    'Connection': "keep-alive",
    'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36",
    'Content-Type': "application/x-www-form-urlencoded",
    'Accept': "*/*",
    'Origin': "http://plm.hdel.co.kr",
    'Referer': "http://plm.hdel.co.kr/etc/JsCopyright.jsp",
    'Accept-Language': "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    'Cookie': "SaveStateCookie=undefined; GWP_RENDER_STYLE=%5Bobject%20Object%5D; Notice=done; .SmartPortalAuthentication=49D607A2F8D71A6A5372811D767B16218E3A6CE3658E9D9168763E79A3A04BC90215A78A6429E5F0755A56A84C8D7FC3F93E337F071194B968FB7E06D54189FDF98E391D61D770BE85AB56BFA9151F3A5CF891FB691710D928287AF2A9889A26CD8ABE79; MSP_TENANT_ID=HEL; MSP_LAST_TNTID_helco2020203=HEL; MSP_PORTAL_ON=helco2020203; GWP_COMPANY_ID=HEL; GWP_LANGUAGE_CODE=ko; lcid=1042; JSESSIONID={}".format(cookeis)
    }


migration_input_data = {"json" :'[{"md$number":"TEST-302104","EL_AMAN":"2"},{"md$number":"TEST-302075","EL_AMAN":"2"}]'}


migraion_result = requests.post("http://plm.hdel.co.kr/jsp/help/migration.jsp", data=migration_input_data, headers=headers)
print(migraion_result.text)