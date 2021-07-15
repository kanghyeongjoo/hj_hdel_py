import re
import requests
from bs4 import BeautifulSoup
from selenium import webdriver

def get_ouid(proj_no):
    ouid = {}
    ouid_url = "http://plm.hdel.co.kr/jsp/help/projectouidList.jsp?md%24number={}".format(proj_no)
    ouid_data = requests.get(ouid_url)
    ouid_address = BeautifulSoup(ouid_data.content, "html.parser")
    ouid_get = ouid_address.findAll("br")
    ouid_get = ouid_address.findAll("br")
    elv, prdt, sales, HL_dwg = ouid_get
    ouid.update({"elv_info": elv.previous_sibling.strip(), "product_id": prdt.previousSibling.strip(),
                 "sales_info": sales.previousSibling.strip()})
    layout_no, layout_id = HL_dwg.previousSibling.strip().split("::")
    ouid.update({"layout_no": layout_no, "layout_id": layout_id})
    return ouid


def layout_download(userid, password, proj_no):
    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    options.add_argument('window-size=1200,1000')
    options.add_argument("disable-gpu")
    options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36")
    options.add_experimental_option("prefs", {'download.default_directory': r"C:\Users\Administrator\Downloads",
                                              'download.prompt_for_download': False,
                                              'download.directory_upgrade': True})
    driver = webdriver.Chrome("D:\Python\chromedriver.exe", options=options)  # 실행 파일이 없을 경우 자동으로 설치되도록할 것.
    driver.get("http://plm.hdel.co.kr/jsp/login/JsLogin.jsp")
    driver.find_element_by_name("userid").send_keys(userid)
    driver.find_element_by_name("pwd").send_keys(password)
    driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[3]/td[3]/table/tbody/tr[6]/td/input").click()
    layout_ouid = get_ouid(proj_no)["layout_id"]
    driver.get("http://plm.hdel.co.kr/UIGenerate.do?cmd=center&gbn=info&cOuid=860c9851&iOuid={}&tabOuid=FT".format(layout_ouid))
    layout_download_page = driver.page_source
    layout_download_source = BeautifulSoup(layout_download_page, "html.parser")

    div_section = layout_download_source.find("div", {"id": "ScrollBody"})
    layout_download_list = div_section.findAll("tr", {"class": "listRow"})
    layout_name = []
    for download_ord in range(len(layout_download_list)):
        download_xpath = '//*[@id={}]/td[2]/a[2]'.format(download_ord)
        down_script = driver.find_element_by_xpath(download_xpath)
        layout_name.append(down_script.text)
        down_script_source = down_script.get_attribute("href")
        down_param = re.findall("(construction.+?)',", down_script_source)[0]
        download_url = "http://plm.hdel.co.kr/FileTransfer.do?cmd=download&downloadparm={}".format(down_param)
        driver.get(download_url)

    return layout_name
