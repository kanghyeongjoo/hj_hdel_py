from selenium import webdriver
import time
import requests
from bs4 import BeautifulSoup

driver=webdriver.Ie("D:\Python\Practice(crawler)\IEDriverServer.exe")
driver.get("http://plm.hdel.co.kr/jsp/login/JsLogin.jsp")

driver.find_element_by_name("userid").send_keys("2020203")
driver.find_element_by_name("pwd").send_keys("S92462010*")
driver.find_element_by_xpath("//input[@src='/img/login_26.jpg']").click()


com_info = ps(page)