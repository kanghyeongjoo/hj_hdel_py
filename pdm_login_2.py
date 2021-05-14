# requests만으로 login 해보기
import requests
from bs4 import BeautifulSoup

header = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36", "Referer":"http://plm.hdel.co.kr/LogIn.do"}
url = "http://plm.hdel.co.kr/jsp/login/JsLogin.jsp"
