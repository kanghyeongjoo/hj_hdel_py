from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup as bs
from tkinter import *
import urllib.request
import csv
import pandas as pd
import numpy as np

import openpyxl
import pyautogui
import time
import os

check_list_final = []

def automation():
    wb = openpyxl.load_workbook('C://업무 자동화//AUTO_PROCESS_LIST.xlsx')
    ws = wb['PROJ_LIST']

    sheet2 = wb['개별입력']
    # sheet3 = wb['표준의장표'] # 추가 2020.04.16 h.s.yun
    pdmid = sheet2['B1'].value
    pdmpw = sheet2['B2'].value
    file_path = str(sheet2['B4'].value)
    file_name = str(sheet2['B5'].value)

    # calcu_part = input('일괄계산할 파트의 번호 입력후 ENTER\n1. 기계 ALL\n2. 전기 ALL\n3. C(DE)\n4. M(AB)\n5. F(CEILING)\n6. 1,3PART\n7. 2PART\n8. 3PART\n:')

    driver_location_path = "C://업무 자동화//IEDriverServer.exe"
    os.environ["webdriver.Ie.driver"] = driver_location_path
    driver = webdriver.Ie(driver_location_path)

    driver.get('http://plm.hdel.co.kr/jsp/login/JsLogin.jsp')
    driver.maximize_window()
    action = ActionChains(driver)
    driver.implicitly_wait(3)

    elem = driver.find_element_by_name("userid")
    elem.clear()
    elem.send_keys(pdmid)
    time.sleep(0.1)
    elem = driver.find_element_by_name("pwd")
    elem.click()
    time.sleep(0.1)
    elem.clear()
    elem.send_keys(pdmpw)
    elem.send_keys(Keys.ENTER)
    driver.implicitly_wait(3)
    time.sleep(2)

    def divide_list(l, n): 
        for i in range(0, len(l), n): 
            yield l[i:i + n] 

    def projectNo1(project_no):
        base_url = "http://plm.hdel.co.kr/jsp/help/ouidList.jsp?md%24number="
        url = base_url + project_no
        driver.get(url)
        soup = bs(driver.page_source,'html.parser')
        soup_str = str(soup)
        tokenList = soup_str.split("@",5)
        elev_info = tokenList[1].replace(" <br/>product$vf","").replace(" <br/>","")
        #print(elev_info)   
        
        return (elev_info)

    def projectNo(project_no):
        base_url = "http://plm.hdel.co.kr/jsp/help/ouidList.jsp?md%24number="
        url = base_url + str(project_no)
        driver.get(url)
        soup = bs(driver.page_source,'html.parser')
        soup_str = str(soup)
        tokenList = soup_str.split("@",5)
        if len(tokenList) == 2:
            product_vf = None
            dec_no = None
        elif len(tokenList) <= 1:
            elev_info = None
            product_vf = None
            dec_no = None   
        else:
            elev_info = tokenList[1].replace(" <br/>product$vf","").replace("<br/>","").strip()
            product_vf = tokenList[2].replace(" <br/>","").strip()
            a = int(product_vf, 16)
            dec_no = str(a)                 # 10진수 번호 생성

        return (elev_info,product_vf,dec_no)


    if Cpart_del.get() or Mpart_del.get() or Fpart_del.get() or ONE_part_del.get() or TWO_part_del.get() or THREE_part_del.get():
        for r in ws.rows:                
            row_index = r[0].row   # 행 인덱스
            project_no_temp = str(r[0].value)
            ouid = projectNo(project_no_temp)
            if ouid[1] == None:
                continue
            del_bom = "http://plm.hdel.co.kr/jsp/plmetc/part/deleteAllBom.jsp?iOuid=product$vf@"+str(ouid[1])+"&cOuid=860c9a2c"
            driver.get(del_bom)
            driver.implicitly_wait(3)
            bom_delete_Cpart = driver.find_element_by_name('b_c')
            bom_delete_Mpart = driver.find_element_by_name('b_m')
            bom_delete_Fpart = driver.find_element_by_name('b_f')
            bom_delete_1part = driver.find_element_by_name('b_1')
            bom_delete_2part = driver.find_element_by_name('b_2')
            bom_delete_3part = driver.find_element_by_name('b_3')
            bom_delete = driver.find_element_by_xpath('/html/body/form/table[1]/tbody/tr/td[2]/table/tbody/tr/td[3]/a')

            if Cpart_del.get():
                driver.execute_script("arguments[0].click();", bom_delete_Cpart)        
                time.sleep(0.1)
            else:
                pass

            if Mpart_del.get():
                driver.execute_script("arguments[0].click();", bom_delete_Mpart)        
                time.sleep(0.1)
            else:
                pass

            if Fpart_del.get():
                driver.execute_script("arguments[0].click();", bom_delete_Fpart)        
                time.sleep(0.1)
            else:
                pass

            if ONE_part_del.get():
                driver.execute_script("arguments[0].click();", bom_delete_1part)        
                time.sleep(0.1)
            else:
                pass

            if TWO_part_del.get():
                driver.execute_script("arguments[0].click();", bom_delete_2part)        
                time.sleep(0.1)
            else:
                pass

            if THREE_part_del.get():
                driver.execute_script("arguments[0].click();", bom_delete_3part)        
                time.sleep(0.1)
            else:
                pass

            driver.execute_script("arguments[0].click();", bom_delete)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/form/table[4]/thead/tr/td/span")))

    else:
        pass

# 2020.04.17 추가 : 일괄 승인 일괄 wip버전 생성

    if all_approval.get():
        project_list_for_approval_temp = []
        for r in ws.rows:
            project_no_temp = str(r[0].value)
            project_list_for_approval_temp.append(project_no_temp)

        project_list_for_approval = '`versionup$\n'.join(project_list_for_approval_temp) + '`versionup$'
        url_for_approval = 'http://plm.hdel.co.kr/jsp/plmetc/part/elv_change/elv_change_left.jsp?iOuid=aaa&cOuid=860c9bb8'
        driver.get(url_for_approval)
        driver.implicitly_wait(3)
        driver.refresh()
        text_area_for_approval = driver.find_element_by_id('cmtText')
        text_area_for_approval.send_keys(project_list_for_approval)
        driver.execute_script("javascript:searchProduct();")
        time.sleep(1)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        driver.implicitly_wait(3)
        time.sleep(1)
    else:
        pass

    # eo_prducer_check = []
    # for r in ws.rows:
    #     project_no_temp = str(r[0].value)    
    #     ouid = projectNo1(project_no_temp)
    #     code_value_check = "http://plm.hdel.co.kr/jsp/plmetc/elvInfo/elvinfomation.jsp?cOuid=860c9bb8&iOuid=elv_info$vf@" + ouid
    #     df_read_code = pd.read_html(code_value_check, header=0)
    #     df_read_code1 = df_read_code[0]
    #     df_read_code2 = df_read_code1.loc[df_read_code1['특성코드'] == 'MANAGER_E']
    #     MANAGER_E = df_read_code2['특성값'].values.tolist()
    #     df_read_code2 = df_read_code1.loc[df_read_code1['특성코드'] == 'fcUser']
    #     fcUser = df_read_code2['특성값'].values.tolist()            
    #     if MANAGER_E != fcUser:
    #         eo_prducer_check.append(project_no_temp)

    if all_make_wip.get():
        for r in ws.rows:
            project_no_temp = str(r[0].value)
            make_wip = "http://plm.hdel.co.kr/ElvMakeWip.do?elvNumber=" + project_no_temp + "&userId=" + str(pdmid)
            driver.get(make_wip)
            driver.implicitly_wait(1)
            time.sleep(0.1)
    else:
        pass


    if four_cycle_check.get():
        for r in ws.rows:
            time.sleep(0.5)
            row_index = r[0].row   # 행 인덱스

            project_no_temp = r[0].value

            ouid = projectNo1(project_no_temp)
            four_cycle_save = "http://plm.hdel.co.kr/UIGenerate.do?cmd=center&gbn=info&cOuid=860c9bb8&iOuid=elv_info$vf@" + ouid + "&udBtnId="
            # four_cycle_save_url = 'http://plm.hdel.co.kr/jsp/plmetc/elvInfo/cyclefour.jsp?cOuid=860c9bb8&iOuid=elv_info$vf@ <http://plm.hdel.co.kr/jsp/plmetc/elvInfo/cyclefour.jsp?cOuid=860c9bb8&iOuid=elv_info%24vf%4093a0fa68>' + ouid
            # driver.get(four_cycle_save_url)

            driver.get(four_cycle_save)
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/ul[2]/li[15]/a")))
            four_cycle_save1 = driver.find_element_by_xpath("/html/body/div[2]/ul[2]/li[15]/a") #4cycle저장
            # driver.execute_script("javascript:cycle4Save_ACTION('860c9bb8', 'elv_info$vf@'"+ ouid + ");")
            # driver.execute_script("arguments[0].click();", four_cycle_save1)   
            four_cycle_save1.send_keys(Keys.ENTER)
            WebDriverWait(driver, 5).until(EC.number_of_windows_to_be(2))
            print(driver.window_handles)
            WebDriverWait(driver, 30).until(EC.number_of_windows_to_be(1))
            print(driver.window_handles)

    else:
        pass


    if C_calcul_check.get() or M_calcul_check.get() or F_calcul_check.get() or ONE_calcul_check.get() or TWO_calcul_check.get() or THREE_calcul_check.get():
        for r in ws.rows:

            row_index = r[0].row   # 행 인덱스

            project_no_temp = r[0].value

            ouid = projectNo1(project_no_temp)

            four_cycle_save = "http://plm.hdel.co.kr/UIGenerate.do?cmd=center&gbn=info&cOuid=860c9bb8&iOuid=elv_info$vf@" + ouid + "&udBtnId="
            bom_calculation = "http://plm.hdel.co.kr/jsp/plmetc/elvInfo/bomCalculation.jsp?iOuid=elv_info$vf@" + ouid + "&cOuid=860c9bb8"

            driver.get(bom_calculation)
            driver.implicitly_wait(2)
            bom_calculation_Mpart = driver.find_element_by_xpath('//input[@name="b_m"]')
            bom_calculation_Cpart = driver.find_element_by_xpath('//input[@name="b_c"]')
            bom_calculation_Fpart = driver.find_element_by_xpath('//input[@name="b_f"]')
            bom_calculation_1part = driver.find_element_by_xpath('//input[@name="b_1"]')
            bom_calculation_2part = driver.find_element_by_xpath('//input[@name="b_2"]')
            bom_calculation_3part = driver.find_element_by_xpath('//input[@name="b_3"]')
            bom_calculation_execute = driver.find_element_by_xpath('//a[@id="bomCalcBtn"]')

            if C_calcul_check.get():
                driver.execute_script("arguments[0].click();", bom_calculation_Cpart)
                time.sleep(0.1)
            else:
                pass

            if M_calcul_check.get():
                driver.execute_script("arguments[0].click();", bom_calculation_Mpart)   
                time.sleep(0.1)
            else:
                pass

            if F_calcul_check.get():
                driver.execute_script("arguments[0].click();", bom_calculation_Fpart)        
                time.sleep(0.1)
            else:
                pass
            
            if ONE_calcul_check.get():
                driver.execute_script("arguments[0].click();", bom_calculation_1part)        
                time.sleep(0.1)
            else:
                pass   

            if TWO_calcul_check.get():
                driver.execute_script("arguments[0].click();", bom_calculation_2part)        
                time.sleep(0.1)
            else:
                pass

            if THREE_calcul_check.get():
                driver.execute_script("arguments[0].click();", bom_calculation_3part)        
                time.sleep(0.1)
            else:
                pass

            driver.execute_script("arguments[0].click();", bom_calculation_execute)
            WebDriverWait(driver, 90).until(EC.element_to_be_clickable((By.ID, "resultbtn")))
        print('BOM계산완료!')
    else:
        pass

        # driver.implicitly_wait(3)
        # def projectNo1(project_no):
        #     base_url = "http://plm.hdel.co.kr/jsp/help/ouidList.jsp?md%24number="
        #     url = base_url + project_no
        #     driver.get(url)
        #     soup = bs(driver.page_source,'html.parser')
        #     soup_str = str(soup)
        #     soup_str_len = len(soup_str)

        #     if soup_str_len > 20:
        #         tokenList = soup_str.split("@",5)
        #         elev_info = tokenList[1].replace(" <br/>product$vf","")
        #         product_vf = tokenList[2].replace(" <br/>","")
                
        #         a = int(product_vf, 16)
        #         dec_no = str(a)                 # 10진수 번호 생성
                        
        #         return [elev_info,product_vf,dec_no]

        #     else:
        #         print("끝!")

        # 엑셀파일 열기

        # 현재 Active Sheet 얻기

    if check_bom_excel_C.get() or check_bom_excel_M.get() or check_bom_excel_F.get() or check_bom_excel_1.get() or check_bom_excel_2.get() or check_bom_excel_3.get():
        
        for r in ws.rows:
            csv_code_value = []

            row_index = r[0].row   # 행 인덱스

            project_no_temp = str(r[0].value)

            ouid = projectNo(project_no_temp)
            if not ouid[1] == None:
                csv_first_row = []   
                four_cycle_save = "http://plm.hdel.co.kr/UIGenerate.do?cmd=center&gbn=info&cOuid=860c9bb8&iOuid=elv_info$vf@" + ouid[0] + "&udBtnId="
                unfoundItem  = 'http://plm.hdel.co.kr/jsp/plmetc/part/bomchecking/redirect_UnfoundItem.jsp?iOuid='+ ouid[2] +'&B_M=N&B_C=N&B_F=N&B_1=N&B_2=N&B_3=N&option=null'
                if check_bom_excel_C.get():
                    unfoundItem = unfoundItem.replace("B_C=N","B_C=Y")
                else:
                    pass
                if check_bom_excel_M.get():
                    unfoundItem = unfoundItem.replace("B_M=N","B_M=Y")
                else:
                    pass
                if check_bom_excel_F.get():
                    unfoundItem = unfoundItem.replace("B_F=N","B_F=Y")
                else:
                    pass                
                if check_bom_excel_1.get():
                    unfoundItem = unfoundItem.replace("B_1=N","B_1=Y")
                else:
                    pass
                if check_bom_excel_2.get():
                    unfoundItem = unfoundItem.replace("B_2=N","B_2=Y")
                else: 
                    pass
                if check_bom_excel_3.get():
                    unfoundItem = unfoundItem.replace("B_3=N","B_3=Y")
                else:
                    pass

                productSpecDiff = 'http://plm.hdel.co.kr/jsp/plmetc/product/productSpecDiff.jsp?iOuid=product$vf@'+ ouid[1] + '&cOuid=860c9a2c'
                checkingData = 'http://plm.hdel.co.kr/jsp/plmetc/product/checkingData.jsp?iOuid=product$vf@'+ ouid[1] +'&cOuid=860c9a2c'
                partDrawCompare = 'http://plm.hdel.co.kr/jsp/plmetc/part/bomchecking/partDrawCompare.jsp?iOuid='+ouid[2]
                redirect_duplicateItem = 'http://plm.hdel.co.kr/jsp/plmetc/part/bomchecking/redirect_duplicateItem.jsp?iOuid='+ouid[2]
                partConditionCheck = 'http://plm.hdel.co.kr/jsp/plmetc/part/bomchecking/partConditionCheck.jsp?iOuid=product$vf@'+ouid[1]
                code_value_check = "http://plm.hdel.co.kr/jsp/plmetc/elvInfo/elvinfomation.jsp?cOuid=860c9bb8&iOuid=elv_info$vf@" + ouid[0]

                # driver.get(code_value_check)
                # driver.implicitly_wait(6)
                # soup = bs(driver.page_source,'html.parser')
                # tbody = soup.find_all('td')
                # row = [tr.text for tr in tbody]

                df_read_code = pd.read_html(code_value_check, header=0)
                df_read_code1 = df_read_code[0]
                for code in sheet2.rows:
                    code_I_want = code[4].value
                    if not code_I_want == None:
                        df_read_code2 = df_read_code1.loc[df_read_code1['특성코드'] == code_I_want]
                        df_read_code_value = df_read_code2['특성값'].values.tolist()
                        df_read_code_name = df_read_code2['특성명'].values.tolist()
                        
                        # for design_name in sheet3.rows:    # 추가 2020.04.16 h.s.yun
                        #     if df_read_code_value == [design_name[0].value]:
                        #         df_read_code_value = [design_name[1].value]
                        #     else:
                        #         pass      

                        csv_first_row = csv_first_row + df_read_code_name
                        csv_code_value = csv_code_value + df_read_code_value
                    else:
                        pass

                driver.get(four_cycle_save)
                driver.implicitly_wait(6)
                soup = bs(driver.page_source,'html.parser')
                common_input_error_temp = soup.find_all('textarea')[2] # 공통입력오류
                common_input_error_str = str(common_input_error_temp.text)
                common_input_error1 = common_input_error_str.split("\r\n- ")
                common_input_error2 = common_input_error1[0:]
                common_input_error2[0] = common_input_error2[0].replace("- ","")
                common_input_error3 = list(divide_list(common_input_error2, 1))
                for i in common_input_error3:
                    if i != ['']:
                        common_input_error4 = [project_no_temp] + csv_code_value +['공통입력오류']+i
                        check_list_final.append(common_input_error4)
                    else:
                        pass

                duty_error_temp = soup.find_all('textarea')[3] # DUTY 비교사항
                duty_error_str = str(duty_error_temp.text)
                duty_error = duty_error_str.split("\r\n- ")
                duty_error2 = duty_error[0:]
                duty_error2[0] = duty_error2[0].replace("- ","")
                duty_error3 = list(divide_list(duty_error2, 1))
                for i in duty_error3:
                    if i != ['']:
                        duty_error4 = [project_no_temp] + csv_code_value +['DUTY비교사항']+i
                        check_list_final.append(duty_error4)
                    else:
                        pass

                meca_input_error_temp = soup.find_all('textarea')[4] # 기계입력오류
                meca_input_error_str = str(meca_input_error_temp.text)
                meca_input_error = meca_input_error_str.split("\r\n- ")
                meca_input_error2 = meca_input_error[0:]
                meca_input_error2[0] = meca_input_error2[0].replace("- ","")
                meca_input_error3 = list(divide_list(meca_input_error2, 1))
                for i in meca_input_error3:
                    if i != ['']:
                        meca_input_error4 = [project_no_temp] + csv_code_value +['기계입력오류']+i
                        check_list_final.append(meca_input_error4)
                    else:
                        pass

                elec_input_error_temp = soup.find_all('textarea')[5] # 전기입력오류
                elec_input_error_str = str(elec_input_error_temp.text)
                elec_input_error = elec_input_error_str.split("\r\n- ")
                elec_input_error2 = elec_input_error[0:]
                elec_input_error2[0] = elec_input_error2[0].replace("- ","")
                elec_input_error3 = list(divide_list(elec_input_error2, 1))
                for i in elec_input_error3:
                    if i != ['']:
                        elec_input_error4 = [project_no_temp] + csv_code_value +['전기입력오류']+i
                        check_list_final.append(elec_input_error4)
                    else:
                        pass

                design_error_temp = soup.find_all('textarea')[6] # 의장정합성
                design_error_str = str(design_error_temp.text)
                design_error = design_error_str.split("\r\n- ")
                design_error2 = design_error[0:]
                design_error2[0] = design_error2[0].replace("- ","")
                design_error3 = list(divide_list(design_error2, 1))
                for i in design_error3:
                    if i != ['']:
                        design_error4 = [project_no_temp] + csv_code_value +['의장정합성']+i
                        check_list_final.append(design_error4)
                    else:
                        pass

                auto_input_error_temp = soup.find_all('textarea')[7] # 자동입력오류
                auto_input_error_str = str(auto_input_error_temp.text)
                auto_input_error = auto_input_error_str.split("\r\n- ")
                auto_input_error2 = auto_input_error[0:]
                auto_input_error2[0] = auto_input_error2[0].replace("- ","")
                auto_input_error3 = list(divide_list(auto_input_error2, 1))
                for i in auto_input_error3:
                    if i != ['']:
                        auto_input_error4 = [project_no_temp] + csv_code_value +['자동입력오류']+i
                        check_list_final.append(auto_input_error4)
                    else:
                        pass

                car_remark_temp = soup.find_all('textarea')[8] # 카내의장리마크
                car_remark_str = str(car_remark_temp.text)
                car_remark = car_remark_str.split("\n")
                car_remark2 = car_remark[0:]
                car_remark3 = list(divide_list(car_remark2, 1))
                for i in car_remark3:
                    if i != ['']:
                        car_remark4 = [project_no_temp] + csv_code_value +['카내의장 REMARK']+i
                        check_list_final.append(car_remark4)
                    else:
                        pass

                hall_remark_temp = soup.find_all('textarea')[9] # 승강장의장리마크
                hall_remark_str = str(hall_remark_temp.text)
                hall_remark = hall_remark_str.split("\n")
                hall_remark2 = hall_remark[0:]
                hall_remark3 = list(divide_list(hall_remark2, 1))
                for i in hall_remark3:
                    if i != ['']:
                        hall_remark4 = [project_no_temp] + csv_code_value +['승강장의장 REMARK']+i
                        check_list_final.append(hall_remark4)
                    else:
                        pass


                opt_remark_temp = soup.find_all('textarea')[10] # 옵션리마크
                opt_remark_str = str(opt_remark_temp.text)
                opt_remark = opt_remark_str.split("\n")
                opt_remark2 = opt_remark[0:]
                opt_remark3 = list(divide_list(opt_remark2, 1))
                for i in opt_remark3:
                    if i != ['']:
                        opt_remark4 = [project_no_temp] + csv_code_value +['옵션 REMARK']+i
                        check_list_final.append(opt_remark4)
                    else:
                        pass

                layout_remark_temp = soup.find_all('textarea')[11] # layout리마크
                layout_remark_str = str(layout_remark_temp.text)
                layout_remark = layout_remark_str.split("\n")
                layout_remark2 = layout_remark[0:]
                layout_remark3 = list(divide_list(layout_remark2, 1))
                for i in layout_remark3:
                    if i != ['']:
                        layout_remark4 = [project_no_temp] + csv_code_value +['LAYOUT REMARK']+i
                        check_list_final.append(layout_remark4)
                    else:
                        pass

                driver.get(unfoundItem) # 미수배 항목
                driver.implicitly_wait(6)
                soup = bs(driver.page_source,'html.parser')
                tbody = soup.find_all('td')
                # for tr in tbody:
                row = [tr.text for tr in tbody]
                unfound_list = row[9:]
                unfound_list2 = list(divide_list(unfound_list, 2))
                for i in unfound_list2:
                    unfound_list3 = ' '.join(i)
                    unfound_list4 = [project_no_temp] + csv_code_value +['미수배 항목']+i
                    check_list_final.append(unfound_list4)
            
                # print(check_list_final)
                productSpecDiff_list_temp1 = []
                driver.get(productSpecDiff) #주석 사양 확인
                driver.implicitly_wait(6)
                soup = bs(driver.page_source,'html.parser')
                tbody = soup.find_all('td')
                # for tr in tbody:
                row = [tr.text for tr in tbody]
                productSpecDiff_list = row[4:]
                productSpecDiff_list2 = list(divide_list(productSpecDiff_list, 6))
                for i in productSpecDiff_list2:
                    productSpecDiff_list3 = ','.join(i)
                    for aaa in i:
                        productSpecDiff_list_temp = aaa.replace("\xa0", "0")
                        productSpecDiff_list_temp1.append(productSpecDiff_list_temp)
                    productSpecDiff_list4 = [project_no_temp] + csv_code_value +['주석 사양 확인']+ productSpecDiff_list_temp1
                    # print(productSpecDiff_list4)
                    check_list_final.append(productSpecDiff_list4)
            
                # print(check_list_final)
            
                driver.get(checkingData) #자동 입력 오류
                driver.implicitly_wait(6)
                soup = bs(driver.page_source,'html.parser')
                tbody = soup.find_all('td')
                # for tr in tbody:
                row = [tr.text for tr in tbody]
                checkingData_list = row[2:]
                checkingData_list2 = list(divide_list(checkingData_list, 1))
                for i in checkingData_list2:
                    checkingData_list3 = ' '.join(i)
                    checkingData_list4 = [project_no_temp] + csv_code_value +['자동입력오류']+i
                    check_list_final.append(checkingData_list4)

                driver.get(partDrawCompare) # 수배 부품별 정합성
                driver.implicitly_wait(6)
                soup = bs(driver.page_source,'html.parser')
                tbody = soup.find_all('td')
                # for tr in tbody:
                row = [tr.text for tr in tbody]
                partDrawCompare_list = row[2:]
                partDrawCompare_list2 = list(divide_list(partDrawCompare_list, 8))
                for i in partDrawCompare_list2:
                    partDrawCompare_list3 = ' '.join(i)
                    partDrawCompare_list4 = [project_no_temp] + csv_code_value +['수배부품별 정합성']+i
                    check_list_final.append(partDrawCompare_list4)
            
                driver.get(redirect_duplicateItem) #중복수배리스트
                driver.implicitly_wait(6)
                soup = bs(driver.page_source,'html.parser')
                tbody = soup.find_all('td')
                # for tr in tbody:
                row = [tr.text for tr in tbody]
                redirect_duplicateItem_list = row[7:]
                redirect_duplicateItem_list2 = list(divide_list(redirect_duplicateItem_list, 4))
                for i in redirect_duplicateItem_list2:
                    redirect_duplicateItem_list3 = ' '.join(i)
                    redirect_duplicateItem_list4 = [project_no_temp] + csv_code_value +['중복수배 BLOCK.NO 리스트']+i
                    check_list_final.append(redirect_duplicateItem_list4)

                # driver.get(partConditionCheck) #수배로직 역체크
                # driver.implicitly_wait(6)
                # soup = bs(driver.page_source,'html.parser')
                # tbody = soup.find_all('td')
                # # for tr in tbody:
                # row = [tr.text for tr in tbody]
                # partConditionCheck_list = row[2:]
                # partConditionCheck_list2 = list(divide_list(partConditionCheck_list, 7))
                # for i in partConditionCheck_list2:
                #     partConditionCheck_list3 = i #"'"+
                #     partConditionCheck_list4 = [project_no_temp] + csv_code_value +['수배로직 역체크']+partConditionCheck_list3
                #     check_list_final.append(partConditionCheck_list4)
            else:
                pass
    
        # print(check_list_final)

        now = datetime.now()
        now_date = now.strftime('%Y_%m_%d')
        now_time = now.strftime('%H%M')
        #ws.move_range("A:A", rows=1, translate=True)
        with open(file_path+file_name+now_date+".csv",'w',newline='', encoding='CP949',errors = 'backslashreplace') as f:
        # with open(file_path+file_name+now_date+"_"+now_time+".csv",'w',newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['공사번호']+ csv_first_row + ['항목','result1','result2'])
            try:
                writer.writerows(check_list_final)
            except UnicodeEncodeError:
                print('유니코드에러')
                pass
        wb.close()
        
    else:
        pass
    # pyautogui.alert(str(eo_prducer_check) + '는 eo전송시 주의 요망!')
    pyautogui.alert('완료')
root=Tk()

all_approval=IntVar()
all_make_wip=IntVar()
four_cycle_check=IntVar()
check_bom_excel_M=IntVar()
check_bom_excel_C=IntVar()
check_bom_excel_F=IntVar()
check_bom_excel_1=IntVar()
check_bom_excel_2=IntVar()
check_bom_excel_3=IntVar()
Cpart_del=IntVar()
Mpart_del=IntVar()
Fpart_del=IntVar()
ONE_part_del=IntVar()
TWO_part_del=IntVar()
THREE_part_del=IntVar()
C_calcul_check=IntVar()
M_calcul_check=IntVar()
F_calcul_check=IntVar()
ONE_calcul_check=IntVar()
TWO_calcul_check=IntVar()
THREE_calcul_check=IntVar()

AUTO_EXE=IntVar()
# la 실행항목 선택
Label(root, text="진행순서 : BOM 일괄삭제").grid(column=0, row=0, sticky=W)
Label(root, text=" →승인, wip생성, 4CYCLE 저장").grid(column=1, row=0, sticky=W)
Label(root, text=" → BOM 일괄계산 → 미품목,입력오류 엑셀출력").grid(column=2, row=0, sticky=W)
# Label(root, text="→ BOM 일괄계산)").grid(column=2, row=0, sticky=W)
Checkbutton(root,text="공사정보 일괄 승인",variable=all_approval).grid(column=0, row=1, sticky=W)
Checkbutton(root,text="공사정보 일괄 wip생성",variable=all_make_wip).grid(column=1, row=1, sticky=W)
Checkbutton(root,text="일괄 4CYLCE 저장",variable=four_cycle_check).grid(column=2, row=1, sticky=W)
Label(root, text="BOM 일괄삭제").grid(column=0, row=2, sticky=W)
Checkbutton(root,text="C(DE)",variable=Cpart_del).grid(column=0, row=3, sticky=W)
Checkbutton(root,text="M(AB)",variable=Mpart_del).grid(column=0, row=4, sticky=W)
Checkbutton(root,text="F(CEILING)",variable=Fpart_del).grid(column=0, row=5, sticky=W)
Checkbutton(root,text="1",variable=ONE_part_del).grid(column=0, row=6, sticky=W)
Checkbutton(root,text="2",variable=TWO_part_del).grid(column=0, row=7, sticky=W)
Checkbutton(root,text="3",variable=THREE_part_del).grid(column=0, row=8, sticky=W)
Label(root, text="BOM 일괄계산").grid(column=1, row=2,sticky=W)
Checkbutton(root,text="C(DE)",variable=C_calcul_check).grid(column=1, row=3, sticky=W)
Checkbutton(root,text="M(AB)",variable=M_calcul_check).grid(column=1, row=4, sticky=W)
Checkbutton(root,text="F(CEILING)",variable=F_calcul_check).grid(column=1, row=5, sticky=W)
Checkbutton(root,text="1",variable=ONE_calcul_check).grid(column=1, row=6, sticky=W)
Checkbutton(root,text="2",variable=TWO_calcul_check).grid(column=1, row=7, sticky=W)
Checkbutton(root,text="3",variable=THREE_calcul_check).grid(column=1, row=8, sticky=W)
Label(root, text="미품목, 입력오류 엑셀 출력").grid(column=2, row=2,sticky=W)
Checkbutton(root,text="C(DE)",variable=check_bom_excel_C).grid(column=2, row=3, sticky=W)
Checkbutton(root,text="M(AB)",variable=check_bom_excel_M).grid(column=2, row=4, sticky=W)
Checkbutton(root,text="F(CEILING)",variable=check_bom_excel_F).grid(column=2, row=5, sticky=W)
Checkbutton(root,text="1",variable=check_bom_excel_1).grid(column=2, row=6, sticky=W)
Checkbutton(root,text="2",variable=check_bom_excel_2).grid(column=2, row=7, sticky=W)
Checkbutton(root,text="3",variable=check_bom_excel_3).grid(column=2, row=8, sticky=W)

Button(root, text='실행', command= automation, width=15,height=2).grid(column=1, row=9, sticky=W)
Button(root, text='닫기', command= root.quit, width=15,height=2).grid(column=2, row=9, sticky=W)
root.mainloop()

# , width=10,height=5
# .grid(column=2, row=3, sticky=W, rowspan=2)

# finish_bom_list = ['TEST-251366','TEST-251367']
# first_checker = '윤현식'
# second_checker = '장동운'
# approver = '왕세궁'
# eco_jeonsong(finish_bom_list,first_checker,second_checker,approver)