import requests
import json
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import threading
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
import numpy as np
from bs4 import BeautifulSoup
import pandas as pd

from selenium.common.exceptions import TimeoutException

#建照瀏覽器
def init_browser():
    service=Service(executable_path=r"C:\\Users\\user\\geckodriver.exe")
    # 初始化 Chrome 瀏覽器
    browser = webdriver.Firefox(service=service)
    # 設定瀏覽器視窗最大化
    browser.set_window_size(640, 960)
    return browser

def parse_table(companyName,year,html,company_sex,company_Flow):
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table', id='t214sb01')

    if not table:
        return None
    data = []
    try:
        men=company_sex["男性比率%"].values[0]/100
        women=company_sex["女性比率%"].values[0]/100
    except:
        men=""
        women=""
    try:
        flow=company_Flow["員工流動率(%)"].values[0]/100
    except:
        flow=""
    data.append(companyName)
    data.append(year)
    current_section = None
    current_subsection = None
    for row in table.find_all('tr'):
        if 'tblHead' in row.get('class', []):
            section_title = row.get_text(strip=True)
            if '環境構面' in section_title:
                current_section = '一、環境構面'
            elif '社會構面' in section_title:
                current_section = '二、社會構面'
            elif '治理構面' in section_title:
                current_section = '三、治理構面'
            else:
                current_section = None
            current_subsection = None  # Reset subsection when a new section is found
        else:
            cells = row.find_all('td')
            if not cells:
                continue
            if 'rowspan' in cells[0].attrs:
                try:
                    value = cells[2].get_text(strip=True)
                except:
                    value=""
                data.append(value)
            else:
                if '女性董事席次及比率' == cells[0].get_text(strip=True):
                    try:
                        value = cells[2].get_text(strip=True)
                    except:
                        value=""
                    data.append(value)
                else:
                    try:
                        value = cells[1].get_text(strip=True)
                        if "\n" in value:
                            value=value.replace("\n","")
                        if "\t" in value:
                            value=value.replace("\t","")

                        if current_section:
                            if current_subsection:
                                data.append(value)
                            else:
                                data.append(value)
                    except IndexError:
                        value=""
                        data.append(value)
    data.append(men)
    data.append(women)
    data.append(flow)   
    return data

#得到個股的排放資訊
def get_emission(browser,company_name,company_code,year,company_sex,company_Flow):
    url = "https://mops.twse.com.tw/mops/web/t214sb01"
    payload={
        'co_id':company_code,
        'YEAR1':year
    }
    headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
    }
    response=requests.post(url,data=payload,headers=headers)
    html=response.content
    # 解析表格
    table = parse_table(company_name,year,html,company_sex,company_Flow)
    return table
    
#下載永續報告書
def get_pdf(year):

    employeeSex=pd.read_excel("C:\\Users\\user\\OneDrive\\桌面\\專題\\Sustainable_ESG\\"+str(year+1911)+"EmployeeSex.xlsx")
    employeeFlow=pd.read_excel("C:\\Users\\user\\OneDrive\\桌面\\專題\\Sustainable_ESG\\"+str(year+1911)+"EmployeeFlow.xlsx")
    emission={}
    data=pd.read_csv("C:\\Users\\user\\OneDrive\\桌面\\專題\\Sustainable_ESG\\112年Listed_info_emission_Mod.csv",encoding='utf-8-sig')

    for company_info in data["公司名稱"].values:
        company_name = company_info
        company_code = data[data["公司名稱"]==company_info]["股票代號"].values[0]
        company_Sex=employeeSex[employeeSex["名稱"]==company_name]
        company_Flow=employeeFlow[employeeFlow["名稱"]==company_name]
        try:
            emission[company_code]=get_emission(company_name,company_code,year,company_Sex,company_Flow)
        except TimeoutException:
            continue
        print(emission)
        print("=====================================")
        time.sleep(10)        
    columns=["公司名稱","年份","直接溫室氣體排放量(範疇一)(噸CO2e)","能源間接(範疇二)(噸CO2e)","其他間接(範疇三)(噸CO2e)","溫室氣體排放密集度(噸CO2e/熟料產量)","溫室氣體管理之策略、方法、目標","再生能源使用率","提升能源使用效率","使用再生物料政策","用水量(公噸)","用水密集度(公噸)","水資源管理或減量目標","有害廢棄物(公噸)","非有害廢棄物(公噸)","總重量(有害+非有害)(公噸)","廢棄物密集度(公噸)","廢棄物管理政策或減量目標","員工福利平均數(仟元/人)","員工薪資平均數(仟元/人)","非擔任主管職務之全時員工薪資平均數(仟元/人)","非擔任主管職務之全時員工薪資中位數(仟元/人)","管理職女性主管占比","職業災害人數","職業災害人數比率(職災人數/總人數)","董事會席次(席)","獨立董事席次(席)","女性董事比率","董事出席董事會出席率","董監事進修時數符合進修要點比率","公司年度召開法說會次數(次)","男性比例","女性比例","員工流動率"]
    df=pd.DataFrame.from_dict(emission,orient='index',columns=columns)
    df.to_csv(str(year)+"年Listed_info_emission"+".csv",encoding='utf-8-sig')

#得到透明足跡的汙染裁罰紀錄
def Get_footprint_violations_Link(corp_id):
    browser=init_browser()
    url=r"https://thaubing.gcaa.org.tw/envmap?qt-front_content=1#{%22latlng%22:[24.292754,120.653797],%22zoom%22:10,%22basemap%22:%22satellite%22,%22factory%22:{%22id%22:%22%22,%22address%22:%22%22,%22name%22:%22%22,%22enabled%22:1,%22type%22:%22All%22,%22poltype%22:%22All%22,%22fine%22:1,%22finehard%22:0,%22realtime%22:0,%22illegal%22:0,%22overhead%22:0},%22airquality%22:{%22enabled%22:1},%22airbox%22:{%22id%22:%22%22,%22enabled%22:1}}"
    browser.get(url)
       # 等待網頁加載完畢
    element = WebDriverWait(browser, 15).until(EC.visibility_of_element_located((By.XPATH, "//*[@id=\"views-exposed-form-corp-corp-list-block-1\"]/div/div/div")))
    # 找到公司統一編號的輸入框
    corp_id_input = browser.find_element(By.XPATH, "//*[@id=\"edit-corp-id\"]")
    # 清空輸入框中的值
    corp_id_input.clear()
    # 將要填入的公司統一編號輸入到輸入框中
    corp_id_input.send_keys(str(corp_id))
    search_button=browser.find_element(By.XPATH,"//*[@id=\"edit-submit-corp\"]")
    search_button.click()
    # 等待網頁加載完畢
    element = WebDriverWait(browser, 15).until(EC.visibility_of_element_located((By.XPATH, "//*[@id=\"block-views-corp-corp-list-block-1\"]/div/div/div[2]")))
    html = browser.page_source
    soup=BeautifulSoup(html, 'html.parser')
    pollution_recordLink=[]
    try:
        for link in soup.find_all('div', class_='views-field views-field-facility-name factory-name'):
            a_tag = link.find('a')  
            href = a_tag['href'] 
            corp_href="https://thaubing.gcaa.org.tw/"+href
            pollution_recordLink.append(corp_href)
    except:
        pollution_recordLink=[]
    print(pollution_recordLink)
    browser.quit()
    return pollution_recordLink
    
    
#爬取證期的開罰案件
def get_Securites_and_futures_violations(type):
    browser=init_browser()
    url="https://mops.twse.com.tw/mops/web/t168sb01"
    browser.get(url)
    if type =="上市":
        company_type_select = browser.find_element(By.XPATH, "//*[@id=\"TYPEK\"]/option[1]")
    else:
        company_type_select = browser.find_element(By.XPATH, "//*[@id=\"TYPEK\"]/option[2]")
    company_type_select.click()
    search_button = browser.find_element(By.XPATH, "/html/body/center/table/tbody/tr/td/div[4]/table/tbody/tr/td/div/table/tbody/tr/td[3]/div/div[3]/form/table/tbody/tr/td[4]/table/tbody/tr/td[2]/div/div/input")
    search_button.click()
    # 等待網頁加載完畢
    element = WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, "//*[@id=\"t168Form\"]/table")))
    html = browser.page_source
    soup=BeautifulSoup(html, 'html.parser')
    table=soup.find("table",{"class":"hasBorder"})
    events_violation=[]
    for row in table.find_all('tr'):
        violated_list=[]
        cells = row.find_all('td')
        if len(cells) == 0:
            continue
        violated_list.append(cells[0].text.strip())
        violated_list.append(cells[1].text.strip())
        violated_list.append(cells[2].text.strip())
        violated_list.append(cells[3].text.strip())
        violated_list.append(cells[4].text.strip())
        violated_list.append(cells[5].text.strip())
        violated_list.append("證期局")
        events_violation.append(violated_list)

    browser.quit()
    return events_violation


#下載上櫃資料
def get_otc(esg_company):
    browser=init_browser()
    browser.get("https://mops.twse.com.tw/mops/web/t51sb01")
    # 選擇上櫃公司
    company_type_select = browser.find_element(By.XPATH, "//*[@id=\"search\"]/table/tbody/tr/td/select[1]/option[2]")
    company_type_select.click()

    #選擇公司產業
    industry_select = browser.find_element(By.XPATH, "//*[@id=\"search\"]/table/tbody/tr/td/select[2]/option[1]")
    industry_select.click()
    # 點擊查詢按鈕
    search_button = browser.find_element(By.XPATH, "//*[@id=\"search_bar1\"]/div/input")
    search_button.click()

    # 等待網頁加載完畢
    element = WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.XPATH, "/html/body/center/table/tbody/tr/td/div[4]/table/tbody/tr/td/div/table/tbody/tr/td[3]/div/div[5]/div/table[2]")))
    # 在這裡解析網頁中的資料，你可以使用Beautiful Soup等工具來解析表格資料或其他資訊
    html = browser.page_source
    company_info = {}
    soup=BeautifulSoup(html, 'html.parser')
    browser.quit()

    #取用上市櫃公司被證期局開罰的資料
    otc_violations=get_Securites_and_futures_violations(browser,"上櫃")
    table=soup.find("table",{"style":"width:100%;"})
    #解析每個<tr>標籤
    for row in table.find_all('tr'):
        cells = row.find_all('td')
        if len(cells) == 0:
            continue  # 跳過表頭
        # 提取企業資訊
        company_name=cells[2].text.strip()
        # if company_name in esg_company.keys():
        #     otc_violated=[]
        #     if company_name in otc_violations.keys():
        #         otc_violated.append(cells[0].text.strip())
        #         otc_violated.append(company_name)
        #         otc_violated.append(cells[3].text.strip())
        #         otc_violated.append(cells[6].text.strip())
        #         otc_violated.append(cells[7].text.strip())
        #         otc_violated.append(otc_violations[company_name])
        #         company_info[company_name] = otc_violated
        #     else:
        #         otc_violated.append(cells[0].text.strip())
        #         otc_violated.append(company_name)
        #         otc_violated.append(cells[3].text.strip())
        #         otc_violated.append(cells[6].text.strip())
        #         otc_violated.append(cells[7].text.strip())
        #         otc_violated.append("")
        #         company_info[company_name] = otc_violated

    return company_info



#下載上市資料
def get_listed(esg_company):
    url="https://mops.twse.com.tw/mops/web/t51sb01"
    # 準備表單數據 (payload)
    payload = {
        'TYPEK': 'sii',  # 市場別：上市公司
        'code': '',  # 產業別：水泥工業，根據需要修改
        'step': '1',
        'firstin': '1',
    }
    headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
    }
    try:
        response=requests.post(url,data=payload,headers=headers)
        if response.status_code == 200:
            # 使用 BeautifulSoup 解析網頁內容
            soup = BeautifulSoup(response.content, 'html.parser')
            print(soup)
            # 找到目標表格
            table = soup.find("table", {"style": "width:100%;"})
            company_info = {}

            # 解析每個 <tr> 標籤
            for row in table.find_all('tr'):
                cells = row.find_all('td')
                if len(cells) == 0:
                    continue  # 跳過表頭
                if cells[2].text.strip() in esg_company:
                    # 提取企業資訊
                    company_code = cells[0].text.strip()
                    listed_info = []
                    listed_info.append(company_code)  # 資料項目1
                    listed_info.append(cells[2].text.strip())
                    listed_info.append(cells[3].text.strip())  # 資料項目2
                    listed_info.append(cells[6].text.strip())  # 資料項目3
                    listed_info.append(cells[7].text.strip())  # 資料項目4
                    listed_info.append(cells[8].text.strip())  # 資料項目5
                    listed_info.append(cells[16].text.strip())  # 資料項目6
                    company_info[company_code] = listed_info
            return company_info
        else:
            print(f"請求失敗，狀態碼：{response.status_code}")
            return None
    except TimeoutException as t:
        print(f"{t}")
        


#下載portfolio
def get_portfolio(code):
    headers={
         'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:129.0) Gecko/20100101 Firefox/129.0'
    }
    json_data=requests.get(f"https://www.cmoney.tw/api/cm/MobileService/ashx/GetDtnoData.ashx?action=getdtnodata&DtNo=59449513&ParamStr=AssignID%3D{code}%3BMTPeriod%3D0%3BDTMode%3D0%3BDTRange%3D1%3BDTOrder%3D1%3BMajorTable%3DM722%3B&FilterNo=0",headers=headers).json()
    return json_data


# 下載永續相關的etf
def get_esg_etf():
    url = "https://www.twse.com.tw/rwd/zh/esg-index-product/etf"
    html_content = requests.get(url).text
    # 使用Beautiful Soup解析HTML
    soup = BeautifulSoup(html_content, "html.parser")
    # 找到所有<tr>標籤
    rows = soup.find_all("tr")
    etf_portfolio={}
    # 迭代每個<tr>標籤，解析資訊
    for row in rows:
        # 找到<tr>標籤中的<td>標籤，並獲取相關資訊
        cells = row.find_all("td")
        if len(cells) == 0:
            continue  # 跳過表頭
        # 獲取相關資訊
        etf_code = cells[0].text.strip()
        etf_name = cells[1].text.strip()
        etf_index = cells[2].text.strip()
        etf_date = cells[3].text.strip()
        etf_type = cells[4].text.strip()
        etf_description = cells[5].text.strip()
        try:
            etf_portfolio[etf_code]={"etf_name":etf_name,"etf_index":etf_index,"etf_date":etf_date,"etf_type":etf_type,"etf_description":etf_description,"portfolio":get_portfolio(etf_code)}
        except requests.exceptions.ConnectionError as e:
            print(f"Error connecting to the server: {e}")
        time.sleep(5)
    return etf_portfolio


def BuildEmissionViolationLink(listedInfo):
    emission={}
    listedCompany=listedInfo["公司名稱"].values
    for company in listedCompany:
        companyTaxID=listedInfo[listedInfo["公司名稱"]==company]["統一編號"].values[0]
        companyCode=listedInfo[listedInfo["公司名稱"]==company]["Unnamed: 0"].values[0]
        violationData=Get_footprint_violations_Link(companyTaxID)
        if violationData ==[]:
            continue
        else:
            emission[companyCode]="、".join(violationData)
    df=pd.DataFrame.from_dict(emission,orient="index",columns=["連結"])
    df.to_csv("EmissionViolationLink.csv",encoding='utf-8-sig')

def Get_Company_Emission_Violation(links):
    headers={
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:129.0) Gecko/20100101 Firefox/129.0'
    }
    website=requests.get(links,headers=headers)
    soup=BeautifulSoup(website.text,'html.parser')
    

def Parse_EmissionTable(companyName, year, html, ESG,company_sex, company_Flow):
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('div', id='individual-table-box1').find('table')

    if not table:
        return None

    data = []
    try:
        men = company_sex["男性比率%"].values[0] / 100
        women = company_sex["女性比率%"].values[0] / 100
    except:
        men = ""
        women = ""

    try:
        flow = company_Flow["員工流動率(%)"].values[0] / 100
    except:
        flow = ""

    data.append(companyName)
    data.append(year)

    current_section = None
    current_subsection = None

    for row in table.find_all('tr', id=True):
        cells = row.find_all('td')
        if not cells:
            continue
        
        if 'rowspan' in cells[0].attrs:
            current_section = cells[0].get_text(strip=True)
        else:
            if len(cells) >= 2:
                sub_title = cells[1].get_text(strip=True)
                value = cells[2].get_text(strip=True) if len(cells) > 2 else ""
                data.append(value)
                
                # 獲取其他數據
                if len(cells) > 3:
                    data_boundary = cells[3].get_text(strip=True)
                    audit_firm = cells[4].get_text(strip=True)
                    audit_standard = cells[5].get_text(strip=True)
                    audit_scope = cells[6].get_text(strip=True)
                    data.extend([data_boundary, audit_firm, audit_standard, audit_scope])
                else:
                    data.extend([""] * 4)

    data.append(men)
    data.append(women)
    data.append(flow)
    data.append(ESG)

    return data

def Get_112_EmissionData():
    greenHouse=pd.read_csv("C:\\Users\\user\\Desktop\\專題\\Sustainable_ESG\\上市112溫室氣體.csv")
    waterManage=pd.read_csv("C:\\Users\\user\\Desktop\\專題\\Sustainable_ESG\\上市112水資源管理.csv")
    wasteManage=pd.read_csv("C:\\Users\\user\\Desktop\\專題\\Sustainable_ESG\\上市112廢棄物管理.csv")
    renewableManage=pd.read_csv("C:\\Users\\user\\Desktop\\專題\\Sustainable_ESG\\上市112再生能源.csv")
    employeeManage=pd.read_csv("C:\\Users\\user\\Desktop\\專題\\Sustainable_ESG\\上市112人力發展.csv")
    president=pd.read_csv("C:\\Users\\user\\Desktop\\專題\\Sustainable_ESG\\上市112董事會.csv")
    committee=pd.read_csv("C:\\Users\\user\\Desktop\\專題\\Sustainable_ESG\\上市112功能性委員會.csv")
    inventor=pd.read_csv("C:\\Users\\user\\Desktop\\專題\\Sustainable_ESG\\上市112投資人溝通.csv")
    esg=pd.read_excel("C:\\Users\\user\\Desktop\\專題\\Sustainable_ESG\\專題esg平台資料.xlsx")
    CorpID=greenHouse["公司代號"].values
    emission={}
    for id in CorpID:
        try:
            CorpGreen=greenHouse[greenHouse["公司代號"]==id].values[0]
            CorpWater=waterManage[waterManage["公司代號"]==id].values[0]
            CorpWaste=wasteManage[wasteManage["公司代號"]==id].values[0]
            CorpRenewable=renewableManage[renewableManage["公司代號"]==id].values[0]
            CorpEmployee=employeeManage[employeeManage["公司代號"]==id].values[0]
            CorpPresident=president[president["公司代號"]==id].values[0]
            CorpCommittee=committee[committee["公司代號"]==id].values[0]
            CorpInventor=inventor[inventor["公司代號"]==id].values[0]
            # 保留公司代號和公司名稱，並合併其他數據（跳過前兩個欄位）
            merged_data = [
                CorpGreen[0],CorpGreen[1],CorpGreen[2],CorpGreen[5],CorpGreen[8],CorpGreen[11],
                CorpWater[2],CorpWater[4],CorpWaste[2],CorpWaste[3],CorpWaste[4],CorpWaste[6],
                CorpRenewable[2],CorpEmployee[2],CorpEmployee[3],CorpEmployee[4],CorpEmployee[5],CorpEmployee[6],CorpEmployee[7],CorpEmployee[8],CorpEmployee[10],CorpEmployee[11],CorpEmployee[12],
                CorpPresident[2],CorpPresident[3],CorpPresident[4],CorpPresident[6],CorpPresident[7],CorpCommittee[2],CorpCommittee[3],CorpCommittee[5],CorpCommittee[6],CorpInventor[2]]
            column=["股票代號","公司名稱","直接(範疇一)溫室氣體排放量-數據(公噸CO₂e)","能源間接(範疇二)溫室氣體排放量-數據(公噸CO₂e)","其他間接(範疇三)溫室氣體排放量-數據(公噸CO₂e)","溫室氣體排放密集度-密集度",
                    "用水量(公噸(t))","用水密集度-密集度","有害廢棄物(公噸(t))","非有害廢棄物(公噸(t))","總重量(有害+非有害)(公噸(t))","廢棄物密集度-密集度","再生能源使用率","員工福利平均數(每年6/2起公開)(仟元/人)",
                    "員工薪資平均數(每年6/2起公開)(仟元/人)","非擔任主管職務之全時員工薪資平均數(每年7/1起公開)(仟元/人)","非擔任主管職務之全時員工薪資中位數(每年7/1起公開)(仟元/人)","管理職女性主管占比","職業災害-人數","職業災害-比率",
                    "火災-件數","火災-死傷人數","火災-比率(死傷人數/員工總人數)","董事會席次(席)","獨立董事席次(席)","董事出席董事會出席率","女性董事席次及比率-席","女性董事席次及比率-比率","薪酬委員會席次(席)","薪酬委員會獨立董事席次(席)","審計委員會席次(席)","審計委員會出席率","公司年度召開法說會次數(次)","ESG"
                    ]
            try:
                CorpESG = esg[esg["代號"] == id].values[0]
                merged_data.append(CorpESG[2])
            except IndexError:
                merged_data.append("")  # 或者其他處理方式
            emission[id]=merged_data
        except:
            continue
    dataFrame=pd.DataFrame.from_dict(emission,orient="index",columns=column)
    dataFrame.to_csv("112年Listed_info_emission.csv",encoding='utf-8-sig')

    
def Get_Listed_Mission(companySet):
    url='https://mops.twse.com.tw/mops/web/t05st03'
    for i in companySet.keys():
        # 模擬表單數據
        payload = {
            'co_id': i,
            'step': '1',
            'firstin': '1',
            'off': '1',
            'TYPEK': 'all',
            'queryName': 'co_id',
            'inpuType': 'co_id',
        }

        # 模擬請求的 Headers
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
        }
        try:
            response = requests.post(url, data=payload, headers=headers)
            # 檢查請求是否成功
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, 'html.parser')
                # 更精確地查找 "主要經營業務" 標籤
                main_business_element = soup.find('th', class_='dColor nowrap', colspan="2", string="主要經營業務")
                if main_business_element:
                    td_element = main_business_element.find_next('td')
                    if td_element:
                        main_business = " ".join(td_element.stripped_strings)
                        companySet[i].append(main_business)
                        print(f"{i}的主要經營業務: {main_business}")
                    else:
                        print(f"{i}的主要經營業務資料未找到")
                        companySet[i].append("未找到")
                else:
                    print(f"{i}的主要經營業務標籤未找到")
                    return companySet
                time.sleep(5)
            else:
                print(f"請求失敗，狀態碼：{response.status_code}")
        except :
            return companySet
    return companySet

def esg_combine():
    esg_score=pd.read_excel(r"C:\Users\user\Desktop\ESGBackEnd\兆豐ESG平台分數.xlsx")
    data=pd.read_csv(r"C:\Users\user\Desktop\ESGBackEnd\112年Listed_info_emission_原檔.csv",encoding='utf-8-sig')
    merged_data = pd.merge(data, esg_score[['股票代號', 'ESG','E','S','G']], on="股票代號", how="left")
    merged_data.to_csv("112年Emission_兆豐ESG.csv",encoding='utf-8-sig')
    
if __name__ == "__main__":
    Get_112_EmissionData()

    #載入前100家ESG優良企業排放資訊、個股資訊
    # get_pdf(111)

    # get_pdf(110)


    
    # #下面是優良esg前100家個股資訊
    # emission_data=pd.read_csv("C:\\Users\\user\\OneDrive\\桌面\\專題\\Sustainable_ESG\\Listed_info.csv",encoding='utf-8-sig')
    # emission_company=emission_data["公司名稱"].values
    # print(emission_company)
    # try:
    #     listed_info=get_listed(emission_company)
    # except TimeoutException:
    #     print("等待超時，無法取得上市資料")
    # # #統整前100家裡是上市公司的永續資訊
    # print(listed_info)
    # listed_info=Get_Listed_Mission(listed_info)
    # time.sleep(5)
    # listed_info=get_esg_etf(listed_info)

    # listed_info_dataFrame=pd.DataFrame.from_dict(listed_info,orient='index',columns=["公司代號","公司名稱","產業類別","統一編號","董事長","總經理","資本額","經營業務","入選投組"])
    # listed_info_dataFrame.to_csv("Listed_info_new.csv",encoding='utf-8-sig')
    

    #證期局開罰案例
    #上市、上櫃開罰案例各一個csv
    

    
    # listed_violated=get_Securites_and_futures_violations("上市")
    # listed_violated_dataFrame=pd.DataFrame(listed_violated,columns=["發函日期","公司代號","公司名稱","違規事由","違反法規","裁處情形","處分機關"])
    # print(listed_violated_dataFrame)
    # listed_violated_dataFrame.to_csv("SercuritiesViolations.csv",encoding='utf-8-sig'
    # esg_Portfolio = get_esg_etf()
    # print(esg_Portfolio)

    # data = pd.read_excel(r"C:\Users\user\OneDrive\桌面\專題\Sustainable_ESG\112年ListedInfo.xlsx", thousands=',')
    # newData = data.copy()

    # for i in newData["公司名稱"].values:
    #     listEsg = []  # 移动listEsg到外层循环
    #     for etf_code, etf_details in esg_Portfolio.items():
    #         for record in etf_details["portfolio"]["Data"]:
    #             if i == record[2]:  # 比较公司名称
    #                 listEsg.append(etf_code+etf_details["etf_name"])
        
    #     if not listEsg:
    #         listed = "無"
    #     else:
    #         listed = '、'.join(listEsg)
        
    #     # 使用.loc方法更新DataFrame
    #     newData.loc[newData["公司名稱"] == i, "投資組合"] = listed

    # newData.to_csv(r'C:\Users\user\OneDrive\桌面\專題\Sustainable_ESG\updated_ListedInfo.csv', index=False, encoding='utf-8-sig')
    #esg_combine()
    data=pd.read_csv(r"C:\Users\user\Desktop\ESGBackEnd\112年Emission_兆豐ESG.csv",encoding='utf-8-sig')
    data=data.replace("--","")
    data.to_csv(r"C:\Users\user\Desktop\ESGBackEnd\112年Emission_兆豐ESG.csv",encoding='utf-8-sig')

        
    
    
  
    


    