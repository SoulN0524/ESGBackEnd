import os
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import time

def get_sustainability_reports(save_folder):
    # 使用Selenium啟動瀏覽器（假設使用Chrome）
    # service=Service(r"C:\Users\lin78\geckodriver\geckodriver-v0.34.0-win64\geckodriver.exe")
    # driver = webdriver.Firefox(service=service)
    url = "https://mops.twse.com.tw/mops/web/t100sb11"  # 永續報告書的網頁
    # driver.get(url)



    # #等待網頁加載完成
    # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "year")))
    # input_box = driver.find_element(By.ID, 'year')
    # input_box.send_keys("110")  # 輸入年份
    # search_button = driver.find_element(By.ID, 'search_bar1')
    # search_button.click()

    # html = driver.page_source
    # # 抓取頁面源代碼
    # soup = BeautifulSoup(html, 'html.parser')
    payload = {
        'TYPEK': 'sii',  
        'year': '110',  
        'step': '1',
        'firstin': '1',
    }
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
    }
    session = requests.Session()
    # retries = Retry(total=5, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
    # session.mount('http://', HTTPAdapter(max_retries=retries))
    # session.mount('https://', HTTPAdapter(max_retries=retries))
    try:
        response = session.post(url, data=payload, headers=headers, timeout=30)
        download_links = []
        if response.status_code == 200:
            # 使用 BeautifulSoup 解析網頁內容
            
             soup = BeautifulSoup(response.text, 'html.parser')
             #print(soup)
     
             # 找到所有的下載鏈接
             report_links = soup.find_all('a', href=True)  # 獲取所有帶有 href 的 <a> 標籤
             #print(report_links)

             for link in report_links:
                 href = link['href']
                 if 't100' in href:
                     if href.startswith('/'):
                         full_url = "https://mops.twse.com.tw" + href  # 拼接主域名和相對路徑
                         download_links.append(full_url)
                     elif href.startswith('http'):
                    # 如果 href 是完整的網址，直接使用
                         download_links.append(href)
                #  if 't100' in href and not href.startswith('http'):
                #      #print(href)

                #      full_url = "https://mops.twse.com.tw" + href  # 如果是相對路徑，構造完整 URL
                #          #print(full_url)
                #      download_links.append(full_url)
                #  elif 't100' in href and href.startswith('/'):
                #      full_url = "https://mops.twse.com.tw" + href
                #          #print(full_url)
                #      download_links.append(full_url)
                #  else:
                #      full_url = href
                #      #download_links.append(full_url)
                #      #print(full_url)
                #      download_links.append(full_url)
             print(download_links)
            # 下載報告
             for download_link in download_links:
                 time.sleep(25)
                 download_reports(download_link,save_folder)
                
        else:
            print(f"請求失敗，狀態碼：{response.status_code}")
            return None
    except Exception as e:
        print(f"發生錯誤：{e}")

    # 下載報告書
    #download_reports(report_links, save_folder)

    # 關閉瀏覽器
    # driver.quit()


def is_valid_pdf(file_path):
    """
    檢查文件是否為有效的 PDF 格式。
    :param file_path: 文件的路徑
    :return: 如果文件是有效的 PDF，則返回 True；否則返回 False。
    """
    try:
        with open(file_path, 'rb') as f:
            header = f.read(4)  # 讀取文件頭的前4個字節
            if header == b'%PDF':
                return True
            else:
                print(f"無效的 PDF 文件: {file_path}")
                return False
    except Exception as e:
        print(f"檢查文件時發生錯誤：{e}")
        return False


def download_reports(download_url,save_folder):
    """
    根據提供的報告書下載連結，將每間公司的報告書下載並保存到指定資料夾中。
    
    :param report_links: 字典，包含公司名稱和報告書下載連結。
    :param save_folder: 報告書保存的本地資料夾路徑。
    """
    # 確認保存報告書的資料夾是否存在，若不存在則創建
    #\if not os.path.exists(save_folder):
    #    os.makedirs(save_folder)

    # 逐一下載每間公司的報告書
    #for company, link in report_links.items():
        # print(f"正在下載 {company} 的永續報告書: {link}")
        
        # try:
        #     # 發送HTTP請求以取得PDF內容
        #     response = requests.get(link)
            
        #     # 確認請求成功，狀態碼為200
        #     if response.status_code == 200:
        #         # 設定檔案保存的路徑與名稱
        #         file_path = os.path.join(save_folder, f"{company}_sustainability_report.pdf")
                
        #         # 以二進制寫入模式將內容保存到本地檔案
        #         with open(file_path, 'wb') as f:
        #             f.write(response.content)
                
        #         print(f"{company} 的報告書已保存至 {file_path}")
        #     else:
        #         print(f"下載 {company} 的報告書失敗，狀態碼: {response.status_code}")
        
        # except Exception as e:
        #     print(f"下載 {company} 的報告書時發生錯誤: {e}")
    try:
        response = requests.get(download_url, stream=True)
        
        if response.status_code == 200:
             file_name = download_url.split('/')[-1]  # 獲取文件名
             file_path = os.path.join(save_folder, file_name)  # 指定下載位置

             content_type = response.headers.get('Content-Type')
             if content_type != "application/pdf" and not file_name.endswith('.pdf'):
                 print(f"無法下載 {download_url}，非 PDF 文件，Content-Type: {content_type}")
                 return  # 退出下載該文件的過程

             if os.path.exists(file_path):
                 print(f"{file_name} 已經存在，跳過下載")
                 return  # 跳過已存在的文件
             with open(file_path, 'wb') as f:
                 for chunk in response.iter_content(chunk_size=8192):
                     f.write(chunk)
                     #f.write(response.content)  # 下載並保存文件
                 print(f"已下載: {file_name}到{save_folder}")

             if is_valid_pdf(file_path):
                 print(f"已成功下載並驗證: {file_name} 到 {save_folder}")
             else:
                 print(f"下載的文件無效或損壞，刪除: {file_name}")
                 os.remove(file_path)  # 如果文件無效，刪除它
        else:
            print(f"無法下載 {download_url}，狀態碼：{response.status_code}")
    except Exception as e:
        print(f"下載時發生錯誤：{e}")


# 範例使用
if __name__ == "__main__":
    save_folder = r"C:\Users\lin78\OneDrive\文件\永續report"  # 指定保存報告書的資料夾
    if not os.path.exists(save_folder):
        os.makedirs(save_folder)  # 如果資料夾不存在，則創建

    # 呼叫函式來抓取並下載永續報告書
    get_sustainability_reports(save_folder)
