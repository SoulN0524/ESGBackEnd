import os
import requests
from bs4 import BeautifulSoup

def get_expected_reports_from_web():
    """
    使用爬蟲抓取公開資訊觀測站上應該下載的永續報告書清單。
    year: 要抓取的年份，例如 '110'
    """
    url = "https://mops.twse.com.tw/mops/web/t100sb11"
    payload = {
        'TYPEK': 'sii',  
        'year': '110',  
        'step': '1',
        'firstin': '1',
    }
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
    }

    response = requests.post(url, data=payload, headers=headers)

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        report_links = soup.find_all('a', href=True)

        expected_reports = {
            '中文': [],
            '英文': [],
            '修訂版': []
            }
        
        for link in report_links:
            href = link['href']
            if 't100' in href:  # 選擇報告書下載連結
                report_name = link.get_text(strip=True)  # 假設文件名稱是連結文字 + .pdf
                if '110_E' in report_name or '110_E' in href:
                    expected_reports['英文'].append(report_name)
                elif '110_M' in report_name or '110_M' in href:
                    expected_reports['修訂版'].append(report_name)
                elif '110' in report_name or '110' in href:
                    expected_reports['中文'].append(report_name)

        return expected_reports
    else:
        print(f"無法抓取網站資料，狀態碼：{response.status_code}")
        return []

def get_actual_reports(folder_path):
    """
    取得指定資料夾內的所有檔案名稱列表
    """
    return [f for f in os.listdir(folder_path) ]

def normalize_filename(filename):
    """
    將檔名標準化，去除&fileName=等多餘的前綴，並返回檔名主體部分。
    """
    if '&fileName=' in filename:
        return filename.split('&fileName=')[-1]
    return filename

def find_missing_reports(expected_reports, actual_reports):
    """
    比對兩個報告書列表，找出缺失的報告書
    """
    normalized_expected = [normalize_filename(report) for report in expected_reports]
    normalized_actual = [normalize_filename(report) for report in actual_reports]
    missing_reports = [report for report in normalized_expected if report not in normalized_actual]
    return missing_reports

def check_missing_reports(folder_path,report_type):
    """
    主函數，檢查資料夾中缺失的報告書
    """
    expected_reports = get_expected_reports_from_web()
    actual_reports = get_actual_reports(folder_path)
    
    missing_reports = find_missing_reports(expected_reports[report_type], actual_reports)
    
    if missing_reports:
        print(f"缺失的報告書: {missing_reports}")
    else:
        print("所有報告書都已下載。")

# 指定資料夾路徑和年份
chinese_folder = r"C:\Users\lin78\OneDrive\文件\永續report\110永續報告書(中文)"
english_folder = r"C:\Users\lin78\OneDrive\文件\永續report\110永續報告書(英文)"
revised_folder = r"C:\Users\lin78\OneDrive\文件\永續report\110永續報告書(修訂版)"

# 分別檢查不同版本資料夾中的缺失報告
print("檢查中文版本:")
check_missing_reports(chinese_folder,'中文')

print("\n檢查英文版本:")
check_missing_reports(english_folder,'英文')

print("\n檢查修訂版本:")
check_missing_reports(revised_folder,'修訂版')
