import os
import requests
from bs4 import BeautifulSoup

def get_expected_counts_from_web():
    """
    從公開資訊觀測站抓取各分類永續報告的預期數量。
    
    :return: 包含各分類報告數量的字典 expected_counts。
    """
    url = "https://mops.twse.com.tw/mops/web/t100sb11"
    payload = {
        'TYPEK': 'sii',  # 股票分類
        'year': '110',   # 指定年度
        'step': '1',
        'firstin': '1',
    }
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
    }

    try:
        response = requests.post(url, data=payload, headers=headers)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            download_links = soup.find_all('a', href=True)

            chinese_count = 0
            english_count = 0
            revised_count = 0

            # 根據下載連結的格式，判斷語言和版本
            for link in download_links:
                href = link['href']
                if 't100' in href:  # 只處理包含 t100 的連結
                    if '110_E' in href:
                        english_count += 1
                    elif '110_M' in href:
                        revised_count += 1
                    elif '110' in href:
                        chinese_count += 1

            expected_counts = {
                'Chinese': chinese_count,
                'English': english_count,
                'Revised': revised_count
            }

            print(f"抓取到的預期數量: {expected_counts}")
            return expected_counts

        else:
            print(f"請求失敗，狀態碼：{response.status_code}")
            return None

    except Exception as e:
        print(f"發生錯誤：{e}")
        return None

def check_all_folders(base_folder, expected_counts):
    """
    檢查所有分類資料夾內的報告書數量。
    
    :param base_folder: 主資料夾路徑，包含中文、英文和修訂版資料夾。
    :param expected_counts: 每個資料夾的預期檔案數量。
    """
    # 定義各個語言和版本的資料夾路徑
    chinese_folder = os.path.join(base_folder, "110永續報告書(中文)")
    english_folder = os.path.join(base_folder, "110永續報告書(英文)")
    revised_folder = os.path.join(base_folder, "110永續報告書(修訂版)")

    # 檢查每個資料夾
    count_reports_in_folder(chinese_folder, expected_counts['Chinese'])
    count_reports_in_folder(english_folder, expected_counts['English'])
    count_reports_in_folder(revised_folder, expected_counts['Revised'])

def count_reports_in_folder(folder_path, expected_count):
    """
    確認資料夾中的報告數量是否與預期一致。
    
    :param folder_path: 資料夾的路徑。
    :param expected_count: 預期的報告書數量。
    :return: 資料夾內的實際檔案數量與是否符合預期。
    """
    # 確認資料夾是否存在
    if not os.path.exists(folder_path):
        print(f"資料夾 {folder_path} 不存在！")
        return 0, False

    # 獲取資料夾內所有PDF檔案
    files = [file for file in os.listdir(folder_path) if file.endswith('.pdf')]
    actual_count = len(files)

    # 比較實際數量與預期數量
    if actual_count == expected_count:
        print(f"資料夾 '{folder_path}' 的檔案數量正確：{actual_count}/{expected_count}")
        return actual_count, True
    else:
        print(f"資料夾 '{folder_path}' 的檔案數量不符：{actual_count}/{expected_count}")
        return actual_count, False


# 範例使用
if __name__ == "__main__":
    # 設定主資料夾路徑
    base_folder = r"C:\Users\lin78\OneDrive\文件\永續report"

    # 抓取預期檔案數量
    expected_counts = get_expected_counts_from_web()

    # 若成功抓取到數量，則進行檢查
    if expected_counts:
        check_all_folders(base_folder, expected_counts)
