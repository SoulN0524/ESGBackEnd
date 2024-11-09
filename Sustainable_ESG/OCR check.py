import os
import pdfplumber
import pytesseract
import re

# 設定 tesseract 的路徑，請根據實際安裝位置進行調整
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\lin78\OneDrive\文件\Sustainable\tesseract-5.4.1\tesseract-5.4.1'

def is_scanned_pdf(file_path, check_pages=10):
    """檢查 PDF 是否為掃描件"""
    cid_pattern = re.compile(r"(cid:\d+|\(cid:\d+\))") # 用於檢測 CID 字元和特殊符號的模式
    with pdfplumber.open(file_path) as pdf:
        for i in range(min(check_pages, len(pdf.pages))):
            page = pdf.pages[i]
            text = page.extract_text()
           
            # 如果頁面沒有提取到任何有效文字，或提取內容主要是 CID 字元/特殊符號，判斷為掃描件
            if text is None or len(cid_pattern.findall(text)) > (len(text) * 0.025):
                if len(page.images) > 1:  # 確保頁面包含圖片
                    return True
        return False  # 如果可以提取自然語言文本，則認為不是掃描件

def test_scanned_pdf_detection(folder_path, check_pages=10):
    """測試指定資料夾中的所有 PDF 檔案是否能正確判斷為掃描件"""
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".pdf"):
                file_path = os.path.join(root, file)
                is_scanned = is_scanned_pdf(file_path, check_pages=check_pages)
                print(f"{file}: {'掃描件' if is_scanned else '非掃描件'}")



# 指定要測試的 PDF 檔案資料夾
test_folder = r"C:\Users\lin78\OneDrive\文件\永續report\110永續報告書(中文)"
test_scanned_pdf_detection(test_folder)

