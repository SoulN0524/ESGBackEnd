import os
import pdfplumber
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
import pandas as pd
import requests
from langchain_community.document_loaders import PDFPlumberLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
import re
from langchain.schema import Document
import time
from openpyxl import load_workbook

#  定義 API 的 URL 和標頭
API_KEY = "AIzaSyDSPg9Y7i2jL1Ja8PlV12N44PaXB9usYYU"  # 替換為您的 Google AI API Key
MODEL = "models/gemini-1.5-pro"
API_URL = f"https://generativelanguage.googleapis.com/v1/{MODEL}:generateContent?key={API_KEY}"

# 檢查 PDF 文件
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

def pdf_loader(file,size=1000,overlap=100):
  if is_scanned_pdf(file):
    print(f"{file} 是掃描件，使用 OCR 進行處理。")
        # 使用 pytesseract 提取掃描 PDF 的文本
    with pdfplumber.open(file) as pdf:
        text = ""
        for page in pdf.pages:
            # 將 PDF 頁面轉換成 PIL 影像格式
            pil_image = page.to_image(resolution=300).original
            pil_image = pil_image.convert("L")
            # 使用 Tesseract 提取文字
            page_text = pytesseract.image_to_string(pil_image,lang='chi_tra')
            text += page_text
        doc = [Document(page_content=text)]
  else:
    print(f"{file} 不是掃描件，正常加載。")
    loader = PDFPlumberLoader(file)
    doc = loader.load()

  text_splitter = RecursiveCharacterTextSplitter(
                          chunk_size=size,
                          chunk_overlap=overlap)
  new_doc = text_splitter.split_documents(doc)
  return new_doc

def question_and_answer_google(new_doc, question):
    """使用 Google AI Studio 做摘要"""
    context = " ".join([doc.page_content for doc in new_doc])

    # 構造 API 請求
    data= {
    "contents": {
            "parts": [
              {
                "text":f"你是一个整理永續報告書内容的助手，你的工作是幫助用戶根據他們提供的公司或股票代碼整理永續報告書的内容，並生成摘要,如果有明確數據或技術(產品)名稱可以用數據或名稱回答,回答以繁體中文和台灣用語為主，請根據以下的內容來生成摘要："
                "{context}\n\n{question}"
              }
            ]
           }
        }
    time.sleep(1)

    # 調用 Google AI Studio API
    response = requests.post(API_URL, json=data)
    if response.status_code == 200:
        result = response.json()
        try:
            return result['candidates'][0]['output']
        except (KeyError, IndexError):
            return "無法生成回應"
    else:
        print(f"API 請求失敗，狀態碼: {response.status_code}")
        return "API 請求失敗"
    
def write_to_excel(output_file, sheet_name, data):
    if not os.path.exists(output_file):
        # 如果文件不存在，創建新文件
        pd.DataFrame(data).to_excel(output_file, sheet_name=sheet_name, index=False)
    else:
        # 如果文件已存在，追加到對應的 sheet
        book = load_workbook(output_file)
        writer = pd.ExcelWriter(output_file, engine="openpyxl")
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}
        df_existing = pd.read_excel(output_file, sheet_name=sheet_name)
        df_new = pd.concat([df_existing, pd.DataFrame(data)], ignore_index=True)
        df_new.to_excel(writer, sheet_name=sheet_name, index=False)
        writer.save()

def process_reports_in_folder_google(folder_path,sheet_name,output_file, question):
    data = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".pdf"):
                file_path = os.path.join(root, file)
                print(f"正在處理文件: {file_path}")

                # 加載文件
                new_doc = pdf_loader(file_path)

                # 生成摘要
                summary = question_and_answer_google(new_doc, question)
                print(f"摘要: \n{summary}")
                print('_________')

                data.append({"文件名稱": file, "文件路徑": file_path, "摘要": summary})
                write_to_excel(output_file, sheet_name, data)
        
    df = pd.DataFrame(data)
    return df

# 主程序：針對不同版本生成摘要
def generate_summaries_for_versions_google(base_folder, output_file="summaries_google.xlsx"):
    # 定義不同版本的資料夾
    chinese_folder = os.path.join(base_folder, "110永續報告書(中文)")
    english_folder = os.path.join(base_folder, "110永續報告書(英文)")
    revised_folder = os.path.join(base_folder, "110永續報告書(修訂版)")
    question = "請給我此份永續報告書的500字摘要"

    print("正在處理中文版報告書...")
    process_reports_in_folder_google(chinese_folder,"中文版",output_file, question)

    print("正在處理英文版報告書...")
    process_reports_in_folder_google(english_folder,"英文版",output_file, question)

    print("正在處理修訂版報告書...")
    process_reports_in_folder_google(revised_folder,"修訂版",output_file,question)


    print(f"所有摘要已儲存到 {output_file}")

# 執行程式
base_folder = r"C:\Users\lin78\OneDrive\文件\永續report"
generate_summaries_for_versions_google(base_folder)