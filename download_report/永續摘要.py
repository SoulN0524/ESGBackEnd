import os
import pdfplumber
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
import pandas as pd
from langchain_community.document_loaders import PDFPlumberLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import OpenAIEmbeddings
from langchain_community.vectorstores import FAISS
#from langchain_community.chat_models import ChatOpenAI
from langchain.prompts import ChatPromptTemplate
from langchain.chains import RetrievalQA
from langchain_openai import ChatOpenAI
api_key = "your_api_key"
import re
from langchain.schema import Document
# Load the PDF

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

def pdf_loader(file,size=500,overlap=50):
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
  db = FAISS.from_documents(new_doc, OpenAIEmbeddings(api_key=api_key))
  return db

#db = pdf_loader(r"C:\Users\lin78\OneDrive\文件\t100sa11_1101_110.pdf",500,50)

# 提示模板
prompt = ChatPromptTemplate.from_messages([
    ("system",
     "你是一个整理永续报告书内容的助手，你的工作是帮助用户根据他们提供的公司或股票代码整理永续报告书的内容，并生成摘要,"
     "如果有明確數據或技術(產品)名稱可以用數據或名稱回答,"
     "回答以繁體中文和台灣用語為主。"
     "{context}"),
    ("human","{question}")])

# 建立問答函式
def question_and_answer(db,question):
    llm = ChatOpenAI(model="gpt-4", api_key=api_key)
    qa = RetrievalQA.from_llm(llm=llm,
                              prompt=prompt,
                              return_source_documents=True,
                              retriever=db.as_retriever(
                                  search_kwargs={'k':10}))
    result = qa.invoke(question)
    return result

def process_reports_in_folder(folder_path, question):
    """針對資料夾內所有PDF文件生成500字摘要"""
    data = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".pdf"):
                file_path = os.path.join(root, file)
                print(f"正在處理文件: {file_path}")

                # 加載文件
                db = pdf_loader(file_path)

                # 生成摘要
                result = question_and_answer(db, question)
                summary = result['result']
                print(f"摘要: \n{summary}")
                print('_________')

                data.append({"文件名稱":file,"文件路徑":file_path,"摘要": summary})
    df=pd.DataFrame(data)
    return df

#主程序：針對不同版本生成摘要
def generate_summaries_for_versions(base_folder,output_file="summaries.xlsx"):
    # 定義不同版本的資料夾
    chinese_folder = os.path.join(base_folder, "110永續報告書(中文)")
    english_folder = os.path.join(base_folder, "110永續報告書(英文)")
    revised_folder = os.path.join(base_folder, "110永續報告書(修訂版)")
    question = "請給我此份永續報告書的500字摘要"
    # result=question_and_answer(question)
    # print(result['result'])
    # print('_________')
    # 處理各個資料夾中的PDF文件
    print("正在處理中文版報告書...")
    chinese_df=process_reports_in_folder(chinese_folder, question)

    print("正在處理英文版報告書...")
    english_df=process_reports_in_folder(english_folder, question)

    print("正在處理修訂版報告書...")
    revised_df=process_reports_in_folder(revised_folder, question)

    # 使用 pandas ExcelWriter 將三個資料框寫入同一個 Excel 檔案的不同工作表
    with pd.ExcelWriter(output_file) as writer:
        chinese_df.to_excel(writer, sheet_name="中文版", index=False)
        english_df.to_excel(writer, sheet_name="英文版", index=False)
        revised_df.to_excel(writer, sheet_name="修訂版", index=False)
    print(f"所有摘要已儲存到 {output_file}")

# 執行程式
base_folder = r"C:\Users\lin78\OneDrive\文件\永續report"
generate_summaries_for_versions(base_folder)
# query = ""
# docs = db.similarity_search(query,k=3)
# for i in docs:
#     print(i.page_content)
#     print('_________')