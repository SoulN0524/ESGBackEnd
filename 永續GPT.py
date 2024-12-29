import os
import pdfplumber
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
import pandas as pd
import requests
from langchain_community.document_loaders import PDFPlumberLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import OpenAIEmbeddings
from langchain_community.vectorstores import FAISS
from langchain.prompts import ChatPromptTemplate
from langchain.chains import RetrievalQA
from langchain_openai import ChatOpenAI
from langchain.schema import Document, AIMessage
import re

api_key = "your api key"

def is_scanned_pdf(file_path, check_pages=10):
    """檢查 PDF 是否為掃描件"""
    cid_pattern = re.compile(r"(cid:\d+|\(cid:\d+\))")
    with pdfplumber.open(file_path) as pdf:
        for i in range(min(check_pages, len(pdf.pages))):
            page = pdf.pages[i]
            text = page.extract_text()
            
            if text is None or len(cid_pattern.findall(text)) > (len(text) * 0.025):
                if len(page.images) > 1:
                    return True
        return False

def pdf_loader(file, size=500, overlap=50):
    """加載 PDF 文件，處理掃描或非掃描 PDF，並標註頁碼"""
    if is_scanned_pdf(file):
        #print(f"{file} 是掃描件，使用 OCR 進行處理。")
        with pdfplumber.open(file) as pdf:
            text = ""
            documents = []
            for page_num, page in enumerate(pdf.pages, start=1):
                pil_image = page.to_image(resolution=300).original
                pil_image = pil_image.convert("L")
                page_text = pytesseract.image_to_string(pil_image, lang='chi_tra')
                documents.append(Document(page_content=page_text, metadata={'page_number': page_num, 'source': file}))
    else:
        #print(f"{file} 不是掃描件，正常加載。")
        loader = PDFPlumberLoader(file)
        doc = loader.load()
        documents = []
        for page_num, document in enumerate(doc, start=1):
            documents.append(Document(page_content=document.page_content, metadata={'page_number': page_num, 'source': file}))

    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=size,
        chunk_overlap=overlap
    )
    new_doc = text_splitter.split_documents(documents)
    db = FAISS.from_documents(new_doc, OpenAIEmbeddings(api_key=api_key))
    return db

prompt = ChatPromptTemplate.from_messages([
    ("system",
     "你是一个永續報告書的智能助手，可以根據使用者提供的檔案以及透過查詢網路上的資料回答問題，並提供具體的來源頁碼讓使用者可以確認資料來源的真實性。請在回答中標註每個信息的來源頁碼,"
     "如果有明確數據或技術(產品)名稱可以用數據或名稱回答,"
     "回答以繁體中文和台灣用語為主。"
     "{context}"),
    ("human", "{question}")
])

def question_and_answer(db, question):
    """根據使用者問題，查詢並返回答案及來源頁碼"""
    llm = ChatOpenAI(model="gpt-4", api_key=api_key)
    qa = RetrievalQA.from_llm(
        llm=llm,
        prompt=prompt,
        return_source_documents=True,
        retriever=db.as_retriever(
            search_kwargs={'k': 10}
        )
    )
    result = qa.invoke(question)
    answers = result['result']
    source_documents = result['source_documents']

    detailed_answer = answers + "\n\n來源頁碼:\n"
    for doc in source_documents:
        detailed_answer += f"文件: {doc.metadata.get('source', '未知')} - 頁碼: {doc.metadata.get('page_number', '未知')}\n"
    return detailed_answer

# 使用 GPT 模型直接回答問題（無文件時）
def gpt_answer_directly(question):
    """直接使用 GPT 生成回答"""
    llm = ChatOpenAI(model="gpt-4", api_key=api_key)
    messege = [{"role": "system", "content": "你是一個資深永續專家，在國外有做過十年以上ESG相關類型的工作，對於任何有關ESG的問題你都能使用簡單易懂的詞彙解釋給永續小白聽，如果有明確數據或技術(產品)名稱可以用數據或名稱回答，回答以繁體中文和台灣用語為主："},
               {"role": "user", "content": question}]
    response = llm.invoke(messege)
    if isinstance(response, AIMessage):  # 檢查返回值是否為 AIMessage
        return response.content  # 使用屬性訪問內容
    else:
        raise TypeError("Unexpected response type: Expected AIMessage.")
    
def main(file_path=None, question=None):
    # 如果沒有提供檔案，則查詢網路資料
    if file_path is None:
        if question is None:
            raise ValueError("若未提供檔案，必須提供問題進行查詢。")
        #print("正在查詢網路資料...")
        answer = gpt_answer_directly(question)
        print("碧綠知音:", answer)
    else:
        #print(f"正在處理文件: {file_path}")

        db = pdf_loader(file_path)
        detailed_answer = question_and_answer(db, question)

        print("碧綠知音:\n", detailed_answer)

# 測試範例：上傳文件，並提問
if __name__ == '__main__':
    # 測試PDF文件
    # file_path = r"C:\Users\lin78\OneDrive\文件\永續report\110永續報告書(中文)\&fileName=t100sa11_1215_110.pdf"  # 請替換為您的PDF文件路徑
    # question = "這份報告的碳排放量是多少？"
    # main(file_path, question)

    # 測試沒有檔案的情況，直接從網路上查詢
    question = "2020年全球氣候變遷的挑戰是什麼？"
    main(question=question)

