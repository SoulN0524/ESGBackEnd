import os
import getpass
# LangChain 相關套件
from langchain.document_loaders import PDFPlumberLoader             # 讀取 PDF
from langchain.text_splitter import RecursiveCharacterTextSplitter  # 字串切割
from langchain.embeddings import OpenAIEmbeddings                   # 嵌入方法
from langchain.vectorstores import FAISS                            # 向量資料庫
from langchain.chat_models import ChatOpenAI
from langchain.prompts import ChatPromptTemplate # 提示模板
from langchain.chains import RetrievalQA         # 問答模組      
os.environ['OPENAI_API_KEY'] = getpass.getpass('OpenAI API Key:')
llm_4_turbo = ChatOpenAI(model="gpt-4-turbo")
def pdf_loader(file,size,overlap):
  loader = PDFPlumberLoader(file)
  doc = loader.load()
  text_splitter = RecursiveCharacterTextSplitter(
                          chunk_size=size,
                          chunk_overlap=overlap)
  new_doc = text_splitter.split_documents(doc)
  db = FAISS.from_documents(new_doc, OpenAIEmbeddings())
  return db

query="公司近期的永續政策是甚麼?"
db=pdf_loader('sustainable_report\台積電esgreport.pdf',500,50)
docs = db.similarity_search(query,k=3)
for i in docs:
    print(i.page_content)
    print('______')


