import numpy as np
from bs4 import BeautifulSoup
import pandas as pd
import requests
import datetime as dt
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
#建照瀏覽器
def init_browser():
  service=Service(executable_path=r"C:\\Users\\user\\geckodriver.exe")
  # 初始化 Chrome 瀏覽器
  browser = webdriver.Firefox(service=service)
  # 設定瀏覽器視窗最大化
  browser.set_window_size(640, 960)
  return browser
def catch_businesstoday(company_name):
  # 設定搜尋的網址與查詢關鍵字
  search_url = f'https://esg.businesstoday.com.tw/group_search/get_articles_by_keyword?keywords={company_name}'
  headers={
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
  }

  # 送出 GET 請求
  data =requests.get(search_url, headers=headers).json()
  items=data["article_list"]
  columns=['股名','標題','連結','內容']
  news=[]

  for item in items:
    print("----------------------------------")
    print("新聞"+str(item))
    print("日期:"+str(item[""]))
    news_id=item["id"]
    title=item["title"]
    link=item["href"]
    url=requests.get(link).content
    soup=BeautifulSoup(url,'html.parser')
    p_elements=soup.find_all('p')
    p=""
    for paragraph in p_elements:
        p+=paragraph.get_text()
    news.append([company_name,title,link,p])
 # df=pd.DataFrame(news,columns=columns)
  print(df)


def catch_CSR(company_name):
  #company_name="台積電"
  browser=init_browser()
  search_url=f"https://csr.cw.com.tw/search/doSearch.action?keyword={company_name}&channel=0"
  browser.get(search_url)
  element = WebDriverWait(browser, 10).until(EC.visibility_of_element_located((By.XPATH, "//*[@id=\"search_vue\"]/section[2]/section")))
  html = browser.page_source
  soup=BeautifulSoup(html, 'html.parser')
  # 選取文章列表中的所有 <li> 標籤
  articles = soup.select("section.articleList ul li")

  # 初始化結果列表
  result = []

  for article in articles:
    # 提取標題
    title_tag = article.select_one(".txt h2 a")
    title = title_tag.get_text(strip=True) if title_tag else "無標題"

    # 提取描述
    description_tag = article.select_one(".txt p")
    description = description_tag.get_text(strip=True) if description_tag else "無描述"

    # 提取鏈接
    link = title_tag["href"] if title_tag and "href" in title_tag.attrs else "無鏈接"

    result.append({
        "標題": title,
        "描述": description,
        "鏈接": link
    })
  browser.quit()
  # 打印結果
  for item in result:
    print(item)
  
def catch_esgtimes(company_name):
  headers={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
  }
  url=f'https://esgtimes.com.tw/?s={company_name}'
  response=requests.get(url,headers=headers)
  soup=BeautifulSoup(response.content,'html.parser')
  articles=soup.find_all("article",class_="jeg_post")
  # 存儲結果
  news_list = []
  for article in articles:
    # 標題和連結
    title_tag = article.find('h3', class_='jeg_post_title').find('a')
    title = title_tag.text.strip()
    link = title_tag['href']
    
    # 日期
    date_tag = article.find('div', class_='jeg_meta_date')
    date = date_tag.text.strip() if date_tag else None
    
    # 簡介
    excerpt_tag = article.find('div', class_='jeg_post_excerpt')
    excerpt = excerpt_tag.text.strip() if excerpt_tag else None
    
    
    # 組裝結果
    news_list.append({
      '公司':company_name,
      '日期': date,
      '新聞標題': title,
      '新聞連結': link
    })

  # 查看結果
  for news in news_list:
      print(news)
      
#中央社資料(但輸入company_name要像這樣ex:聯發科 ESG)
def catch_cna(company_name):
  headers={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
  }
  catch_esgtimes(company_name[:3])
  url=f"https://www.cna.com.tw/search/hysearchws.aspx?q={company_name}"
  response=requests.get(url,headers=headers)
  soup=BeautifulSoup(response.content,'html.parser')
  mainList=soup.find("ul",class_="mainList")
  news_list=[]

  for row in mainList.find_all("li"):
    # 標題和連結
    title_tag = row.find('h2')
    if not title_tag:
        continue
    title = title_tag.text.strip()
    link_tag = row.find('a')
    link = f"https://www.cna.com.tw{link_tag['href']}" if link_tag else None

    # 日期
    date_tag = row.find('div', class_='date')
    date = date_tag.text.strip()[:10] if date_tag else None  # 僅保留 YYYY/MM/DD

    # 組裝結果
    news_list.append({
        '公司': company_name[:3],
        '日期': date,
        '新聞標題': title,
        '新聞連結': link
    })
  print(news_list)