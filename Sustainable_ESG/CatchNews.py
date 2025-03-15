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
import re
from datetime import datetime
#建照瀏覽器
def init_browser():
  service=Service(executable_path=r"C:\\Users\\user\\geckodriver.exe")
  # 初始化 Chrome 瀏覽器
  browser = webdriver.Firefox(service=service)
  # 設定瀏覽器視窗最大化
  browser.set_window_size(640, 960)
  return browser
#公司名稱、日期、新聞標題、新聞連結、新聞內容、來源
#獲利能力排序介面
def catch_businesstoday(company_name,page):#這只能抓到2022年10-11月的資料，後續再做調整
  news=[]
  headers={
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
  }
  for i in range(page):
    search_url = f'https://esg.businesstoday.com.tw/group_search/get_articles_by_keyword?keywords={company_name}&page={i+1}'
    print("連結"+search_url)
    # 送出 GET 請求
    try:
        response = requests.get(search_url, headers=headers)
        response.raise_for_status()  # 檢查 HTTP 狀態碼

        try:
            data = response.json()
        except requests.exceptions.JSONDecodeError:
            print("無法解析 JSON，回應內容:", response.text)
            continue
        
        items = data.get("article_list", [])
        if not items:
            print(f"頁面 {i+1} 沒有找到新聞")
            break
        
        for item in items:
            date = item.get("url_query", "")[:8]
            title = item.get("title", "無標題")
            link = item.get("href", "無連結")
            p = item.get("part_text", "").strip()
            p = re.sub(r'\t|\n|\r', '', p)

            news.append({
                "公司名稱": company_name,
                "日期": date,
                "新聞標題": title,
                "新聞連結": link,
                "新聞內容": p,
                "新聞來源": "ESG今周刊"
            })
    except requests.RequestException as e:
        print(f"HTTP 請求錯誤: {e}")
        continue

    return news

def CSR_dealer(json_data):
  browser=init_browser()
  for i in json_data:
    browser.get(i["新聞連結"])
    html = browser.page_source
    soup=BeautifulSoup(html, 'html.parser')
    date=soup.find("li",class_="time")
    date=re.sub('-','',date)
    p=''
    p_elements=soup.find_all('p')
    for paragraph in p_elements:
      p+=paragraph.get_text().strip()
    p=re.sub(r'\t|\n|\r','',p)
    json_data[i]["日期"]=date
    json_data[i]["新聞描述"]=p[:100]
  return json_data


#這個抓很久
def catch_CSR(company_name, page):
  browser = init_browser()
  result = []

  for i in range(page):
      time.sleep(5)
      search_url = f"https://csr.cw.com.tw/search/doSearch.action?keyword={company_name}&channel=0&page={i+1}"

      try:
          print(f"抓取頁面 {i+1}: {search_url}")
          browser.get(search_url)

          WebDriverWait(browser, 10).until(
              EC.visibility_of_element_located((By.XPATH, "//*[@id=\"search_vue\"]/section[2]/section"))
          )

          html = browser.page_source
          soup = BeautifulSoup(html, 'html.parser')
          articles = soup.select("section.articleList ul li")

          for article in articles:
              title_tag = article.select_one(".txt h2 a")
              title = title_tag.get_text(strip=True) if title_tag else "無標題"
              description_tag = article.select_one(".txt p")
              description = description_tag.get_text(strip=True) if description_tag else "無描述"
              link = title_tag["href"] if title_tag and "href" in title_tag.attrs else "無鏈接"
              
              result.append({
                  '公司名稱': company_name,
                  "日期": "",
                  "新聞標題": title,
                  "新聞連結": link,
                  "新聞內容": description,
                  '新聞來源': "CSR天下"
              })
      
      except Exception as e:
          print(f"無法獲取 CSR 天下新聞 (第 {i+1} 頁): {e}")
          continue
  
  browser.quit()
  return CSR_dealer(result)

#網頁的更多文章無法讀取
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
    date=datetime.strptime(date,"%Y 年 %m 月 %d 日")
    date = date.strftime("%Y%m%d")
    # 簡介
    excerpt_tag = article.find('div', class_='jeg_post_excerpt')
    excerpt = excerpt_tag.text.strip() if excerpt_tag else None
    
    
    # 組裝結果
    news_list.append({
      '公司':company_name,
      '日期': date,
      '新聞標題': title,
      '新聞連結': link,
      "新聞內容":excerpt,
      "新聞來源":"ESG_TIMES"
    })
  return news_list

      
#中央社資料(但輸入company_name要像這樣ex:聯發科 ESG)
def catch_cna(company_name):
  headers={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
  }
  space_index=company_name.find(" ")
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
    response=requests.get(link,headers=headers)
    soup=BeautifulSoup(response.content,'html.parser')
    p=''
    p_elements=soup.find_all('p')
    for paragraph in p_elements:
      p+=paragraph.get_text().strip()
    p=re.sub(r'\t|\n|\r','',p)
    # 日期
    date_tag = row.find('div', class_='date')
    date = date_tag.text.strip()[:10] if date_tag else None  # 僅保留 YYYY/MM/DD
    date=date.replace("/","")
    # 組裝結果
    news_list.append({
        '公司': company_name[:space_index],
        '日期': date,
        '新聞標題': title,
        '新聞連結': link,
        "新聞內容":p[:100],
        "新聞來源":"中央社"
    })
  return news_list

#想不到怎麼處理
def catch_CSRone(company_name):
  headers={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
  }
  url=f"https://csrone.com/search?q={company_name}&c=news#gsc.tab=0&gsc.q=inurl%3Anews%20intext%3A{company_name}%20OR%20intitle%3A{company_name}&gsc.sort=&gsc.page=1"
  reponse=requests.get(url,headers=headers)
  soup=BeautifulSoup(reponse.content,'html.parser')
  print(soup)

#日後來看
def catch_ESGTW(company_name):
  url=f"https://www.esgtw.net/page/1/?s={company_name}"
  headers={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
  }
  response=requests.get(url,headers=headers)
  soup=BeautifulSoup(response.content,'html.parser')
  

def catch_ESGView(company_name,page):
  headers={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
  }
  news_list=[]
  try:
    for i in range(page):
      url=f"https://esg.gvm.com.tw/page/{i+1}?s={company_name}"
      

      response=requests.get(url,headers=headers)
      soup=BeautifulSoup(response.content,'html.parser')

      articles = soup.find_all('div', class_='list_post row my-4')
      for article in articles:
            # 抓取標題
        title_tag = article.find('h2', class_='content_title')
        title = title_tag.get_text(strip=True) if title_tag else '無標題'
        
        # 抓取新聞連結
        link_tag = title_tag.find('a') if title_tag else None
        link = link_tag['href'] if link_tag else '無連結'
        
        # 抓取日期
        date_tag = article.find('span', class_='list_post_datetime')
        date = date_tag.get_text(strip=True) if date_tag else '無日期'
        date=date.replace("/","")
        # 抓取新聞內容
        content_tag = article.find('div', class_='list_post_excerpt')
        content = content_tag.get_text(strip=True)[:100] if content_tag else '無內容'
        
        # 抓取作者名稱（可視為公司名）
        author_tag = article.find('span', class_='list_post_author')
        author_name = author_tag.get_text(strip=True).replace('文', '').strip() if author_tag else '無作者'

        # 將資料添加至列表
        news_list.append({
            '公司': company_name,
            '日期': date,
            '新聞標題': title,
            '新聞連結': link,
            '新聞內容': content,
            '新聞來源': "ESG遠見"  # 此處的來源可按需更改
        })
  except:
    if len(news_list)<1:
      return None
    else:
      return news_list
  return news_list

#爬失敗
def catch_ETtoday(company_name,page):

  headers={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
  }
  news_list=[]
  try:
    for i in range(page):
      url=f'https://esg.ettoday.net/search/{company_name}/{i+1}'
      response=requests.get(url,headers=headers)
      soup=BeautifulSoup(response.content,'html.parser')
      print(soup)
      # 尋找所有的新聞組件
      articles = soup.find_all("div", class_="piece")
      for article in articles:
        # 提取標題
        title = article.find("h3", itemprop="headline").get_text(strip=True)
        
        # 提取網址
        link = article.find("a", itemprop="url")["href"]
        
        # 提取摘要
        description = article.find("p", class_="summary").get_text(strip=True)
        
        # 提取並格式化日期
        date_str = article.find("span", class_="date").get_text(strip=True)
        formatted_date = date_str.split()[0].replace("/", "")
        news_list.append({
          "公司名稱":company_name,
          "日期":formatted_date,
          "新聞標題":title,
          "新聞連結":link,
          "新聞內容":description,
          "新聞來源":"ETtoday永續雲"
        })
    return news_list
  except:
    if len(news_list)<1:
      return None
    else:
      return news_list
    

def Get_AllNews(company_name):
  result=catch_businesstoday(company_name,5)
  result+=catch_esgtimes(company_name)
  result+=catch_cna(company_name+" ESG")
  result+=catch_ESGView(company_name,5)
  return result
# if __name__ == "__main__":
  # company_name="聯發科"
  # print("="*50+"今周刊"+"="*50)
  # catch_businesstoday(company_name,5)
  # # print("="*50+"CSR天下")
  # # catch_CSR(company_name,5)
  # print("="*50+"esgtimes"+"="*50)
  # catch_esgtimes(company_name)#無法讀取更多文章
  # print("="*50+"中央社"+"="*50)
  # catch_cna("聯發科 ESG")
  # # print("="*50+"CSRone"+"="*50)
  # # catch_CSRone(company_name)
  # # print("="*50+"ESGTW"+"="*50)
  # # catch_ESGTW(company_name)
  # print("="*50+"ESG遠見"+"="*50)
  # catch_ESGView(company_name,5)
  # # print("="*50+"ETtoday"+"="*50)
  # # catch_ETtoday(company_name,5)