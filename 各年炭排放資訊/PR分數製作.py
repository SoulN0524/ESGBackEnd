import pandas as pd

data=pd.read_csv(r"C:\Users\lin78\OneDrive\桌面\ESGBackEnd\各年炭排放資訊\111年Listed_Info_Emission_KNN補值.csv",encoding='utf-8-sig')
data=data.dropna()
data=data[data["產業類別"]!="其他"]
#print(data.dtypes)
#IndustryData=data[data["產業類別"]=="半導體業"]
# #IndustryData
#Industry=IndustryData.sort_values(by="直接(範疇一)溫室氣體排放量(公噸CO₂e)")
# #Industry.index[0]
#Industry["公司名稱"].values
#for i in Industry.index:
  # print(Industry.at[i,"直接(範疇一)溫室氣體排放量(公噸CO₂e)"])

Compare_Columns=data.iloc[:, 3:-4].columns
#print(Compare_Columns)
#data["產業類別"].unique()
company_index=0 #公司索引

def Create_ESG(industry):
  """
  Create ESG (Environmental, Social, and Governance) scores for a given industry.
  This function calculates and assigns PR (Percentile Rank) scores to various columns
  in the dataset for companies within a specified industry. The PR scores are calculated
  based on the ranking of companies within the industry for each column.
  Parameters:
  industry (str): The industry category for which the ESG scores are to be created.
  Returns:
  None
  Notes:
  - The function modifies the global 'data' DataFrame in place.
  - The columns to be compared are specified in the 'Compare_Columns' list.
  - For certain columns, higher values are considered better, while for others, lower values are considered better.
  - The PR scores are calculated such that the highest-ranked company gets the highest PR score (99), and the lowest-ranked company gets the lowest PR score (0).
  - If multiple companies have the same value for a column, they receive the same PR score.
  Example:
  Create_ESG("Technology")
  """
  IndustryData = data.copy()  # Copy the original data to avoid modifying it
  IndustryData=IndustryData[IndustryData["產業類別"]==industry]
  PR_Violatility=round(99/len(IndustryData))
  for col in Compare_Columns:

    #print(col+":")
    if col=="再生能源使用率" or col=="員工福利平均數(仟元/人)" or col=="員工薪資平均數(仟元/人)" or col=="非擔任主管職務之全時員工薪資平均數(仟元/人)" or col=="非擔任主管職務之全時員工薪資中位數(仟元/人)"  or col=="董事會席次(席)" or col=="獨立董事席次(席)"  or col=="董事出席董事會出席率" or col=="公司年度召開法說會次數(次)":
      Industry=IndustryData.sort_values(by=col,ascending=False)
      PR=99+PR_Violatility
      #print(Industry)
      for i in Industry.index:
        if i==Industry.index[0]:
          #print(i)
          PR=PR-PR_Violatility
          company_index=i
          data.at[i,col]=PR
          continue
        if data.at[i,col]==data.at[company_index,col]:

          data.at[i,col]=PR
          company_index=i
          continue
        else:
          if PR-PR_Violatility<0:
            PR=0
          else:
            PR=PR-PR_Violatility
          data.at[i,col]=PR
          company_index=i
    elif col=="女性董事比率" or col=="管理職女性主管占比":
      # 計算距離欄位
      try:
        IndustryData[f"{col}_distance"] = abs(data[col] - 0.5)
        #print(f"成功生成距離欄位 {col}_distance")
      except Exception as e:
        #print(f"生成距離欄位 {col}_distance 時發生錯誤: {e}")
        continue
      #print(IndustryData.columns)  # 檢查 DataFrame 的所有欄位名稱
      #print(IndustryData[[col, f"{col}_distance"]].head())  # 檢查新生成的欄位
      if f"{col}_distance" not in IndustryData.columns:
        #print(f"警告: 欄位 {col}_distance 不存在，跳過此部分")
        continue
      Industry=IndustryData.sort_values(by=f"{col}_distance",ascending=True)
      PR=99+PR_Violatility
      for i in Industry.index:
        if i==Industry.index[0]:
          PR=PR-PR_Violatility
          company_index=i
          data.at[i,col]=PR
          continue
        if IndustryData.at[i,f"{col}_distance"]==IndustryData.at[company_index,f"{col}_distance"]:
          data.at[i,col]=PR
          company_index=i
          continue
        else:
          if PR-PR_Violatility<0:
            PR=0
          else:
            PR=PR-PR_Violatility
          data.at[i,col]=PR
          company_index=i
      try:
        IndustryData.drop(columns=[f"{col}_distance"], inplace=True)
        #print(f"已刪除欄位: {col}_distance")
      except KeyError:
        print(f"警告: 欄位 {col}_distance 不存在，無法刪除")
    
    else:
      Industry=IndustryData.sort_values(by=col)
      PR=99+PR_Violatility
      for i in Industry.index:
        if i==Industry.index[0]:
          PR=PR-PR_Violatility
          company_index=i
          data.at[i,col]=PR
          continue
        if data.at[i,col]==data.at[company_index,col]:
          data.at[i,col]=PR
          company_index=i
          continue
        else:
          if PR-PR_Violatility<0:
            PR=0
          else:
            PR=PR-PR_Violatility
          data.at[i,col]=PR
          company_index=i

for i in data["產業類別"].unique():
  ESG=Create_ESG(i)
data["ESG_ByPR"]=data[Compare_Columns].mean(axis=1).round().astype(int)
data.to_csv("111年碳排放PR處理.csv",encoding='utf-8-sig')