import pandas as pd
data=pd.read_csv(r"C:\Users\lin78\OneDrive\桌面\ESGBackEnd\各年炭排放資訊\112年Listed_Info_Emission_KNN補值.csv",encoding='utf-8-sig')
data=data.dropna()
data=data[data["產業類別"]!="其他"]
print(data.dtypes)

Compare_Columns=data.iloc[:, 3:-4].columns
Compare_Columns

company_index=0 #公司索引

def Create_ESG(industry):
  IndustryData=data[data["產業類別"]==industry]
  PR_Violatility=round(99/len(IndustryData))
  for col in Compare_Columns:
    print(col+":")
    if col=="再生能源使用率" or col=="員工福利平均數(仟元/人)" or col=="員工薪資平均數(仟元/人)" or col=="非擔任主管職務之全時員工薪資平均數(仟元/人)" or col=="非擔任主管職務之全時員工薪資中位數(仟元/人)" or col=="管理職女性主管占比" or col=="董事會席次(席)" or col=="獨立董事席次(席)" or col=="女性董事比率" or col=="董事出席董事會出席率" or col=="公司年度召開法說會次數(次)":
      Industry=IndustryData.sort_values(by=col,ascending=False)
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
data.to_csv("112年碳排放PR處理.csv",encoding='utf-8-sig')