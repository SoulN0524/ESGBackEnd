import pandas as pd

data=pd.read_csv(r"C:\Users\lin78\OneDrive\桌面\ESGBackEnd\各年炭排放資訊\112年碳排放PR處理.csv",encoding='utf-8-sig')

E_column=["用水量(公噸)","使用率(再生能源/總能源)" , "用水密集度-密集度(公噸/單位)" ,"範疇一-排放量(噸CO2e)" ,"範疇二-排放量(噸CO2e)" ,"範疇三-排放量(噸CO2e)" ,"溫室氣體排放密集度-密集度(噸CO2e/單位)","有害廢棄物量-數據(公噸)" ,"非有害廢棄物量-數據(公噸)","總重量(有害+非有害)-數據(公噸)" ,"廢棄物密集度-密集度(公噸/單位)" ]

E_column112=["用水量(公噸(t))","再生能源使用率" , "用水密集度" ,"直接(範疇一)溫室氣體排放量(公噸CO₂e)" ,
             "能源間接(範疇二)溫室氣體排放量(公噸CO₂e)" ,"其他間接(範疇三)溫室氣體排放量(公噸CO₂e)" ,
             "溫室氣體排放密集度","有害廢棄物(公噸(t))" 
             ,"非有害廢棄物(公噸(t))","總重量(有害+非有害)(公噸(t))" ,"廢棄物密集度" ]

S_column112=["員工福利平均數(仟元/人)" ,"員工薪資平均數(仟元/人)" ,"非擔任主管職務之全時員工薪資平均數(仟元/人)" 
             ,"非擔任主管職務之全時員工薪資中位數(仟元/人)" ,"職業災害人數","職業災害比率","火災件數","火災死傷人數","火災比率(死傷人數/員工總人數)"]

S_column=["員工福利平均數(仟元/人)" ,"員工薪資平均數(仟元/人)" ,"非擔任主管職務之全時員工薪資平均數(仟元/人)" ,"非擔任主管之全時員工薪資中位數(仟元/人)" ,"職業災害人數及比率-人數","職業災害人數及比率-比率"]

G_column112=["董事會席次(席)" ,"管理職女性主管占比" ,"獨立董事席次(席)" ,"女性董事比率",
             "董事出席董事會出席率","公司年度召開法說會次數(次)"]

G_column=["董事席次(含獨立董事)(席)" ,"管理職女性主管占比" ,"獨立董事席次(席)" ,"女性董事席次及比率-比率","董事出席董事會出席率","公司年度召開法說會次數(次)"]
for i in data["產業類別"].unique():
    industry_data = data[data["產業類別"] == i]
    #data["ESG_ByPR"]=data[["E","S","G"]].mean(axis=1).round().astype(int)
    data.loc[data["產業類別"] == i,"E"]=industry_data[E_column112].mean(axis=1).round().astype(int)
    
    data.loc[data["產業類別"] == i,"S"]=industry_data[S_column112].mean(axis=1).round().astype(int)
    
    data.loc[data["產業類別"] == i,"G"]=industry_data[G_column112].mean(axis=1).round().astype(int)
data.to_csv("112年碳排放PR處理_最終版.csv",encoding='utf-8-sig')