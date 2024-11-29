import pandas as pd
from sklearn.impute import KNNImputer

data=pd.read_excel(r"C:\Users\lin78\OneDrive\桌面\ESGBackEnd\各年炭排放資訊\111年碳排放資訊.xlsx",thousands=",")
# 定義函數來處理百分比數據轉換為小數
def convert_percent_to_decimal(value):
    try:
        if isinstance(value, str) and "%" in value:
            return float(value.strip('%')) / 100  # 將百分比轉為小數
    except ValueError:
        return value
    return value  # 保持非百分比數據不變

# 過濾數據並應用轉換
percent_columns = [col for col in data.columns if data[col].astype(str).str.contains('%').any()]
data[percent_columns] = data[percent_columns].map(convert_percent_to_decimal)

# 將數據中的百分比格式轉換為小數格式
#data = data.applymap(convert_percent_to_decimal)
#print(data["產業類別"].unique())

# 初始化 KNNImputer
imputer = KNNImputer(n_neighbors=2)
def fill_na(industry):
  data_industry=data[data["產業類別"]==industry]
    # 檢查是否有超過 3 列資料，否則不進行補植
  if len(data_industry) < 3:
    return data_industry
    # 選取公司資料
  #companyData=data_industry.iloc[:,:3]
  #print(companyData)

    # 選取所有資料
  #data_fillna=data_industry.iloc[:,3:-1]
  #print(data_fillna)
    # 提取數值欄位
  numeric_data = data_industry.select_dtypes(include=['number'])

    # 選取非數字資料
  non_numeric_data = data_industry.select_dtypes(exclude=['number'])
    # 檢查是否有超過 3 列資料，否則不進行補植
  if numeric_data.dropna(thresh=3).shape[0] < 3:
    return data_industry
  # 檢查並將所有值為空的欄位設為 0
  for column in numeric_data.columns:
      if numeric_data[column].isna().all():
          numeric_data[column] = 0
  #ESG_Score=data_industry.iloc[:,-1]
  # 使用 KNNImputer 補植缺失值
  numeric_data_imputed = imputer.fit_transform(numeric_data)

  # 確保輸出形狀與列名一致
  if numeric_data_imputed.shape[1] == len(numeric_data.columns):
    numeric_data_imputed = pd.DataFrame(numeric_data_imputed, columns=numeric_data.columns)
  else:
    print(f"警告：{industry} 產業數值欄位形狀不匹配！")
    return data_industry

  # 合併公司資料、補植後的欄位和 ESG 分數
  data_imputed = pd.concat([non_numeric_data.reset_index(drop=True),
                            numeric_data_imputed.reset_index(drop=True)],axis=1)

  #ESG_Score.reset_index(drop=True)], axis=1
  return data_imputed

dataframe=pd.DataFrame()
for i in data["產業類別"].unique():
  data_fill=fill_na(i)
  dataframe=pd.concat([dataframe,data_fill], axis=0)
dataframe.to_csv("111年Listed_Info_Emission_KNN補值.csv",encoding='utf-8-sig', index=False)