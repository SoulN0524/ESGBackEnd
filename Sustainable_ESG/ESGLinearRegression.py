import pandas as pd
import numpy as np 
import pyodbc
import pymysql
import pymssql
import lightgbm as lgb
from sqlalchemy import create_engine
from sklearn.model_selection import train_test_split,cross_val_score
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor,GradientBoostingRegressor
from sklearn.model_selection import RandomizedSearchCV,GridSearchCV
from sklearn.preprocessing import MinMaxScaler
from sklearn.feature_selection import RFE
from sklearn.metrics import r2_score
import os 
import matplotlib.pyplot as plt
from sklearn.metrics import r2_score,mean_squared_error
#https://ycjhuo.gitlab.io/blogs/Python-Pandas-Download-Data-From-MSSQL.html#%E7%94%A8-pandas-%E8%BD%89%E5%AD%98%E8%B3%87%E6%96%99%E7%82%BA-dataframe-%E4%B8%A6%E8%BC%B8%E5%87%BA%E6%88%90-excel
conn=pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LAPTOP-JRA4GUH1\\MSSQLSERVER2022;DATABASE=ESGInfoHub;UID=Yogurt;PWD=Stock0050')


query="select * from [Emission112Data]"
data=pd.read_sql(query,conn)
data.to_csv("ESG.csv",encoding='utf-8-sig')
print(data)
# drop_indics=data[(data["直接(範疇一)溫室氣體排放量(公噸)"]==0.0000) |
#                  (data["能源間接(範疇二)溫室氣體排放量(公噸)"]==0.0000)|
#                  (data["其他間接(範疇三)溫室氣體排放量(公噸)"]==0.0000) | 
#                  (data["溫室氣體排放密集度"]==0.0000) |
#                  (data["再生能源使用率"]==0.0000) |
#                  (data["用水量(公噸(t))"]==0.0000) |
#                  (data["用水密集度"]==0.0000)|
#                  (data["總重量(有害+非有害)(公噸(t))"]==0.0000)|
#                  (data["廢棄物密集度(公噸)"]==0.0000)|
#                  #(data["員工福利平均數(仟元/人)"]==0) |
#                  #(data["員工薪資平均數(仟元/人)"]==0) |
#                  #(data["非擔任主管職務之全時員工薪資平均數(仟元/人)"]==0)|
#                  #(data["非擔任主管職務之全時員工薪資中位數(仟元/人)"]==0)|
#                  (data["管理職女性主管占比"]==0.0000)|
#                  (data["職業災害比率"]==0.0000) |
#                  (data["董事會席次(席)"]==0) |
#                  #(data["獨立董事席次(席)"]==0)|
#                  #(data["女性董事比例"]==0.0000)|
#                  #(data["董事出席董事會出席率"]==0.0000) |
#                  (data["公司年度召開法說會次數(次)"]==0) |
#                  (data["ESG"]==0)].index
# data=data.drop(drop_indics)
# data.to_csv('Adj_Data.csv',encoding='utf-8-sig')
# data['用水量(公噸(t))'] = np.log1p(data['用水量(公噸(t))'])
# data['直接(範疇一)溫室氣體排放量(公噸)']=np.log1p(data['直接(範疇一)溫室氣體排放量(公噸)'])
# data['能源間接(範疇二)溫室氣體排放量(公噸)']=np.log1p(data['能源間接(範疇二)溫室氣體排放量(公噸)'])
# data['其他間接(範疇三)溫室氣體排放量(公噸)']=np.log1p(data['其他間接(範疇三)溫室氣體排放量(公噸)'])
# print(data)
# # 定義特徵名稱列表
# feature_names = [
#     "直接(範疇一)溫室氣體排放量(公噸)",
#     "能源間接(範疇二)溫室氣體排放量(公噸)",
#     "其他間接(範疇三)溫室氣體排放量(公噸)",
#     "溫室氣體排放密集度",
#     "再生能源使用率",
#     "用水量(公噸(t))",
#     "用水密集度",
#     "總重量(有害+非有害)(公噸(t))",
#     "廢棄物密集度(公噸)",
#     #"員工福利平均數(仟元/人)",
#     #"員工薪資平均數(仟元/人)",
#     #"非擔任主管職務之全時員工薪資平均數(仟元/人)",
#     #"非擔任主管職務之全時員工薪資中位數(仟元/人)",
#     "管理職女性主管占比",
#     "職業災害比率",
#     "董事會席次(席)",
#     #"獨立董事席次(席)",
#     #"女性董事比例",
#     #"董事出席董事會出席率",
#     "公司年度召開法說會次數(次)"
# ]


# # 設置 X 變量
# X = np.array(data[feature_names].values)
# y = data.iloc[:,-4].values

# print(X)
# print(y)

# X_train,X_test,y_train,y_test=train_test_split(X,y,test_size=0.35,random_state=0)

    
# # 使用隨機森林做特徵選擇
# model = RandomForestRegressor(
#     n_estimators=60, 
#     random_state=0,
#     max_depth=5,
#     min_samples_split=10,
#     min_samples_leaf=5,
#     max_features='sqrt'
# )

# # 使用 RFE 篩選特徵74

# rfe = RFE(estimator=model, n_features_to_select=7)
# rfe.fit(X_train, y_train)

# # 被選中的特徵
# selected_features = np.array(feature_names)[rfe.support_]
# print("Selected features:", selected_features)

# # 使用選出的特徵重新訓練模型
# X_train_rfe = rfe.transform(X_train)
# X_test_rfe = rfe.transform(X_test)

# # 訓練並評估模型
# def RandomForestRegressionModel(X_train, X_test, y_train, y_test):
#     model.fit(X_train, y_train)
#     y_pred=model.predict(X_test)
#     # 預測
#     score = r2_score(y_test, y_pred)
#     print(" R^2:", score)
   

#     # 使用 5 折交叉驗證
#     scores = cross_val_score(model, X_train, y_train, cv=10, scoring='r2')
#     print("Cross-Validation R^2 Scores:", scores)
#     print("Mean R^2 Score:", scores.mean())

#     # # 輸出特徵重要性
#     # feature_importances = pd.DataFrame({
#     #     'feature': selected_features,
#     #     'importance': model.feature_importances_
#     # }).sort_values(by='importance', ascending=False)
#     # print(feature_importances)

#     # 定義隨機搜索參數
#     param_dist = {
#         'n_estimators': [40, 50, 80],
#         'max_depth': [3, 5, 7],
#         'min_samples_split': [2, 5, 8],
#         'min_samples_leaf':[4,6,8]
#     }

#     # 使用隨機森林進行隨機搜索
#     random_search = RandomizedSearchCV(model, param_dist, n_iter=10, cv=5)
#     random_search.fit(X_train, y_train)

#     # 最佳參數組合
#     print("隨機森林搜索法:",random_search.best_params_)


#     grid_search = GridSearchCV(model, param_dist, cv=5)
#     grid_search.fit(X, y)

#     # 最佳參數組合
#     print("網格搜尋法:",grid_search.best_params_)

# # 呼叫函數時傳入篩選過後的特徵
# RandomForestRegressionModel(X_train_rfe, X_test_rfe, y_train, y_test)


# # 梯度提升回歸模型
# def GradientBoostingRegressionModel(X_train, X_test, y_train, y_test, feature_names):
#     model = GradientBoostingRegressor(
#         n_estimators=100,   # 樹的數量
#         learning_rate=0.1,  # 學習率
#         max_depth=4,        # 每棵樹的最大深度
#         min_samples_split=5, 
#         min_samples_leaf=4,
#         random_state=0
#     )
#     model.fit(X_train, y_train)  # 訓練模型

#     # 預測
#     train_score = model.score(X_train, y_train)
#     test_score = model.score(X_test, y_test)
#     print("Train R^2:", train_score)
#     print("Test R^2:", test_score)

#     # 使用 5 折交叉驗證
#     scores = cross_val_score(model, X_train, y_train, cv=5, scoring='r2')
#     print("Cross-Validation R^2 Scores:", scores)
#     print("Mean R^2 Score:", scores.mean())

#     # 輸出特徵重要性
#     feature_importances = pd.DataFrame({
#         'feature': feature_names,
#         'importance': model.feature_importances_
#     }).sort_values(by='importance', ascending=False)

#     print(feature_importances)

# # 呼叫梯度提升回歸模型
# GradientBoostingRegressionModel(X_train, X_test, y_train, y_test, feature_names)