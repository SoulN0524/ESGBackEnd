import pyodbc
import pandas as pd
import numpy as np
import pymysql
def Insert_Listed_Info():
    ListedInfo=pd.read_excel("C:\\Users\\user\\OneDrive\\桌面\\專題\\Sustainable_ESG\\112年ListedInfo.xlsx",thousands=',')
    ListedInfo["股票代號"]=ListedInfo["股票代號"].astype(str)
    ListedInfo["統一編號"]=ListedInfo['統一編號'].astype(str)
    ListedInfo["資本額"]=ListedInfo["資本額"].astype(float)
    ListedInfo["入選投組"]=ListedInfo["入選投組"].str.replace("'","").replace("無","").astype(str)
    ListedInfo["經營業務"]=ListedInfo["經營業務"].str.replace("\n","").replace(" ","").astype(str)
    ListedInfoData=ListedInfo.values

    try:
        conn=pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LAPTOP-JRA4GUH1\\MSSQLSERVER2022;DATABASE=ESGInfoHub;UID=Yogurt;PWD=Stock0050')
        cur=conn.cursor()
        for i in range(ListedInfo.shape[0]):
            insertSql = f"""
                            INSERT INTO ListedInfo(股票代號,公司名稱,產業類別,統一編號,董事長,總經理,資本額,經營業務,入選投組)
                            VALUES ('{ListedInfoData[i,0]}', '{ListedInfoData[i,1]}', '{ListedInfoData[i,2]}', '{ListedInfoData[i,3]}',
                                    '{ListedInfoData[i,4]}', '{ListedInfoData[i,5]}', {ListedInfoData[i,6]}, '{ListedInfoData[i,7]}','{ListedInfoData[i,8]}')
                        """
            cur.execute(insertSql)
            cur.connection.commit()
        conn.commit()
        cur.close()

        conn.close()
        print("Listed Info has been inserted to sql")
    except pymysql.err.OperationalError as e:
        print(f"連結資料庫出錯:{e}")


def Insert_110_Emission():
    data=pd.read_csv("C:\\Users\\user\\OneDrive\\桌面\\專題\\Sustainable_ESG\\110年Listed_info_emission.csv",encoding='utf-8-sig')

    # Step 2: 將 "無" 替換為 NaN
    data["股票代號"]=data["股票代號"].astype(str)
    data["直接溫室氣體排放量(範疇一)(噸CO2e)"] = data["直接溫室氣體排放量(範疇一)(噸CO2e)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["直接溫室氣體排放量(範疇一)(噸CO2e)"]=data["直接溫室氣體排放量(範疇一)(噸CO2e)"].str.replace(",","").astype(float)

    data["能源間接(範疇二)(噸CO2e)"]=data["能源間接(範疇二)(噸CO2e)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["能源間接(範疇二)(噸CO2e)"]=data["能源間接(範疇二)(噸CO2e)"].str.replace(",","").astype(float)

    data["其他間接(範疇三)(噸CO2e)"]=data["其他間接(範疇三)(噸CO2e)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["其他間接(範疇三)(噸CO2e)"]=data["其他間接(範疇三)(噸CO2e)"].str.replace(",","").astype(float)

    data["溫室氣體排放密集度(噸CO2e/熟料產量)"]=data["溫室氣體排放密集度(噸CO2e/熟料產量)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["溫室氣體排放密集度(噸CO2e/熟料產量)"]=data["溫室氣體排放密集度(噸CO2e/熟料產量)"].str.replace(",","").astype(float)

    data["再生能源使用率"]=data["再生能源使用率"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["再生能源使用率"]=data["再生能源使用率"].str.replace("%","").replace(",","").astype(float)/100


    data["用水量(公噸)"]=data["用水量(公噸)"].apply(lambda x: np.nan if str(x).startswith("無") else x)
    data["用水量(公噸)"]=data["用水量(公噸)"].str.replace(",","").astype(float)

    data["用水密集度(公噸)"]=data["用水密集度(公噸)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["用水密集度(公噸)"]=data["用水密集度(公噸)"].str.replace(",","").astype(float)

    data["有害廢棄物(公噸)"]=data["有害廢棄物(公噸)"].apply(lambda x: np.nan if str(x).startswith("無") else x)
    data["有害廢棄物(公噸)"]=data["有害廢棄物(公噸)"].str.replace(",","").astype(float)

    data["非有害廢棄物(公噸)"]=data["非有害廢棄物(公噸)"].apply(lambda x: np.nan if str(x).startswith("無") else x)
    data["非有害廢棄物(公噸)"]=data["非有害廢棄物(公噸)"].str.replace(",","").astype(float)

    data["總重量(有害+非有害)(公噸)"]=data["總重量(有害+非有害)(公噸)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["總重量(有害+非有害)(公噸)"]=data["總重量(有害+非有害)(公噸)"].str.replace(",","").astype(float)

    data["廢棄物密集度(公噸)"]=data["廢棄物密集度(公噸)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["廢棄物密集度(公噸)"]=data["廢棄物密集度(公噸)"].str.replace(",","").astype(float)


    data["員工福利平均數(仟元/人)"]=data["員工福利平均數(仟元/人)"].str.replace(",","").astype(float)


    data["員工薪資平均數(仟元/人)"]=data["員工薪資平均數(仟元/人)"].str.replace(",","").astype(float)

    data["非擔任主管職務之全時員工薪資平均數(仟元/人)"]=data["非擔任主管職務之全時員工薪資平均數(仟元/人)"].str.replace(",","").astype(float)

    data["非擔任主管職務之全時員工薪資中位數(仟元/人)"]=data["非擔任主管職務之全時員工薪資中位數(仟元/人)"].str.replace(",","").astype(float)

    data["管理職女性主管占比"]=data["管理職女性主管占比"].str.replace("%","").replace(",","").astype(float)/100

    data["職業災害人數"]=data["職業災害人數"].str.replace("人","").replace(",","").astype(float)
    data["職業災害人數比率(職災人數/總人數)"]=data["職業災害人數比率(職災人數/總人數)"].str.replace("%","").replace(",","").astype(float)/100
    data["董事會席次(席)"]=data["董事會席次(席)"].astype(float)
    data["獨立董事席次(席)"]=data["獨立董事席次(席)"].astype(float)
    data["女性董事比率"]=data["女性董事比率"].str.replace("%","").replace(",","").astype(float)/100
    data["董事出席董事會出席率"]=data["董事出席董事會出席率"].str.replace("%","").replace(",","").astype(float)/100
    data["公司年度召開法說會次數(次)"]=data["公司年度召開法說會次數(次)"].astype(float)
    data["男性比例"]=data["男性比例"].astype(float)
    data["女性比例"]=data["女性比例"].astype(float)
    data["員工流動率"]=data["員工流動率"].astype(float)
    data["E"]=data["E"].astype(float)
    data["S"]=data["S"].astype(float)
    data["G"]=data["G"].astype(float)
    data["ESG"]=data["ESG"].astype(float)
    data=data.replace({np.nan: 0})

    print(data.dtypes)

    data=data.values

    try:
        conn=pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LAPTOP-JRA4GUH1\\MSSQLSERVER2022;DATABASE=ESGInfoHub;UID=Yogurt;PWD=Stock0050')
        cur=conn.cursor()
        for i in range(data.shape[0]):
            #欄位有換過
            inserSql=f"""
                        INSERT INTO Emission110Data (STOCK, STOCKNM, [DEMISSION(TON)], [INDEMISSION(TON)],[OINDEMISSION(TON)], GINTENSITY, [RENERGY(%)],[WCONSUMPTION(TON)],
                                                    [WINTENSITY(TON)],[HWASTE(TON)],[UHWASTE(TON)],[TOTALWASTE(TON)],[WASTEINTENSITY(TON)],AVGEMPLOYEEBENEFIT,AVGEMPLOYEESALARY,
                                                    AVGNONEXECUTIVESALARY,MIDNONEXECUTIVESALARY,FEMALEEXECUTIVEPERSENT,OACCIDENT,POACCIDENT,NUMPATTENDANCE,NUMIPATTENDANCE,PFPATTENDANCE,
                                                    PPATTENDANCE,ICONFERENCE,MALE,FEMALE,EMPLOYEEFLOW,E,S,G,ESG)
                                VALUES ('{data[i,0]}', '{data[i,1]}', {data[i,3]}, {data[i,4]},
                                        {data[i,5]}, {data[i,6]}, {data[i,8]}, {data[i,11]},{data[i,12]},{data[i,14]},{data[i,15]},{data[i,16]},{data[i,17]},
                                        {data[i,19]},{data[i,20]},{data[i,21]},{data[i,22]},{data[i,23]},{data[i,24]},{data[i,25]},{data[i,26]},{data[i,27]},
                                        {data[i,28]},{data[i,29]},{data[i,31]},{data[i,32]},{data[i,33]},{data[i,34]},{data[i,35]},{data[i,36]},{data[i,37]},{data[i,38]})
                      """
            cur.execute(inserSql)
            cur.connection.commit()
        conn.commit()
        cur.close()

        conn.close()
        print("110年EMISSION has been inserted to sql")
    except pymysql.err.OperationalError as e:
        print(f"連結資料庫出錯:{e}")
def Insert_111_Emission():
    data=pd.read_csv("C:\\Users\\user\\OneDrive\\桌面\\專題\\Sustainable_ESG\\111年Listed_info_emission.csv",encoding='utf-8-sig')

    # Step 2: 將 "無" 替換為 NaN
    data["股票代號"]=data["股票代號"].astype(str)
    data["直接溫室氣體排放量(範疇一)(噸CO2e)"] = data["直接溫室氣體排放量(範疇一)(噸CO2e)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["直接溫室氣體排放量(範疇一)(噸CO2e)"]=data["直接溫室氣體排放量(範疇一)(噸CO2e)"].str.replace(",","").astype(float)

    data["能源間接(範疇二)(噸CO2e)"]=data["能源間接(範疇二)(噸CO2e)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["能源間接(範疇二)(噸CO2e)"]=data["能源間接(範疇二)(噸CO2e)"].str.replace(",","").astype(float)

    data["其他間接(範疇三)(噸CO2e)"]=data["其他間接(範疇三)(噸CO2e)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["其他間接(範疇三)(噸CO2e)"]=data["其他間接(範疇三)(噸CO2e)"].str.replace(",","").astype(float)

    data["溫室氣體排放密集度(噸CO2e/熟料產量)"]=data["溫室氣體排放密集度(噸CO2e/熟料產量)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["溫室氣體排放密集度(噸CO2e/熟料產量)"]=data["溫室氣體排放密集度(噸CO2e/熟料產量)"].str.replace(",","").astype(float)

    data["再生能源使用率"]=data["再生能源使用率"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["再生能源使用率"]=data["再生能源使用率"].str.replace("%","").replace(",","").astype(float)/100


    data["用水量(公噸)"]=data["用水量(公噸)"].apply(lambda x: np.nan if str(x).startswith("無") else x)
    data["用水量(公噸)"]=data["用水量(公噸)"].str.replace(",","").astype(float)

    data["用水密集度(公噸)"]=data["用水密集度(公噸)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["用水密集度(公噸)"]=data["用水密集度(公噸)"].str.replace(",","").astype(float)

    data["有害廢棄物(公噸)"]=data["有害廢棄物(公噸)"].apply(lambda x: np.nan if str(x).startswith("無") else x)
    data["有害廢棄物(公噸)"]=data["有害廢棄物(公噸)"].str.replace(",","").astype(float)

    data["非有害廢棄物(公噸)"]=data["非有害廢棄物(公噸)"].apply(lambda x: np.nan if str(x).startswith("無") else x)
    data["非有害廢棄物(公噸)"]=data["非有害廢棄物(公噸)"].str.replace(",","").astype(float)

    data["總重量(有害+非有害)(公噸)"]=data["總重量(有害+非有害)(公噸)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["總重量(有害+非有害)(公噸)"]=data["總重量(有害+非有害)(公噸)"].str.replace(",","").astype(float)

    data["廢棄物密集度(公噸)"]=data["廢棄物密集度(公噸)"].apply(lambda x: np.nan if x.startswith("無") else x)
    data["廢棄物密集度(公噸)"]=data["廢棄物密集度(公噸)"].str.replace(",","").astype(float)


    data["員工福利平均數(仟元/人)"]=data["員工福利平均數(仟元/人)"].str.replace(",","").astype(float)


    data["員工薪資平均數(仟元/人)"]=data["員工薪資平均數(仟元/人)"].str.replace(",","").astype(float)

    data["非擔任主管職務之全時員工薪資平均數(仟元/人)"]=data["非擔任主管職務之全時員工薪資平均數(仟元/人)"].str.replace(",","").astype(float)

    data["非擔任主管職務之全時員工薪資中位數(仟元/人)"]=data["非擔任主管職務之全時員工薪資中位數(仟元/人)"].str.replace(",","").astype(float)

    data["管理職女性主管占比"]=data["管理職女性主管占比"].str.replace("%","").replace(",","").astype(float)/100

    data["職業災害人數"]=data["職業災害人數"].str.replace("人","").replace(",","").astype(float)
    data["職業災害人數比率(職災人數/總人數)"]=data["職業災害人數比率(職災人數/總人數)"].str.replace("%","").replace(",","").astype(float)/100
    data["董事會席次(席)"]=data["董事會席次(席)"].astype(float)
    data["獨立董事席次(席)"]=data["獨立董事席次(席)"].astype(float)
    data["女性董事比率"]=data["女性董事比率"].str.replace("%","").replace(",","").astype(float)/100
    data["董事出席董事會出席率"]=data["董事出席董事會出席率"].str.replace("%","").replace(",","").astype(float)/100
    data["公司年度召開法說會次數(次)"]=data["公司年度召開法說會次數(次)"].astype(float)
    data["男性比例"]=data["男性比例"].astype(float)
    data["女性比例"]=data["女性比例"].astype(float)
    data["員工流動率"]=data["員工流動率"].astype(float)
    data["E"]=data["E"].astype(float)
    data["S"]=data["S"].astype(float)
    data["G"]=data["G"].astype(float)
    data["ESG"]=data["ESG"].astype(float)
    data=data.replace({np.nan: 0})

    print(data.dtypes)

    data=data.values

    try:
        conn=pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LAPTOP-JRA4GUH1\\MSSQLSERVER2022;DATABASE=ESGInfoHub;UID=Yogurt;PWD=Stock0050')
        cur=conn.cursor()
        for i in range(data.shape[0]):
            inserSql=f"""
                        INSERT INTO Emission111Data (STOCK, STOCKNM, [DEMISSION(TON)], [INDEMISSION(TON)],[OINDEMISSION(TON)], GINTENSITY, [RENERGY(%)],[WCONSUMPTION(TON)],
                                                    [WINTENSITY(TON)],[HWASTE(TON)],[UHWASTE(TON)],[TOTALWASTE(TON)],[WASTEINTENSITY(TON)],AVGEMPLOYEEBENEFIT,AVGEMPLOYEESALARY,
                                                    AVGNONEXECUTIVESALARY,MIDNONEXECUTIVESALARY,FEMALEEXECUTIVEPERSENT,OACCIDENT,POACCIDENT,NUMPATTENDANCE,NUMIPATTENDANCE,PFPATTENDANCE,
                                                    PPATTENDANCE,ICONFERENCE,MALE,FEMALE,EMPLOYEEFLOW,E,S,G,ESG)
                                VALUES ('{data[i,0]}', '{data[i,1]}', {data[i,3]}, {data[i,4]},
                                        {data[i,5]}, {data[i,6]}, {data[i,8]}, {data[i,11]},{data[i,12]},{data[i,14]},{data[i,15]},{data[i,16]},{data[i,17]},
                                        {data[i,19]},{data[i,20]},{data[i,21]},{data[i,22]},{data[i,23]},{data[i,24]},{data[i,25]},{data[i,26]},{data[i,27]},
                                        {data[i,28]},{data[i,29]},{data[i,31]},{data[i,32]},{data[i,33]},{data[i,34]},{data[i,35]},{data[i,36]},{data[i,37]},{data[i,38]})
                      """
            cur.execute(inserSql)
            cur.connection.commit()
        conn.commit()
        cur.close()

        conn.close()
        print("111年EMISSION has been inserted to sql")
    except pymysql.err.OperationalError as e:
        print(f"連結資料庫出錯:{e}")


def Insert_112_Emission():
    data=pd.read_csv(r"C:\Users\user\Desktop\ESGBackEnd\112年Emission_兆豐ESG.csv",encoding='utf-8-sig',thousands=",")
    print(data)
     # Step 2: 將 "無" 替換為 NaN
    data["股票代號"]=data["股票代號"].astype(str)
    data["直接(範疇一)溫室氣體排放量(公噸CO₂e)"]=data["直接(範疇一)溫室氣體排放量(公噸CO₂e)"].astype(float)

    data["能源間接(範疇二)溫室氣體排放量(公噸CO₂e)"]=data["能源間接(範疇二)溫室氣體排放量(公噸CO₂e)"].astype(float)

    data["其他間接(範疇三)溫室氣體排放量(公噸CO₂e)"]=data["其他間接(範疇三)溫室氣體排放量(公噸CO₂e)"].astype(float)

    data["溫室氣體排放密集度"]=data["溫室氣體排放密集度"].astype(float)
 
    data["再生能源使用率"]=data["再生能源使用率"].str.replace("%","").replace(",","").astype(float)/100

    data["用水量(公噸(t))"]=data["用水量(公噸(t))"].astype(float)

    data["用水密集度"]=data["用水密集度"].astype(float)

    data["有害廢棄物(公噸(t))"]=data["有害廢棄物(公噸(t))"].astype(float)

    data["非有害廢棄物(公噸(t))"]=data["非有害廢棄物(公噸(t))"].astype(float)

    data["總重量(有害+非有害)(公噸(t))"]=data["總重量(有害+非有害)(公噸(t))"].astype(float)

    data["廢棄物密集度"]=data["廢棄物密集度"].astype(float)

    data["員工福利平均數(仟元/人)"]=data["員工福利平均數(仟元/人)"].astype(float)

    data["員工薪資平均數(仟元/人)"]=data["員工薪資平均數(仟元/人)"].astype(float)

    data["非擔任主管職務之全時員工薪資平均數(仟元/人)"]=data["非擔任主管職務之全時員工薪資平均數(仟元/人)"].astype(float)

    data["非擔任主管職務之全時員工薪資中位數(仟元/人)"]=data["非擔任主管職務之全時員工薪資中位數(仟元/人)"].astype(float)

    data["管理職女性主管占比"]=data["管理職女性主管占比"].str.replace("%","").astype(float)/100

    data["職業災害人數"]=data["職業災害人數"].astype(float)
    data["職業災害比率"]=data["職業災害比率"].str.replace("%","").astype(float)/100

    data["火災件數"]=data["火災件數"].astype(float)
    data["火災死傷人數"]=data["火災死傷人數"].astype(float)
    data["火災比率(死傷人數/員工總人數)"]=data["火災比率(死傷人數/員工總人數)"].str.replace("%","").astype(float)
    data["董事會席次(席)"]=data["董事會席次(席)"].astype(float)
    data["獨立董事席次(席)"]=data["獨立董事席次(席)"].astype(float)
    data["女性董事比率"]=data["女性董事比率"].str.replace("%","").astype(float)/100
    data["董事出席董事會出席率"]=data["董事出席董事會出席率"].str.replace("%","").astype(float)/100
    data["公司年度召開法說會次數(次)"]=data["公司年度召開法說會次數(次)"].astype(float)
    data["ESG"]=data["ESG"].astype(float)
    data["E"]=data["E"].astype(float)
    data["S"]=data["S"].astype(float)
    data["G"]=data["G"].astype(float)
    data=data.replace({np.nan: 0.0000})
    try:
        conn=pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LAPTOP-JRA4GUH1\\MSSQLSERVER2022;DATABASE=ESGInfoHub;UID=Yogurt;PWD=Stock0050')
        cur=conn.cursor()
        for i in range(data.shape[0]):
            inserSql=f"""
    INSERT INTO Emission112Data (
        [股票代號], [公司名稱], [直接(範疇一)溫室氣體排放量(公噸)], [能源間接(範疇二)溫室氣體排放量(公噸)],
        [其他間接(範疇三)溫室氣體排放量(公噸)], [溫室氣體排放密集度], [再生能源使用率], [用水量(公噸(t))], [用水密集度],
        [有害廢棄物(公噸(t))], [非有害廢棄物(公噸(t))], [總重量(有害+非有害)(公噸(t))], [廢棄物密集度(公噸)],
        [員工福利平均數(仟元/人)], [員工薪資平均數(仟元/人)], [非擔任主管職務之全時員工薪資平均數(仟元/人)],
        [非擔任主管職務之全時員工薪資中位數(仟元/人)], [管理職女性主管占比], [職業災害人數], [職業災害比率],
        [火災件數], [火災死傷人數], [火災比率(死傷人數/員工總人數)], [董事會席次(席)], [獨立董事席次(席)], [女性董事比例],
        [董事出席董事會出席率], [公司年度召開法說會次數(次)], [ESG],[E],[S],[G])
     VALUES (
        '{data.iloc[i, 0]}', '{data.iloc[i, 1]}', {data.iloc[i, 2]}, {data.iloc[i, 3]}, {data.iloc[i, 4]},
        {data.iloc[i, 5]}, {data.iloc[i, 6]}, {data.iloc[i, 7]}, {data.iloc[i, 8]}, {data.iloc[i, 9]},
        {data.iloc[i, 10]}, {data.iloc[i, 11]}, {data.iloc[i, 12]}, {data.iloc[i, 13]}, {data.iloc[i, 14]},
        {data.iloc[i, 15]}, {data.iloc[i, 16]}, {data.iloc[i, 17]}, {data.iloc[i, 18]}, {data.iloc[i, 19]},
        {data.iloc[i, 20]}, {data.iloc[i, 21]}, {data.iloc[i, 22]}, {data.iloc[i, 23]}, {data.iloc[i, 24]},
        {data.iloc[i, 25]}, {data.iloc[i, 26]}, {data.iloc[i, 27]}, {data.iloc[i,28]},{data.iloc[i,29]},{data.iloc[i,30]},{data.iloc[i,31]})

                      """
            cur.execute(inserSql)
            cur.connection.commit()
        conn.commit()
        cur.close()

        conn.close()
        print("112年EMISSION has been inserted to sql")
    except pymysql.err.OperationalError as e:
        print(f"連結資料庫出錯:{e}")





def Insert_Securities ():
    data=pd.read_csv("C:\\Users\\user\\OneDrive\\桌面\\專題\\Sustainable_ESG\\SercuritiesViolations.csv",encoding='utf-8-sig')
    data["發函日期"]=data["發函日期"].str.replace("/","").astype(str)
    data["股票代號"]=data["股票代號"].astype(str)
    print(data)
    data=data.values
    try:
        conn=pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LAPTOP-JRA4GUH1\\MSSQLSERVER2022;DATABASE=ESGInfoHub;UID=Yogurt;PWD=Stock0050')
        cur=conn.cursor()
        for i in range(data.shape[0]):
            if len(data[i,3]) >=512:
                continue
            else:
                inserSql=f"""
                            INSERT INTO SecuritiesViolations (VDATE,STOCK,STOCKNM,REASON,LAW,SITUATION,MINISTRY)
                                    VALUES ('{data[i,0]}','{data[i,1]}','{data[i,2]}','{data[i,3]}','{data[i,4]}','{data[i,5]}','{data[i,6]}')
                        """
                cur.execute(inserSql)
                cur.connection.commit()
        conn.commit()
        cur.close()

        conn.close()
        print("Futute violation has been inserted to sql")
    except pymysql.err.OperationalError as e:
        print(f"連結資料庫出錯:{e}")

def Insert_financial_data():
    data=pd.read_excel(r"C:\Users\user\OneDrive\桌面\專題\Sustainable_ESG\上市公司ESG資訊\上市.xlsx",thousands=',')
    data["代號"]=data["代號"].astype(str)
    data["名稱"]=data["名稱"].astype(str)
    data["2023年度毛利增減(%)"]=data["2023年度毛利增減(%)"].astype(float)
    data["2022年度毛利增減(%)"]=data["2022年度毛利增減(%)"].astype(float)
    data["2021年度毛利增減(%)"]=data["2021年度毛利增減(%)"].astype(float)
    data["2020年度毛利增減(%)"]=data["2020年度毛利增減(%)"].astype(float)
    data["2019年度毛利增減(%)"]=data["2019年度毛利增減(%)"].astype(float)
    data["2023年度營益增減(%)"]=data["2023年度營益增減(%)"].astype(float)
    data["2022年度營益增減(%)"]=data["2022年度毛利增減(%)"].astype(float)
    data["2021年度營益增減(%)"]=data["2021年度營益增減(%)"].astype(float)
    data["2020年度營益增減(%)"]=data["2020年度營益增減(%)"].astype(float)
    data["2019年度營益增減(%)"]=data["2019年度營益增減(%)"].astype(float)
    data["2023年度淨利增減(%)"]=data["2023年度淨利增減(%)"].astype(float)
    data["2022年度淨利增減(%)"]=data["2022年度淨利增減(%)"].astype(float)
    data["2021年度淨利增減(%)"]=data["2021年度淨利增減(%)"].astype(float)
    data["2020年度淨利增減(%)"]=data["2020年度淨利增減(%)"].astype(float)
    data["2019年度淨利增減(%)"]=data["2019年度淨利增減(%)"].astype(float)
    data=data.replace({np.nan: 0.00})
    data=data.values
    try:
        conn=pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=LAPTOP-JRA4GUH1\\MSSQLSERVER2022;DATABASE=ESGInfoHub;UID=Yogurt;PWD=Stock0050')
        cur=conn.cursor()
        for i in range(data.shape[0]):
            inserSql=f"""
                        INSERT INTO 上市公司五年獲利三率成長率 
                        (股票代號, 公司名稱, 毛利率2023年度成長率, 毛利率2022年度成長率, 毛利率2021年度成長率, 毛利率2020年度成長率, 毛利率2019年度成長率, 
                        營益率2023年度成長率, 營益率2022年度成長率, 營益率2021年度成長率, 營益率2020年度成長率, 營益率2019年度成長率, 
                        淨利率2023年度成長率, 淨利率2022年度成長率, 淨利率2021年度成長率, 淨利率2020年度成長率, 淨利率2019年度成長率)
                        VALUES ('{data[i,0]}', '{data[i,1]}', {data[i,2]}, {data[i,3]}, {data[i,4]}, {data[i,5]}, {data[i,6]}, 
                                {data[i,7]}, {data[i,8]}, {data[i,9]}, {data[i,10]}, {data[i,11]}, 
                                {data[i,12]}, {data[i,13]}, {data[i,14]}, {data[i,15]}, {data[i,16]});
                    """
            cur.execute(inserSql)
            cur.connection.commit()
        conn.commit()
        cur.close()

        conn.close()
        print("Futute violation has been inserted to sql")
    except pymysql.err.OperationalError as e:
        print(f"連結資料庫出錯:{e}")
    



if __name__=='__main__':

    # Insert_Listed_Info()
    # Insert_110_Emission()
    # Insert_111_Emission()
    # Insert_Securities()
    Insert_112_Emission()
    # Insert_financial_data()