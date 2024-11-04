from flask import Flask,jsonify,request
from flask_sqlalchemy import SQLAlchemy
import pyodbc
from flasgger import Swagger, swag_from
from flask_cors import CORS
from sqlalchemy.exc import IntegrityError
app=Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI']=r"mssql+pyodbc://Yogurt:Stock0050@LAPTOP-JRA4GUH1\MSSQLSERVER2022/ESGInfoHub?driver=ODBC+Driver+17+for+SQL+Server"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SWAGGER']={
    "title":"ESGInfoHub API",
    "description": "Retrieve stock information",
    "version": "1.0.2",
    "termsOfService": "",
    "hide_top_bar": True
}


CORS(app)
Swagger(app)
db = SQLAlchemy(app)
class ListedInfo(db.Model):
    __tablename__ = 'ListedInfo'  # 指定表名稱
    stock = db.Column("股票代號", db.String(6), primary_key=True)
    stocknm = db.Column("公司名稱", db.String(32))
    industry = db.Column("產業類別", db.String(32))
    taxid = db.Column("統一編號", db.String(10))
    president = db.Column("董事長", db.String(64))
    ceo = db.Column("總經理", db.String(64))
    capital = db.Column("資本額", db.Float)
    mission = db.Column("經營業務", db.String(512))
    portfolio = db.Column("入選投組", db.String(256))

    def __init__(self, stock, stocknm, industry, taxid, president, ceo, capital, mission, portfolio):
        self.stock = stock
        self.stocknm = stocknm
        self.industry = industry
        self.taxid = taxid
        self.president = president
        self.ceo = ceo
        self.capital = capital
        self.mission = mission
        self.portfolio = portfolio

class Emission112(db.Model):
    __tablename__="Emission112Data"
    stock=db.Column("股票代號",db.String(6),primary_key=True)
    stocknm=db.Column("公司名稱",db.String(32))
    directed_emission=db.Column("直接(範疇一)溫室氣體排放量(公噸)",db.Float)
    indirected_emission=db.Column("能源間接(範疇二)溫室氣體排放量(公噸)",db.Float)
    other_indirected_emission=db.Column("其他間接(範疇三)溫室氣體排放量(公噸)",db.Float)
    greenHouse_emission_intensity=db.Column("溫室氣體排放密集度",db.Float)
    renewable_energy=db.Column("再生能源使用率",db.Float)
    waterConsumption=db.Column("用水量(公噸(t))",db.Float)
    water_intensity=db.Column("用水密集度",db.Float)
    hazardous=db.Column("有害廢棄物(公噸(t))",db.Float)
    unhazardous=db.Column("非有害廢棄物(公噸(t))",db.Float)
    Total_waste=db.Column("總重量(有害+非有害)(公噸(t))",db.Float)
    waste_intensity=db.Column("廢棄物密集度(公噸)",db.Float)
    employee_avg_benefit=db.Column("員工福利平均數(仟元/人)",db.Float)
    employee_avg_salary=db.Column("員工薪資平均數(仟元/人)",db.Float)
    non_executive_avg_salary=db.Column("非擔任主管職務之全時員工薪資平均數(仟元/人)",db.Float)
    non_executive_med_salary=db.Column("非擔任主管職務之全時員工薪資中位數(仟元/人)",db.Float)
    proportion_of_female_executive=db.Column("管理職女性主管占比",db.Float)
    occupational_accident=db.Column("職業災害人數",db.Float)
    proportion_occupational_accident=db.Column("職業災害比率",db.Float)
    num_of_fire=db.Column("火災件數",db.Float)
    death_of_fire=db.Column("火災死傷人數",db.Float)
    proportion_death_of_fire=db.Column("火災比率(死傷人數/員工總人數)",db.Float)
    num_of_president=db.Column("董事會席次(席)",db.Float)
    num_of_indpresident=db.Column("獨立董事席次(席)",db.Float)
    proportion_of_female_president=db.Column("女性董事比例",db.Float)
    attence_of_president=db.Column("董事出席董事會出席率",db.Float)
    investor_conference=db.Column("公司年度召開法說會次數(次)",db.Float)
    ESG=db.Column("ESG",db.Float)
    def __init__(self,stock,stocknm, directed_emission,indirected_emission, other_indirected_emission,greenHouse_emission_intensity,renewable_energy, waterConsumption, water_intensity,
                 hazardous,unhazardous,Total_waste, waste_intensity,employee_avg_benefit,employee_avg_salary,non_executive_avg_salary,non_executive_med_salary,proportion_of_female_executive,
                 occupational_accident,proportion_occupational_accident,num_of_fire,death_of_fire,proportion_death_of_fire,num_of_president, num_of_indpresident, proportion_of_female_president,
                  attence_of_president, investor_conference,ESG):
        self.stock=stock
        self.stocknm=stocknm
        self.directed_emission=directed_emission
        self.indirected_emission=indirected_emission
        self.other_indirected_emission=other_indirected_emission
        self.greenHouse_emission_intensity=greenHouse_emission_intensity
        self.renewable_energy=renewable_energy
        self.waterConsumption=waterConsumption
        self.water_intensity=water_intensity
        self.hazardous=hazardous
        self.unhazardous=unhazardous
        self.Total_waste=Total_waste
        self.waste_intensity=waste_intensity
        self.employee_avg_benefit=employee_avg_benefit
        self.employee_avg_salary=employee_avg_salary
        self.non_executive_avg_salary=non_executive_avg_salary
        self.non_executive_med_salary=non_executive_med_salary
        self.proportion_of_female_executive=proportion_of_female_executive
        self.occupational_accident=occupational_accident
        self.proportion_occupational_accident=proportion_occupational_accident
        self.num_of_fire=num_of_fire
        self.death_of_fire=death_of_fire
        self.proportion_death_of_fire=proportion_death_of_fire
        self.num_of_president=num_of_president
        self.num_of_indpresident=num_of_indpresident
        self.proportion_of_female_president=proportion_of_female_president
        self.attence_of_president=attence_of_president
        self.investor_conference=investor_conference
        self.ESG=ESG

class FinancialGrowthRate(db.Model):
    __tablename__="上市公司五年獲利三率成長率"
    stock = db.Column("股票代號", db.String(6), primary_key=True)  # stock code
    stocknm = db.Column("公司名稱", db.String(32))  # company name
    gross_margin_2023 = db.Column("毛利率2023年度成長率(%)", db.Float)  # gross margin growth rate 2023
    gross_margin_2022 = db.Column("毛利率2022年度成長率(%)", db.Float)  # gross margin growth rate 2022
    gross_margin_2021 = db.Column("毛利率2021年度成長率(%)", db.Float)  # gross margin growth rate 2021
    gross_margin_2020 = db.Column("毛利率2020年度成長率(%)", db.Float)  # gross margin growth rate 2020
    gross_margin_2019 = db.Column("毛利率2019年度成長率(%)", db.Float)  # gross margin growth rate 2019
    operating_income_2023 = db.Column("營益率2023年度成長率(%)", db.Float)  # operating income growth rate 2023
    operating_income_2022 = db.Column("營益率2022年度成長率(%)", db.Float)  # operating income growth rate 2022
    operating_income_2021 = db.Column("營益率2021年度成長率(%)", db.Float)  # operating income growth rate 2021
    operating_income_2020 = db.Column("營益率2020年度成長率(%)", db.Float)  # operating income growth rate 2020
    operating_income_2019 = db.Column("營益率2019年度成長率(%)", db.Float)  # operating income growth rate 2019
    net_profit_2023 = db.Column("淨利率2023年度成長率(%)", db.Float)  # net profit growth rate 2023
    net_profit_2022 = db.Column("淨利率2022年度成長率(%)", db.Float)  # net profit growth rate 2022
    net_profit_2021 = db.Column("淨利率2021年度成長率(%)", db.Float)  # net profit growth rate 2021
    net_profit_2020 = db.Column("淨利率2020年度成長率(%)", db.Float)  # net profit growth rate 2020
    net_profit_2019 = db.Column("淨利率2019年度成長率(%)", db.Float)  # net profit growth rate 2019
    def __init__(self, stock, stocknm, gross_margin_2023, gross_margin_2022, gross_margin_2021, gross_margin_2020, gross_margin_2019, 
                 operating_income_2023, operating_income_2022, operating_income_2021, operating_income_2020, operating_income_2019, 
                 net_profit_2023, net_profit_2022, net_profit_2021, net_profit_2020, net_profit_2019):
        self.stock = stock
        self.stocknm = stocknm
        self.gross_margin_2023 = gross_margin_2023
        self.gross_margin_2022 = gross_margin_2022
        self.gross_margin_2021 = gross_margin_2021
        self.gross_margin_2020 = gross_margin_2020
        self.gross_margin_2019 = gross_margin_2019
        self.operating_income_2023 = operating_income_2023
        self.operating_income_2022 = operating_income_2022
        self.operating_income_2021 = operating_income_2021
        self.operating_income_2020 = operating_income_2020
        self.operating_income_2019 = operating_income_2019
        self.net_profit_2023 = net_profit_2023
        self.net_profit_2022 = net_profit_2022
        self.net_profit_2021 = net_profit_2021
        self.net_profit_2020 = net_profit_2020
        self.net_profit_2019 = net_profit_2019


@app.route("/get_listed_info", methods=["GET"])
@swag_from({
    'tags': ['All_STOCK_API'],
    'Method':'GET',
    'summary': 'Retrieve all listed company information',
    'responses': {
        200: {
            'description': 'List of all companies',
            'type': {
                'application/json': [{
                    '股票代號': {'type': 'string'},
                    '公司名稱': {'type': 'string'},
                    '產業類別': {'type': 'string'},
                    '董事長': {'type': 'string'},
                    '總經理': {'type': 'string'},
                    '資本額': {'type': 'float'},
                    '經營業務': {'type': 'string'},
                    '入選投組': {'type': 'string'}
                }]
            }
        }
    }
})
def get_listed_info():
    # 查詢所有 ListedInfo 資料
    listings = ListedInfo.query.all()
    # 將查詢結果轉換為 JSON 格式
    results = []
    for listing in listings:
        results.append({
            "股票代號": listing.stock,
            "公司名稱": listing.stocknm,
            "產業類別": listing.industry,
            "董事長": listing.president,
            "總經理": listing.ceo,
            "資本額": listing.capital,
            "經營業務": listing.mission,
            "入選投組": listing.portfolio
        })
    
    return jsonify(results)


@app.route('/StockInfo/<string:stock_code>',methods=["GET"])
@swag_from({
    'tags': ['StockInfo'],
    'summary': 'Retrieve information for a specific stock',
    'Method':'GET',
    'parameters': [
        {
            'name': 'stock_code',
            'in': 'path',
            'type': 'string',
            'required': True,
            'description': 'Stock code of the company (e.g., 2330 for TSMC)'
        }
    ],
    'produces': ['application/json'],
    'responses': {
        200: {
            'description': 'Company information based on stock code',
            'type': {
                'application/json': {
                    '股票代號': {'type': 'string'},
                    '公司名稱': {'type': 'string'},
                    '產業類別': {'type': 'string'},
                    '董事長': {'type': 'string'},
                    '總經理': {'type': 'string'},
                    '資本額': {'type': 'float'},
                    '經營業務': {'type': 'string'},
                    '入選投組': {'type': 'string'}
                }
            }
        },
        404: {
            'description': 'Stock not found'
        }
    }
})
def get_stock_info(stock_code):
    try:
        stock_info = ListedInfo.query.filter_by(stock=stock_code).first()
        if stock_info:
            return jsonify({
                "股票代號": stock_info.stock,
                "公司名稱": stock_info.stocknm,
                "產業類別": stock_info.industry,
                "董事長": stock_info.president,
                "總經理": stock_info.ceo,
                "資本額": stock_info.capital,
                "經營業務": stock_info.mission,
                "入選投組": stock_info.portfolio
            })
        else:
            return jsonify({"message": "Stock not found"}), 404
    except Exception as e:
        return jsonify({"message": f"Error occurred: {str(e)}"}), 500


@app.route('/Stock112Emission/<string:stock_code>',methods=["GET"])
@swag_from({
    'tags': ['112y Emission'],
    'summary': 'Retrieve emission data for a specific stock',
    'Method':'Get',
    'parameters': [
        {
            'name': 'stock_code',
            'in': 'path',
            'type': 'string',
            'required': True,
            'description': 'Stock code of the company (e.g., 2330 for TSMC)'
        }
    ],
    'produces': ['application/json'],
    'responses': {
        200: {
            'description': 'Emission data and other ESG-related metrics based on stock code',
            'type': {
                'application/json': {
                    '股票代號': {'type': 'string'},
                    '公司名稱': {'type': 'string'},
                    '直接(範疇一)溫室氣體排放量(公噸)': {'type': 'float'},
                    '能源間接(範疇二)溫室氣體排放量(公噸)': {'type': 'float'},
                    '其他間接(範疇三)溫室氣體排放量(公噸)': {'type': 'float'},
                    '溫室氣體排放密集度': {'type': 'float'},
                    '再生能源使用率': {'type': 'float'},
                    '用水量(公噸(t))': {'type': 'float'},
                    '用水密集度': {'type': 'float'},
                    '有害廢棄物(公噸(t))': {'type': 'float'},
                    '非有害廢棄物(公噸(t))': {'type': 'float'},
                    '總重量(有害+非有害)(公噸(t))': {'type': 'float'},
                    '廢棄物密集度(公噸)': {'type': 'float'},
                    '員工福利平均數(仟元/人)': {'type': 'float'},
                    '員工薪資平均數(仟元/人)': {'type': 'float'},
                    '非擔任主管職務之全時員工薪資平均數(仟元/人)': {'type': 'float'},
                    '非擔任主管職務之全時員工薪資中位數(仟元/人)': {'type': 'float'},
                    '管理職女性主管占比': {'type': 'float'},
                    '職業災害人數': {'type': 'float'},
                    '職業災害比率': {'type': 'float'},
                    '火災件數': {'type': 'float'},
                    '火災死傷人數': {'type': 'float'},
                    '火災比率(死傷人數/員工總人數)': {'type': 'float'},
                    '董事會席次(席)': {'type': 'float'},
                    '獨立董事席次(席)': {'type': 'float'},
                    '女性董事比例': {'type': 'float'},
                    '董事出席董事會出席率': {'type': 'float'},
                    '公司年度召開法說會次數(次)': {'type': 'float'},
                    'ESG': {'type': 'float'}
                }
            }
        },
        404: {
            'description': 'Stock not found'
        }
    }
})
def get_listed_112_emission(stock_code):
    try:
        stock_emission=Emission112.query.filter_by(stock=stock_code).first()
        if stock_emission:
            return jsonify({
                "股票代號": stock_emission.stock,
                "公司名稱": stock_emission.stocknm,
                "直接(範疇一)溫室氣體排放量(公噸)": stock_emission.directed_emission,
                "能源間接(範疇二)溫室氣體排放量(公噸)": stock_emission.indirected_emission,
                "其他間接(範疇三)溫室氣體排放量(公噸)": stock_emission.other_indirected_emission,
                "溫室氣體排放密集度": stock_emission.greenHouse_emission_intensity,
                "再生能源使用率": stock_emission.renewable_energy,
                "用水量(公噸(t))": stock_emission.waterConsumption,
                "用水密集度": stock_emission.water_intensity,
                "有害廢棄物(公噸(t))": stock_emission.hazardous,
                "非有害廢棄物(公噸(t))": stock_emission.unhazardous,
                "總重量(有害+非有害)(公噸(t))": stock_emission.Total_waste,
                "廢棄物密集度(公噸)": stock_emission.waste_intensity,
                "員工福利平均數(仟元/人)": stock_emission.employee_avg_benefit,
                "員工薪資平均數(仟元/人)": stock_emission.employee_avg_salary,
                "非擔任主管職務之全時員工薪資平均數(仟元/人)": stock_emission.non_executive_avg_salary,
                "非擔任主管職務之全時員工薪資中位數(仟元/人)": stock_emission.non_executive_med_salary,
                "管理職女性主管占比": stock_emission.proportion_of_female_executive,
                "職業災害人數": stock_emission.occupational_accident,
                "職業災害比率": stock_emission.proportion_occupational_accident,
                "火災件數": stock_emission.num_of_fire,
                "火災死傷人數": stock_emission.death_of_fire,
                "火災比率(死傷人數/員工總人數)": stock_emission.proportion_death_of_fire,
                "董事會席次(席)": stock_emission.num_of_president,
                "獨立董事席次(席)": stock_emission.num_of_indpresident,
                "女性董事比例": stock_emission.proportion_of_female_president,
                "董事出席董事會出席率": stock_emission.attence_of_president,
                "公司年度召開法說會次數(次)": stock_emission.investor_conference,
                "ESG": stock_emission.ESG
            })
        else:
            return jsonify({'error': 'Stock not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/FinancialGrowthRate/<string:stock>', methods=['GET'])
@swag_from({
    'tags': ['Financial Growth Rate'],
    'summary': 'Retrieve 5-year growth rates for a specific stock',
    'Method':'Get',
    'parameters': [
        {
            'name': 'stock_code',
            'in': 'path',
            'type': 'string',
            'required': True,
            'description': 'Stock code of the company (e.g., 2330 for TSMC)'
        }
    ],
    'responses': {
        200: {
            'description': '5-year financial growth rates for the stock',
            'type': {
                'application/json':{
                    '股票代號': {'type': 'string'},
                    '公司名稱': {'type': 'string'},
                    '毛利率2023年度成長率(%)': {'type': 'float'},
                    '毛利率2022年度成長率(%)': {'type': 'float'},
                    '毛利率2021年度成長率(%)': {'type': 'float'},
                    '毛利率2020年度成長率(%)': {'type': 'float'},
                    '毛利率2019年度成長率(%)': {'type': 'float'},
                    '營益率2023年度成長率(%)': {'type': 'float'},
                    '營益率2022年度成長率(%)': {'type': 'float'},
                    '營益率2021年度成長率(%)': {'type': 'float'},
                    '營益率2020年度成長率(%)': {'type': 'float'},
                    '營益率2019年度成長率(%)': {'type': 'float'},
                    '淨利率2023年度成長率(%)': {'type': 'float'},
                    '淨利率2022年度成長率(%)': {'type': 'float'},
                    '淨利率2021年度成長率(%)': {'type': 'float'},
                    '淨利率2020年度成長率(%)': {'type': 'float'},
                    '淨利率2019年度成長率(%)': {'type': 'float'}
                    }
                },
            },
        404: {
            'description': 'Stock not found'
        }
    }
})
def get_financial_growth_rate(stock):
    # 根據個股代號查詢數據
    result = FinancialGrowthRate.query.filter_by(stock=stock).first()
    if result:
        return jsonify({
            '股票代號': result.stock,
            '公司名稱': result.stocknm,
            '毛利率2023年度成長率(%)': result.gross_margin_2023,
            '毛利率2022年度成長率(%)': result.gross_margin_2022,
            '毛利率2021年度成長率(%)': result.gross_margin_2021,
            '毛利率2020年度成長率(%)': result.gross_margin_2020,
            '毛利率2019年度成長率(%)': result.gross_margin_2019,
            '營益率2023年度成長率(%)': result.operating_income_2023,
            '營益率2022年度成長率(%)': result.operating_income_2022,
            '營益率2021年度成長率(%)': result.operating_income_2021,
            '營益率2020年度成長率(%)': result.operating_income_2020,
            '營益率2019年度成長率(%)': result.operating_income_2019,
            '淨利率2023年度成長率(%)': result.net_profit_2023,
            '淨利率2022年度成長率(%)': result.net_profit_2022,
            '淨利率2021年度成長率(%)': result.net_profit_2021,
            '淨利率2020年度成長率(%)': result.net_profit_2020,
            '淨利率2019年度成長率(%)': result.net_profit_2019
        }), 200
    else:
        return jsonify({"error": "Stock not found"}), 404


if __name__=="__main__":
    with app.app_context():
        listings = ListedInfo.query.all()
    app.run(debug=True)
