from flask import Flask,jsonify,request
from flask_sqlalchemy import SQLAlchemy
import pyodbc
from flasgger import Swagger, swag_from
from flask_cors import CORS
from sqlalchemy.exc import IntegrityError
from Sustainable_ESG import CatchNews
app=Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI']=r"mssql+pyodbc://Yogurt:stock0050@LAPTOP-GO7LV6HA\MSSQLSERVER2022/ESGInfoHub?driver=ODBC+Driver+17+for+SQL+Server"
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
    __tablename__ = 'ListedInfo'  
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

class Emission(db.Model):
    __tablename__="CompanyESGTable"
    stock=db.Column("公司代號",db.String(6),primary_key=True)
    stocknm=db.Column("公司名稱",db.String(32))
    year = db.Column("資料年度",db.String(4),primary_key =True)
    industry = db.Column("產業類別",db.String(16))
    E = db.Column("E",db.Float)
    S = db.Column("S",db.Float)
    G = db.Column("G",db.Float)
    ESG=db.Column("ESG",db.Float)
    def __init__(self,stock,stocknm, year,E,S,G,ESG):
        self.stock=stock
        self.stocknm=stocknm
        self.year = year
        self.E = E
        self.S = S
        self.G = G
        self.ESG=ESG

class FinancialGrowthRate(db.Model):
    __tablename__ = "FinancialGrowthRate"
    stock = db.Column("股票代號", db.String(6), primary_key=True)         # 股票代號
    stocknm = db.Column("公司名稱", db.String(32))                          # 公司名稱
    industry = db.Column("產業類別", db.String(32))                         # 產業類別

    gross_margin_2023 = db.Column("2023年度毛利增減(%)", db.Float)          # 2023年度毛利增減(%)
    gross_margin_2022 = db.Column("2022年度毛利增減(%)", db.Float)          # 2022年度毛利增減(%)
    gross_margin_2021 = db.Column("2021年度毛利增減(%)", db.Float)          # 2021年度毛利增減(%)
    gross_margin_2020 = db.Column("2020年度毛利增減(%)", db.Float)          # 2020年度毛利增減(%)
    gross_margin_2019 = db.Column("2019年度毛利增減(%)", db.Float)          # 2019年度毛利增減(%)

    operating_income_2023 = db.Column("2023年度營益增減(%)", db.Float)       # 2023年度營益增減(%)
    operating_income_2022 = db.Column("2022年度營益增減(%)", db.Float)       # 2022年度營益增減(%)
    operating_income_2021 = db.Column("2021年度營益增減(%)", db.Float)       # 2021年度營益增減(%)
    operating_income_2020 = db.Column("2020年度營益增減(%)", db.Float)       # 2020年度營益增減(%)
    operating_income_2019 = db.Column("2019年度營益增減(%)", db.Float)       # 2019年度營益增減(%)

    net_profit_2023 = db.Column("2023年度淨利增減(%)", db.Float)             # 2023年度淨利增減(%)
    net_profit_2022 = db.Column("2022年度淨利增減(%)", db.Float)             # 2022年度淨利增減(%)
    net_profit_2021 = db.Column("2021年度淨利增減(%)", db.Float)             # 2021年度淨利增減(%)
    net_profit_2020 = db.Column("2020年度淨利增減(%)", db.Float)             # 2020年度淨利增減(%)
    net_profit_2019 = db.Column("2019年度淨利增減(%)", db.Float)             # 2019年度淨利增減(%)

    def __init__(self, stock, stocknm, industry,
                 gross_margin_2023, gross_margin_2022, gross_margin_2021, gross_margin_2020, gross_margin_2019,
                 operating_income_2023, operating_income_2022, operating_income_2021, operating_income_2020, operating_income_2019,
                 net_profit_2023, net_profit_2022, net_profit_2021, net_profit_2020, net_profit_2019):
        self.stock = stock
        self.stocknm = stocknm
        self.industry = industry
        
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

#爬取個股基本資訊的API
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

#爬取公司的ESG數據
@app.route('/Emission/<string:stock_code>/<string:year>',methods=["GET"])
@swag_from({
    'tags': ['Emission'],
    'summary': 'Retrieve emission data for a specific stock',
    'Method':'Get',
    'parameters': [
        {
            'name': 'stock_code',
            'in': 'path',
            'type': 'string',
            'required': True,
            'description': 'Stock code of the company (e.g., 2330 for TSMC)'
        },
        {
            'name': 'year',
            'in': 'path',
            'type': 'string',
            'required': True,
            'description': 'year of the company data (e.g., 110 for TSMC)'
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
                    '資料年度': {'type':'string'},
                    'E' : {'type':'float'},
                    'S' : {'type':'float'},
                    'G' : {'type':'float'},
                    'ESG': {'type': 'float'}
                }
            }
        },
        404: {
            'description': 'Stock not found'
        }
    }
})
def get_listed_112_emission(stock_code,year):
    try:
        stock_emission=Emission.query.filter_by(stock=stock_code,year = year).first()
        if stock_emission:
            return jsonify({
                "股票代號": stock_emission.stock,
                "公司名稱": stock_emission.stocknm,
                "資料年度": stock_emission.year,
                "E" : stock_emission.E,
                "S" : stock_emission.S,
                "G" : stock_emission.G,
                "ESG": stock_emission.ESG
            })
        else:
            return jsonify({'error': 'ESG數據沒有抓取成功'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

#爬取公司的三率成長率
@app.route('/FinancialGrowthRate/<string:stock_code>', methods=['GET'])
@swag_from({
    'tags': ['Financial Growth Rate'],
    'summary': 'Retrieve 5-year growth rates for a specific stock',
    'Method': 'Get',
    'parameters': [
        {
            'name': 'stock_code',
            'in': 'path',
            'type': 'string',
            'required': True,
            'description': '公司股票代號（例如：2330 代表台積電）'
        }
    ],
    'responses': {
        200: {
            'description': '回傳該股票 5 年內各項指標的增減數據',
            'schema': {
                'type': 'object',
                'properties': {
                    '股票代號': {'type': 'string'},
                    '公司名稱': {'type': 'string'},
                    '產業類別': {'type': 'string'},
                    '2023年度毛利增減(%)': {'type': 'number'},
                    '2022年度毛利增減(%)': {'type': 'number'},
                    '2021年度毛利增減(%)': {'type': 'number'},
                    '2020年度毛利增減(%)': {'type': 'number'},
                    '2019年度毛利增減(%)': {'type': 'number'},
                    '2023年度營益增減(%)': {'type': 'number'},
                    '2022年度營益增減(%)': {'type': 'number'},
                    '2021年度營益增減(%)': {'type': 'number'},
                    '2020年度營益增減(%)': {'type': 'number'},
                    '2019年度營益增減(%)': {'type': 'number'},
                    '2023年度淨利增減(%)': {'type': 'number'},
                    '2022年度淨利增減(%)': {'type': 'number'},
                    '2021年度淨利增減(%)': {'type': 'number'},
                    '2020年度淨利增減(%)': {'type': 'number'},
                    '2019年度淨利增減(%)': {'type': 'number'}
                }
            }
        },
        404: {
            'description': '找不到該股票資料'
        }
    }
})
def get_financial_growth_rate(stock_code):
    # 利用 SQLAlchemy 查詢該股票資料
    stock_data = FinancialGrowthRate.query.filter_by(stock=stock_code).first()
    if stock_data:
        response = {
            '股票代號': stock_data.stock,
            '公司名稱': stock_data.stocknm,
            '產業類別': stock_data.industry,
            '2023年度毛利增減(%)': stock_data.gross_margin_2023,
            '2022年度毛利增減(%)': stock_data.gross_margin_2022,
            '2021年度毛利增減(%)': stock_data.gross_margin_2021,
            '2020年度毛利增減(%)': stock_data.gross_margin_2020,
            '2019年度毛利增減(%)': stock_data.gross_margin_2019,
            '2023年度營益增減(%)': stock_data.operating_income_2023,
            '2022年度營益增減(%)': stock_data.operating_income_2022,
            '2021年度營益增減(%)': stock_data.operating_income_2021,
            '2020年度營益增減(%)': stock_data.operating_income_2020,
            '2019年度營益增減(%)': stock_data.operating_income_2019,
            '2023年度淨利增減(%)': stock_data.net_profit_2023,
            '2022年度淨利增減(%)': stock_data.net_profit_2022,
            '2021年度淨利增減(%)': stock_data.net_profit_2021,
            '2020年度淨利增減(%)': stock_data.net_profit_2020,
            '2019年度淨利增減(%)': stock_data.net_profit_2019,
        }
        return jsonify(response), 200
    else:
        return jsonify({"error": "Stock not found"}), 404

@app.route('/Catch_News/<string:stock>', methods=['GET'])
@swag_from({
    'tags': ['News'],
    'summary': 'Retrieve ESG News for a specific company',
    'Method': 'Get',
    'parameters': [
        {
            'name': 'stock',
            'in': 'path',
            'type': 'string',
            'required': True,
            'description': '公司名稱或股號（請確保可以對應到公司名稱，例如 "聯發科"）'
        }
    ],
    'responses': {
        200: {
            'description': '回傳指定公司相關的 ESG 新聞列表',
            'schema': {
                'type': 'array',
                'items': {
                    'type': 'object',
                    'properties': {
                        '公司名稱': {'type': 'string'},
                        '日期': {'type': 'string'},
                        '新聞標題': {'type': 'string'},
                        '新聞連結': {'type': 'string'},
                        '新聞內容': {'type': 'string'},
                        '新聞來源': {'type': 'string'}
                    }
                }
            }
        },
        404: {
            'description': '找不到該股票相關的資料'
        }
    }
})
def Get_News(stock):
    # 假設 ListedInfo 資料表中包含 stock 與公司名稱的對應關係
    stock_info = ListedInfo.query.filter_by(stock=stock).first()
    if not stock_info:
        return jsonify({"error": "Stock not found"}), 404

    try:
        # 以公司名稱去爬取新聞
        result = CatchNews.Get_AllNews(stock_info.stocknm)
        # 將股票代號加入每一筆新聞資料中
        for news_item in result:
            news_item["股票代號"] = stock_info.stock
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    if result:
        return jsonify(result), 200
    else:
        return jsonify({"error": "No news found"}), 404

if __name__=="__main__":
    with app.app_context():
        listings = ListedInfo.query.all()
    app.run(debug=True)
