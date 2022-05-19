from http.server import HTTPServer, SimpleHTTPRequestHandler
from multiprocessing.spawn import WINSERVICE
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
from dateutil.relativedelta import relativedelta
from pprint import pprint
import pandas

def get_name_year(years):
    remainder = years % 10
    if remainder == 1:
        return 'год'
    if remainder > 1 and remainder < 5:
        return 'года'
    else:
        return 'лет'


def make_wines(wine_data):
    wines = {}
    for wine in wine_data:
        if wine['Категория'] not in wines:
            wines[wine['Категория']] = []
        wines[wine['Категория']].append(wine)
    return wines


def get_excel_data(file, sheet):
    excel_data_df = pandas.read_excel(file, sheet_name=sheet,
                                      usecols=["Категория", "Название", "Сорт", "Цена", "Картинка", "Акция"],
                                      na_values='NaN', keep_default_na=False)
    excel_data_df = excel_data_df.astype({'Цена': 'int32'})
    wine_data = excel_data_df.to_dict(orient='records')
    return wine_data


def get_years():
    date_start = datetime(1920, 1, 1)
    date_now = datetime.now()
    years = relativedelta(date_now, date_start).years
    return years


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    excel_file = 'wine3.xlsx'
    sheet = 'Лист1'
    wine_data = get_excel_data(excel_file, sheet)
    wines = make_wines(wine_data)
    # сортируем словарь, чтобы вина были первыми, потом напитки
    wines = sorted(wines.items())
    
    template = env.get_template('template.html')
    years = get_years()
    name_year = get_name_year(years)
    rendered_page = template.render(
        years=years, 
        name_year=name_year,
        wines=wines,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
