from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from dateutil.relativedelta import relativedelta
from jinja2 import Environment, FileSystemLoader, select_autoescape
from settings import return_filepath_and_sheet

EXCEL_FILENAME = "wine_table.xlsx"
EXCEL_SHEETNAME = 'Вина'


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


def get_excel_data(filepath, sheet_name):
    excel_data_df = pandas.read_excel(filepath, sheet_name=sheet_name,
                                    #   usecols=["Категория", "Название", "Сорт", "Цена", "Картинка", "Акция"],
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
    filepath, sheet_name = return_filepath_and_sheet()
    wine_data = get_excel_data(filepath, sheet_name)
    wines = make_wines(wine_data)
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
