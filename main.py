from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from dateutil.relativedelta import relativedelta
from jinja2 import Environment, FileSystemLoader, select_autoescape

from settings import YEAR_FOUNDATION, return_filepath_and_sheet


def get_name_year(years):
    remainder = years % 10
    if remainder == 1:
        return 'год'
    if remainder > 1 and remainder < 5:
        return 'года'
    else:
        return 'лет'


def make_wine_dict_by_category(wine_list):
    wines_by_categories = defaultdict(list)
    for wine in wine_list:
        wines_by_categories[wine['Категория']].append(wine)
    return wines_by_categories


def get_wine_list_from_excel(filepath, sheet_name):
    excel_data_df = pandas.read_excel(filepath, sheet_name=sheet_name,
                                    #   usecols=["Категория", "Название", "Сорт", "Цена", "Картинка", "Акция"],
                                      na_values='NaN', keep_default_na=False)
    excel_data_df = excel_data_df.astype({'Цена': 'int32'})
    wine_list = excel_data_df.to_dict(orient='records')
    return wine_list


def years_from_foundation():
    date_start = datetime(YEAR_FOUNDATION, 1, 1)
    date_now = datetime.now()
    years = relativedelta(date_now, date_start).years
    return years


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    filepath, sheet_name = return_filepath_and_sheet()
    try:
        wine_list = get_wine_list_from_excel(filepath, sheet_name)
    except FileNotFoundError:
        return print('Такого файла не существует!')
    except ValueError:
        return print('Такого листа в файле нет!')

    wines_by_categories = make_wine_dict_by_category(wine_list)
    categories = sorted(wines_by_categories.keys())
    
    template = env.get_template('template.html')
    years = years_from_foundation()
    name_year = get_name_year(years)
    rendered_page = template.render(
        years=years, 
        name_year=name_year,
        wines=wines_by_categories,
        categories=categories
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
