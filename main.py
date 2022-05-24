from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from dateutil.relativedelta import relativedelta
from jinja2 import Environment, FileSystemLoader, select_autoescape

from settings import (DEFAULT_ENVPATH, DEFAULT_EXCELPATH, DEFAULT_SHEET,
                      YEAR_FOUNDATION, return_filepath_and_sheet)


def make_wine_dict_by_category(wines_table):
    wines = defaultdict(list)
    for wine in wines_table:
        wines[wine['Категория']].append(wine)
    return wines


def get_wine_list_from_excel(filepath, sheet_name):
    excel_data_df = pandas.read_excel(filepath, sheet_name=sheet_name,
                                      na_values='NaN', keep_default_na=False)
    excel_data_df = excel_data_df.astype({'Цена': 'int32'})
    wine_list = excel_data_df.to_dict(orient='records')
    return wine_list


def years_from_foundation(foundation_year):
    date_start = datetime(foundation_year, 1, 1)
    date_now = datetime.now()
    years = relativedelta(date_now, date_start).years
    return years


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    kwargs = {'default_envpath': DEFAULT_ENVPATH,
              'default_excelpath': DEFAULT_EXCELPATH,
              'default_sheet': DEFAULT_SHEET}
    filepath, sheet_name = return_filepath_and_sheet(**kwargs)
    try:
        wines_table = get_wine_list_from_excel(filepath, sheet_name)
    except FileNotFoundError:
        return print('Такого файла не существует!')
    except ValueError:
        return print('Такого листа в файле нет!')

    wines = make_wine_dict_by_category(wines_table)
    
    template = env.get_template('template.html')
    years = years_from_foundation(YEAR_FOUNDATION)
    rendered_page = template.render(
        years=years,
        wines=wines,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
