import os
from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader, select_autoescape

from terminal_utils import parse_terminal_commands

DEFAULT_ENVPATH = '.env'
DEFAULT_EXCELPATH = 'wine_table.xlsx'
DEFAULT_SHEET = 'Вина'
YEAR_FOUNDATION = 1920


def get_wines_table_from_excel(filepath, sheet_name):
    excel_data_df = pandas.read_excel(filepath, sheet_name=sheet_name,
                                      na_values='NaN', keep_default_na=False)
    excel_data_df = excel_data_df.astype({'Цена': 'int32'})
    wine_table = excel_data_df.to_dict(orient='records')
    return wine_table


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

    terminal_args = parse_terminal_commands(**kwargs)

    if terminal_args.method == "env":
        load_dotenv(terminal_args.envpath)
        excel_filepath = os.getenv('EXCEL_FILEPATH')
        excel_sheet = os.getenv('EXCEL_SHEET')
    else:
        excel_filepath = terminal_args.excelpath
        excel_sheet = terminal_args.sheet

    try:
        wines_table = get_wines_table_from_excel(excel_filepath, excel_sheet)
    except FileNotFoundError:
        return print('Такого файла не существует!')
    except ValueError:
        return print('Такого листа в файле нет!')
    wines = defaultdict(list)
    for wine in wines_table:
        wines[wine['Категория']].append(wine)
    template = env.get_template('template.html')
    years = years_from_foundation(YEAR_FOUNDATION)
    rendered_page = template.render(years=years, wines=wines)

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
