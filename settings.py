import argparse
import os

from dotenv import load_dotenv

DEFAULT_ENVPATH = '.env'
DEFAULT_EXCELPATH = 'wine_table.xlsx'
DEFAULT_SHEET = 'Вина'
YEAR_FOUNDATION = 1920


def get_excel_name_and_sheet(args):
    if args.method == "env":
        load_dotenv(args.envpath)
        excel_filepath = os.getenv('EXCEL_FILEPATH', default=DEFAULT_EXCELPATH)
        excel_sheet = os.getenv('EXCEL_SHEET', default=DEFAULT_SHEET)
    else:
        excel_filepath = args.excelpath
        excel_sheet = args.sheet
    return excel_filepath, excel_sheet


def parse_terminal_commands():
    parser = argparse.ArgumentParser(
    description='Описание что делает программа:'
                )
    parser.add_argument("-m", "--method", choices=['exc', 'env'], help="Выбор метода указания пути к Excel файлу: "
                        "и имени листа с винами, 'env' - из файла .env, "
                        "'exc' - из консоли'",
                        default="env")
    parser.add_argument("-pexc", "--excelpath", help="Путь к файлу Excel (по умолчанию - 'wine_table.xlsx'",
                        default=DEFAULT_EXCELPATH)
    parser.add_argument("-penv", "--envpath", help="Путь к файлу .env (по умолчанию - '.env')", 
                        default=DEFAULT_ENVPATH)
    parser.add_argument("-s", "--sheet", help="Имя листа в файле Excel (По умолчанию - 'Вина'",
                        default=DEFAULT_SHEET)
    args = parser.parse_args()
    return args


def return_filepath_and_sheet():
    args = parse_terminal_commands()
    filepath, excel_sheet = get_excel_name_and_sheet(args)
    return filepath, excel_sheet
