import argparse


def parse_terminal_commands(default_excelpath, default_envpath, default_sheet):
    parser = argparse.ArgumentParser(description='Описание что делает программа:')
    
    parser.add_argument("-m", "--method", choices=['exc', 'env'], help="Выбор метода указания пути к Excel файлу "
                        "и имени листа с винами: 'env' - из файла .env, "
                        "'exc' - из консоли (по умолчанию env)",
                        default="env")
    
    parser.add_argument("-pexc", "--excelpath", help="Путь к файлу Excel (по умолчанию - 'wine_table.xlsx')",
                        default=default_excelpath)
    
    parser.add_argument("-penv", "--envpath", help="Путь к файлу .env (по умолчанию - '.env')", 
                        default=default_envpath)
    
    parser.add_argument("-s", "--sheet", help="Имя листа в файле Excel (По умолчанию - 'Вина')",
                        default=default_sheet)
    
    args = parser.parse_args()
    return args
