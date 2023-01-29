"""
Файл содержит вспомогательные функции проекта

"""

import pandas as pd
import pandas.io.formats.excel
from os.path import basename

from settings import BASE_DIR, NUM_MONTHS, DATE_COL
from settings import TABLE_PURCHASES, TABLE_SALES_HOLDINGS, TABLE_SALES


def get_filtered_df(excel, dict_warehouses, name_column_whs, skiprows=0):
    """
    Возвращает dataframe из excel, отфильтрованный по складам
    с заменой наименований складов
    """

    full_df = excel.parse(skiprows=skiprows)
    filter_df = full_df[full_df[name_column_whs].isin(
        list(dict_warehouses.keys())
    )]
    return filter_df.replace({name_column_whs: dict_warehouses})


def save_to_excel(path, df):
    """Записывает данные в excel с настройкой форматов"""

    sheet = 'Лист1'
    num_row_header = 0
    height_header = 50
    min_wdh_col = 14
    max_wdh_col = 50

    # Убираем формат заголовка таблицы по умолчанию
    pandas.io.formats.excel.ExcelFormatter.header_style = None

    for column in df:
        if column in DATE_COL:
            to_date_dict = {}
            for val in df[column]:
                to_date_dict[val] = val.date()
            df = df.replace({column: to_date_dict})

    writer = pd.ExcelWriter(path)
    df.to_excel(writer, sheet_name=sheet, index=False)
    workbook = writer.book
    worksheet = writer.sheets[sheet]
    cell_form = workbook.add_format({
        'text_wrap': True, 'bold': True,
        'align': 'center', 'valign': 'vcenter'
    })
    worksheet.set_row(num_row_header, height_header, cell_format=cell_form)

    for column in df:
        column_width = min(
            max(
                df[column].astype(str).map(len).max(),
                min_wdh_col,
                len(column) / 2
            ),
            max_wdh_col
        )
        col_idx = df.columns.get_loc(column)
        writer.sheets[sheet].set_column(col_idx, col_idx, column_width)
    writer.close()


def get_col_purch(df):
    """Возвращает словарь с числовыми столбцами закупок"""

    columns = list(df.columns)
    last_month = columns[-1]
    first_month = columns[-NUM_MONTHS]
    col_dict = {
        'first': first_month,
        'last': last_month,
        'all': columns
    }
    return col_dict


def get_col_sales(df):
    """Возвращает словарь с числовыми столбцами продаж"""
    col_view = 3
    avg_col = 3

    columns = list(df.columns)[-(NUM_MONTHS * col_view + avg_col):]
    cuts = columns[:NUM_MONTHS]
    sales = columns[NUM_MONTHS: NUM_MONTHS + NUM_MONTHS]
    cuts_sales = columns[-(NUM_MONTHS + NUM_MONTHS): -NUM_MONTHS]
    avg = columns[-NUM_MONTHS:]
    last_cut = cuts[-1]
    last_sale = sales[-1]
    last_cut_sale = cuts_sales[-1]
    penultimate_sale = sales[-2]
    last = [last_cut, last_sale, last_cut_sale]
    avg_cut = avg[0]
    avg_sale = avg[1]
    avg_cut_sale = avg[2]

    col_dict = {
        'cuts': cuts,
        'sales': sales,
        'cuts_sales': cuts_sales,
        'avg': avg,
        'last_cut': last_cut,
        'last_sale': last_sale,
        'last_cut_sale': last_cut_sale,
        'pntm_sale': penultimate_sale,
        'last': last,
        'avg_cut': avg_cut,
        'avg_sale': avg_sale,
        'avg_cut_sale': avg_cut_sale,
    }

    return col_dict


def get_data(file):
    """Возвращает dataframe из обработанных данных"""
    excel = pd.ExcelFile(f'{BASE_DIR}/Результаты/{file}')
    df = excel.parse()
    if file == TABLE_PURCHASES:
        return df, get_col_purch(df)
    if file in [TABLE_SALES_HOLDINGS, TABLE_SALES]:
        return df, get_col_sales(df)
    return df


def print_complete(file):
    """Функция выводит сообщение об отработке скрипта"""
    print('Скрипт {} выполнен!'.format(basename(file)))


def get_mult_clients_dict(df, col):
    """
    Возвращает словарь по строкам с несколькими клиентами, где ключ все
    клиенты одной строкой, а значение все клиенты раздельно в списке
    """
    clients = list(set(df[col].to_list()))
    mult_clients = [i for i in clients if '), ' in str(i)]
    mult_clients_dict = {}
    for mult_client in mult_clients:
        mult_client_list = mult_client.split('), ')
        mult_client_list_pure = []

        for client in mult_client_list:
            if client != mult_client_list[len(mult_client_list) - 1]:
                mult_client_list_pure.append(client + ')')
            else:
                mult_client_list_pure.append(client)
        mult_clients_dict[mult_client] = mult_client_list_pure
    return mult_clients_dict
