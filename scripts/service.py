"""
Файл содержит вспомогательные функции и глобальные переменные проекта

"""

import pandas as pd
import pandas.io.formats.excel


NUM_MONTHS = 3
EAN = 'EAN штуки'
PRODUCT = 'Наименование товара'
SU = 'SU штуки'
MSU = 'MSU штуки'
LEVEL_1 = '1-й уровень'
LEVEL_2 = '2-й уровень'
LEVEL_3 = '3-й уровень'
HOLDING = 'Код основного холдинга'
NAME_HOLDING = 'Наименование основного холдинга'
CODES = 'Код Точки доставки / Плательщика / Холдинга / Основного холдинга'
BASE_PRICE = 'Базовая цена без НДС, шт'
ELB_PRICE = 'Цена NIV Эльбрус с НДС, шт'
WHS = 'Склад'
LINK = 'Сцепка Склад-ШК'
LINK_HOLDING = 'Сцепка Склад-Наименование холдинга-ШК'
FULL_REST = 'Общий остаток, шт'
AVAILABLE_REST = 'Доступный остаток, шт'
QUOTA = 'Квота, шт'
FREE_REST = 'Свободный остаток, шт'
SOFT_RSV = 'Мягкий резерв, шт'
HARD_RSV = 'Жесткий резерв, шт'
SOFT_HARD_RSV = 'Мягкие + жёсткие резервы, шт'
QUOTA_BY_AVAILABLE = 'Резерв квота с учётом доступного стока'
TOTAL_RSV = 'В резерве всего'
CUTS = 'Урезания '
SALES = 'Продажи '
CUTS_SALES = 'Урезания + продажи '
AVARAGE = 'Средние '


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

    # Убираем формат заголовка таблицы по умолчанию
    pandas.io.formats.excel.ExcelFormatter.header_style = None

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
        column_width = max(
            df[column].astype(str).map(len).max(),
            min_wdh_col,
            len(column) / 2
        )
        col_idx = df.columns.get_loc(column)
        writer.sheets[sheet].set_column(col_idx, col_idx, column_width)

    writer.close()


def numeric_col_purch(df):
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
