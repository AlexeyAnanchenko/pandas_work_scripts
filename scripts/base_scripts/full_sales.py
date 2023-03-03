"""
Скрипт формирует продажи за большой период (с октября 2021)

"""
import utils
utils.path_append()

import pandas as pd

from hidden_settings import WAREHOUSES_SALES
from service import get_filtered_df, save_to_excel, print_complete
from service import get_mult_clients_dict
from settings import WHS, EAN, NAME_HOLDING, SOURCE_DIR, RESULT_DIR, LINK
from settings import TABLE_FULL_SALES_CLIENTS, TABLE_FULL_SALES, LINK_HOLDING


SOURCE_FILE_SALES = 'Продажи ВСЕ по клиентам.xlsx'
SOURCE_FILE_FACTORS = 'Факторы Реестр.xlsx'
FIRST_DATE_REGISTRY = 'Октябрь 2021'
EMPTY_ROWS = 12
WHS_LOC = 'Склад'
NAME_HOLDING_LOC = 'Основной холдинг'
EAN_LOC = 'EAN'
DIRECTION = 'Направление продаж'
NAME_HOLDING__LOC_FACTORS = 'Корректный холдинг'


def get_factors():
    """Получаем реестр факторы"""
    excel = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE_FACTORS)
    df = excel.parse()
    mult_clients = get_mult_clients_dict(df, NAME_HOLDING__LOC_FACTORS)
    list_clients = df[NAME_HOLDING__LOC_FACTORS].tolist()
    for clients in mult_clients:
        list_clients = list_clients + mult_clients[clients]
    list_clients = list(set(list_clients))
    return list_clients


def get_sales(list_clients):
    """Формирование фрейма данных продаж в разрезе клиент-склад-шк"""
    excel = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE_SALES)
    df_sales = get_filtered_df(
        excel, WAREHOUSES_SALES, WHS_LOC, skiprows=EMPTY_ROWS
    )
    df_sales.drop(DIRECTION, axis=1, inplace=True)
    df_sales = df_sales.rename(columns={
        WHS_LOC: WHS, NAME_HOLDING_LOC: NAME_HOLDING, EAN_LOC: EAN
    })

    index_col = df_sales.columns.get_loc(FIRST_DATE_REGISTRY)
    sales_col = df_sales.columns[index_col:]

    agg_sales_col = {}
    for col in sales_col:
        agg_sales_col[col] = 'sum'

    df_sales_whs = df_sales.groupby(
        [WHS, EAN]
    ).agg(agg_sales_col).reset_index()
    df_sales = df_sales[df_sales[NAME_HOLDING].isin(list_clients)]

    df_sales.insert(
        0, LINK_HOLDING,
        (df_sales[WHS] + df_sales[NAME_HOLDING]
         + df_sales[EAN].map(int).map(str))
    )
    for df in [df_sales, df_sales_whs]:
        df.insert(
            0, LINK,
            df[WHS] + df[EAN].map(int).map(str)
        )

    for col in sales_col:
        df_sales = utils.void_to(df_sales, col, 0)

    for df in [df_sales, df_sales_whs]:
        count = 1
        sales_col_copy = sales_col.copy()
        for col in sales_col:
            df[str(count)] = df[sales_col_copy].sum(axis=1)
            sales_col_copy = sales_col_copy[1:]
            count += 1

    df_sales = df_sales.drop(columns=sales_col, axis=1)
    df_sales_whs = df_sales_whs.drop(columns=sales_col, axis=1)
    return df_sales, df_sales_whs


def main():
    df_sales, df_sales_whs = get_sales(get_factors())
    save_to_excel(RESULT_DIR + TABLE_FULL_SALES_CLIENTS, df_sales)
    save_to_excel(RESULT_DIR + TABLE_FULL_SALES, df_sales_whs)
    print_complete(__file__)


if __name__ == "__main__":
    main()
