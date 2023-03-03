"""
Отчёт по выявлению недогруженных объёмов по факторам

"""
import utils
utils.path_append()

import numpy as np
import pandas as pd

from service import save_to_excel, get_data, print_complete
from service import get_mult_clients_dict
from settings import REPORT_NOT_SOLD_PAST, REPORT_DIR, TABLE_FACTORS, PAST
from settings import CURRENT, FACTOR_PERIOD, PURPOSE_PROMO, INACTIVE_PURPOSE
from settings import SALES_FACTOR_PERIOD, LINK, LINK_HOLDING, WHS, NAME_HOLDING
from settings import EAN, PRODUCT, LEVEL_3, PLAN_NFE, ADJUSTMENT_PBI, FACT_NFE
from settings import DESCRIPTION, USER, SALES_CURRENT_FOR_PAST, TABLE_REMAINS
from settings import RSV_FACTOR_PERIOD_TOTAL, TABLE_PURCHASES, FREE_REST
from settings import ARCHIVE_DIR, FULL_REST, SOFT_HARD_RSV, QUOTA, OVERSTOCK
from settings import TABLE_SALES, TABLE_DIRECTORY, MSU, BASE_PRICE
from settings import REPORT_NOT_SOLD_CURRENT
from settings import SOURCE_DIR, REPORT_TRACKING_NOT_SOLD, TABLE_RESERVE
from settings import TABLE_FULL_SALES_CLIENTS, TABLE_FULL_SALES, TOTAL_RSV


SOURCE_FILE_FACTORS = 'Факторы Реестр.xlsx'
NAME_HOLDING_FACTORS = 'Корректный холдинг'
WHS_FACTORS = 'Город'
EAN_LOC = 'EAN'
NUM_MONTH = 'Порядковый номер месяца'
SALES_START_FROM_MONTH = ('Продажи клиента (-ов) начиная с'
                          ' месяца подачи фактора, шт')
SALES_WHS_START_FROM_MONTH = ('Продажи по складу после окончания'
                              ' действия фактора, шт')

LINK_HOLDING_PERIOD = 'Сцепка Склад-Холдинг-Штрихкод-Период'
LINK_PERIOD = 'Сцепка Склад-Штрихкод-Период'
PLAN = 'План, шт'
SALES_TOTAL = 'Продажи Тотал, шт'
SALES = 'Продажи в пределах фактора, шт'
RSV = 'Резервы (с квотой), шт'
FREE_REST_CURRENT = 'Свободные остатки текущего месяца, шт'
PURCH_PAST = 'Закупки прошлого месяца, шт'
PURCH_CURRENT = 'Закупки текущего месяца, шт'
FREE_REST_PAST = 'Свободные остатки для факторов прошедшего периода, шт'
PLAN_MINUS_SALES = 'План - Продажи, шт'
PLAN_MINUS_SALES_RSV = 'План - Продажи - Резервы, шт'
RSV_BY_PLAN = 'Резерв по Плану, шт'
RSV_BY_PLAN_TOTAL = 'В резерве по плану всего, шт'
RSV_BY_PLAN_PURCH = 'Резерв по Плану (с учётом закупа), шт'
PLAN_MINUS_SALES_RSV_TOTAL = 'План - Продажи - Резервы по сцепке, шт'
FREE_REST_BY_PLAN = 'Свободный остаток по плану, шт'
FREE_REST_BY_PLAN_TOTAL = 'Свободный остаток по плану по сцепке, шт'
FREE_REST_BY_PLAN_PURCH = 'Свободный остаток по плану (с учётом закупа), шт'
NOT_SOLD = 'Не продано, шт'
NOT_SOLD_PURCH = 'Не продано (с учётом закупа), шт'
NOT_SOLD_TOTAL = 'Не продано по сцепке, шт'
NOT_SOLD_PURCH_TOTAL = 'Не продано (с учётом закупа) по сцепке, шт'
RES_FOR_STOCK = 'Вклад в пересток, шт'
RES_FOR_STOCK_PURCH = 'Вклад в пересток (с учётом закупа), шт'
PRICE = 'Цена по направлениям бизнеса, руб'


def get_factors():
    """Получаем реестр факторов"""
    excel = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE_FACTORS)
    df = excel.parse()
    df = df.rename(columns={
        WHS_FACTORS: WHS, NAME_HOLDING_FACTORS: NAME_HOLDING, EAN_LOC: EAN
    })
    df.insert(
        0, LINK_HOLDING,
        df[WHS] + df[NAME_HOLDING] + df[EAN].map(int).map(str)
    )
    df.insert(
        0, LINK,
        df[WHS] + df[EAN].map(int).map(str)
    )
    return df


def process_merge_col(df_factors, df_merge, col_merge):
    """Процесс добавления данных по клиентам"""
    dict_mult_cl = get_mult_clients_dict(df_factors, NAME_HOLDING)
    if dict_mult_cl:
        df_factors_pure = df_factors[
            ~df_factors[NAME_HOLDING].isin(list(dict_mult_cl.keys()))
        ].copy()
        df_factors_pure = df_factors_pure.merge(
            df_merge[[LINK_HOLDING, col_merge]], on=LINK_HOLDING, how='left'
        )
        df_factors_clean = df_factors_pure.copy()
        df_factors_clean = df_factors_clean.drop(
            df_factors_clean.index[:len(df_factors_clean)]
        )

        for mult_client, mult_client_list in dict_mult_cl.items():
            df_factors_client = df_factors[
                df_factors[NAME_HOLDING] == mult_client
            ].copy()
            df_merge_clients = df_merge[
                df_merge[NAME_HOLDING].isin(mult_client_list)
            ].groupby([LINK]).agg({col_merge: 'sum'}).reset_index()
            df_factors_client = df_factors_client.merge(
                df_merge_clients, on=LINK, how='left'
            )
            df_factors_clean = pd.concat(
                [df_factors_clean, df_factors_client], ignore_index=True
            )

        df_factors = pd.concat(
            [df_factors_clean, df_factors_pure], ignore_index=True
        )
        return df_factors

    df_factors = df_factors.merge(
        df_merge[[LINK_HOLDING, col_merge]], on=LINK_HOLDING, how='left'
    )
    return df_factors


def process_merge_sales(df, table, new_col, whs=False):
    """Процесс добавления продаж по месяцам"""
    nums_month = list(set(df[NUM_MONTH].tolist()))
    df[new_col] = 0
    df_clean = df.copy().drop(df.index[:len(df)])
    for num_month in nums_month:
        df_iter = df[df[NUM_MONTH] == num_month].copy()
        df_iter = df_iter.drop(columns=[new_col], axis=1)
        df_sales = get_data(table)

        if whs is True:
            df_iter = df_iter.merge(
                df_sales[[LINK, str(num_month + 1)]], on=LINK, how='left'
            )
            df_iter = df_iter.rename(columns={
                str(num_month + 1): new_col
            })
        else:
            df_iter = process_merge_col(df_iter, df_sales, str(num_month))
            df_iter = df_iter.rename(columns={
                str(num_month): new_col
            })
        df_clean = pd.concat([df_clean, df_iter], ignore_index=True)
    return df_clean


def merge_sales(df):
    """Добавляем продажи по клиентам и по складам"""
    df = process_merge_sales(
        df, TABLE_FULL_SALES_CLIENTS, SALES_START_FROM_MONTH
    )
    df = process_merge_sales(
        df, TABLE_FULL_SALES, SALES_WHS_START_FROM_MONTH, whs=True
    )
    return df


def merge_reserve(df):
    """Добавляем резервы по клиентам"""
    df_rsv = get_data(TABLE_RESERVE)
    df = process_merge_col(df, df_rsv, TOTAL_RSV)
    return df


def main():
    df = merge_reserve(merge_sales(get_factors()))
    save_to_excel(REPORT_DIR + REPORT_TRACKING_NOT_SOLD, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
