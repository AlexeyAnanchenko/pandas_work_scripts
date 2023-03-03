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
from settings import EAN, PRODUCT, LEVEL_3
from settings import DESCRIPTION, USER, SALES_CURRENT_FOR_PAST, TABLE_REMAINS
from settings import RSV_FACTOR_PERIOD_TOTAL, TABLE_PURCHASES, FREE_REST
from settings import ARCHIVE_DIR, FULL_REST, SOFT_HARD_RSV, QUOTA, OVERSTOCK
from settings import TABLE_SALES, TABLE_DIRECTORY, MSU, BASE_PRICE
from settings import REPORT_NOT_SOLD_CURRENT
from settings import SOURCE_DIR, REPORT_TRACKING_NOT_SOLD, TABLE_RESERVE
from settings import TABLE_FULL_SALES_CLIENTS, TABLE_FULL_SALES, TOTAL_RSV
from settings import ALL_CLIENTS, NAME_TRAD, TABLE_REMAINS


SOURCE_FILE_FACTORS = 'Факторы Реестр.xlsx'
NAME_HOLDING_FACTORS = 'Корректный холдинг'
WHS_FACTORS = 'Город'
EAN_LOC = 'EAN'
NUM_MONTH = 'Порядковый номер месяца'
SALES_START_FROM_MONTH = ('Продажи клиента (-ов) начиная с'
                          ' месяца подачи фактора, шт')
SALES_WHS_START_FROM_MONTH = ('Продажи по складу после окончания'
                              ' действия фактора, шт')
PLAN_NFE = 'План в NFE, шт'
ADJUSTMENT_NFE = 'Корректировка в NFE, шт'
PURCHASES = 'Закупки в течение месяца, шт'
FULL_REST_MONTH = 'Сток на 1 число следующего месяца, шт'
MONTH = 'Месяц'
PLAN = 'Максимальный План, шт'
SALES = 'Продажи в пределах объёма фактора, шт'
RSV = 'Резервы (с квотой), шт'
FREE_REST_CALC = 'Свободный остаток для расчёта, шт'
OVERSTOCK_CALC = OVERSTOCK + '(Для расчёта)'
FULL_REST_CALC = 'Полный остаток для расчёта, шт'
PLAN_MINUS_SALES = 'План - Продажи, шт'
PLAN_MINUS_SALES_RSV = 'План - Продажи - Резервы, шт'
LINK_PERIOD = 'Сцепка Месяц фактора-Склад-ШК'

RSV_BY_PLAN = 'Резерв по Плану, шт'
RSV_BY_PLAN_DUBLE = 'Резерв по Плану, шт (копия)'
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


def process_merge_col(df_factors, df_merge, col_merge, df_merge_whs):
    """Процесс добавления данных по клиентам"""
    df_all_clients = df_factors[
        df_factors[NAME_HOLDING].isin([ALL_CLIENTS, NAME_TRAD])
    ].copy()
    if df_all_clients.empty:
        df_all_clients = df_factors.copy()
        df_all_clients[col_merge] = 0
        df_all_clients = df_all_clients.drop(
            df_all_clients.index[:len(df_all_clients)]
        )
    else:
        df_all_clients = df_all_clients.merge(
            df_merge_whs[[LINK, col_merge]],
            on=LINK, how='left'
        )

    df_factors = df_factors[
        ~df_factors[NAME_HOLDING].isin([ALL_CLIENTS, NAME_TRAD])
    ]
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
            [df_all_clients, df_factors_clean, df_factors_pure],
            ignore_index=True
        )
        return df_factors

    df_factors = df_factors.merge(
        df_merge[[LINK_HOLDING, col_merge]], on=LINK_HOLDING, how='left'
    )
    df_factors = pd.concat([df_all_clients, df_factors], ignore_index=True)
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
            df_iter = process_merge_col(
                df_iter, df_sales, str(num_month),
                get_data(TABLE_FULL_SALES)
            )
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
    for col in [SALES_START_FROM_MONTH, SALES_WHS_START_FROM_MONTH]:
        df = utils.void_to(df, col, 0)
    return df


def merge_reserve(df):
    """Добавляем резервы по клиентам"""
    df_rsv = get_data(TABLE_RESERVE)
    df_rsv_group = df_rsv.copy().groupby([LINK]).agg({
        TOTAL_RSV: 'sum'
    }).reset_index()
    df = process_merge_col(df, df_rsv, TOTAL_RSV, df_rsv_group)
    df = utils.void_to(df, TOTAL_RSV, 0)
    return df


def gen_plan_sales_rsv(df):
    """Формируем итоговый столбцы плана и продаж"""
    df[PLAN] = np.maximum(df[PLAN_NFE], df[ADJUSTMENT_NFE]).round(0)
    df[SALES] = df[[PLAN, SALES_START_FROM_MONTH]].min(axis=1)
    df = df.rename(columns={TOTAL_RSV: RSV})
    return df


def merge_remains(df):
    """Подтягивает остатки и закупки для расчёта"""
    df_rem = get_data(TABLE_REMAINS)[[
        LINK, FREE_REST, SOFT_HARD_RSV, QUOTA, OVERSTOCK, FULL_REST
    ]]
    df = df.merge(df_rem, on=LINK, how='left')
    for col in [FREE_REST, SOFT_HARD_RSV, QUOTA, OVERSTOCK, FULL_REST]:
        df = utils.void_to(df, col, 0)

    df[FREE_REST_CALC] = np.minimum(
        np.maximum(
            (df[FULL_REST_MONTH] - df[SALES_WHS_START_FROM_MONTH]
             - df[SOFT_HARD_RSV] - df[QUOTA]),
            0
        ),
        df[FREE_REST]
    )
    df[FULL_REST_CALC] = np.minimum(
        np.maximum(df[FULL_REST_MONTH] - df[SALES_WHS_START_FROM_MONTH], 0),
        df[FULL_REST]
    )
    df[OVERSTOCK_CALC] = df[[OVERSTOCK, FULL_REST_CALC]].min(axis=1)
    return df


def merge_group_col(df, group_col, new_name_col):
    """Подтягивает сгруппированный и переименованный столбец """
    group_df = df.merge(
        df[[LINK_PERIOD, group_col]].groupby([LINK_PERIOD]).agg({
            group_col: 'sum'
        }).reset_index().rename(columns={group_col: new_name_col}),
        on=LINK_PERIOD, how='left'
    )
    return group_df


def merge_group_rsv(df):
    """Подтягивает сгруппированный и переименованный столбец по резервам"""
    df_copy = df.copy()
    idx = df_copy[df_copy[NAME_HOLDING].isin([ALL_CLIENTS, NAME_TRAD])].index
    df_copy.loc[idx, RSV_BY_PLAN] = 0
    group_df = df.merge(
        df_copy[[LINK_PERIOD, RSV_BY_PLAN]].groupby([LINK_PERIOD]).agg({
            RSV_BY_PLAN: 'sum'
        }).reset_index().rename(columns={RSV_BY_PLAN: RSV_BY_PLAN_TOTAL}),
        on=LINK_PERIOD, how='left'
    )
    df_copy = df[df[NAME_HOLDING].isin([ALL_CLIENTS, NAME_TRAD])].copy()
    temp_col = RSV_BY_PLAN + '2'
    df_copy = df_copy.groupby([LINK_PERIOD]).agg({
        RSV_BY_PLAN: 'max'
    }).reset_index().rename(columns={RSV_BY_PLAN: temp_col})
    group_df = group_df.merge(
        df_copy[[LINK_PERIOD, temp_col]], on=LINK_PERIOD, how='left'
    )
    idx = group_df[~group_df[temp_col].isnull()].index
    group_df.loc[idx, RSV_BY_PLAN_TOTAL] = np.maximum(
        group_df.loc[idx, RSV_BY_PLAN_TOTAL], group_df.loc[idx, temp_col]
    )
    group_df = group_df.drop(columns=[temp_col], axis=1)
    return group_df


def not_sold(df):
    """Расчёт показателей недогруза"""
    df[LINK_PERIOD] = df[LINK] + df[MONTH]
    df[PLAN_MINUS_SALES] = np.maximum(df[PLAN] - df[SALES], 0)
    df[RSV_BY_PLAN] = df[[RSV, PLAN_MINUS_SALES]].min(axis=1)
    df = merge_group_rsv(df)
    df[RSV_BY_PLAN_PURCH] = (
        df[RSV_BY_PLAN] / df[RSV_BY_PLAN_TOTAL]
        * df[[PURCHASES, RSV_BY_PLAN_TOTAL]].min(axis=1)
    ).round(0)
    df = utils.void_to(df, RSV_BY_PLAN_PURCH, 0)

    df[PLAN_MINUS_SALES_RSV] = np.maximum(df[PLAN] - df[SALES] - df[RSV], 0)
    df = merge_group_col(df, PLAN_MINUS_SALES_RSV, PLAN_MINUS_SALES_RSV_TOTAL)
    df[FREE_REST_BY_PLAN] = (
        df[PLAN_MINUS_SALES_RSV] / df[PLAN_MINUS_SALES_RSV_TOTAL]
        * df[[PLAN_MINUS_SALES_RSV_TOTAL, FREE_REST_CALC]].min(axis=1)
    ).round(0)
    df = utils.void_to(df, FREE_REST_BY_PLAN, 0)
    df = merge_group_col(df, FREE_REST_BY_PLAN, FREE_REST_BY_PLAN_TOTAL)
    df[FREE_REST_BY_PLAN_PURCH] = (
        df[FREE_REST_BY_PLAN] / df[FREE_REST_BY_PLAN_TOTAL]
        * df[[FREE_REST_BY_PLAN_TOTAL, PURCHASES]].min(axis=1)
    ).round(0)
    df = utils.void_to(df, FREE_REST_BY_PLAN_PURCH, 0)

    df[NOT_SOLD] = df[RSV_BY_PLAN] + df[FREE_REST_BY_PLAN]
    df[NOT_SOLD_PURCH] = df[RSV_BY_PLAN_PURCH] + df[FREE_REST_BY_PLAN_PURCH]
    return df


# def responsibility_for_overstock(df):
#     """Расчёт вклад в пересток по заказчикам"""
#     df = merge_group_col(df, NOT_SOLD, NOT_SOLD_TOTAL)
#     df = merge_group_col(df, NOT_SOLD_PURCH, NOT_SOLD_PURCH_TOTAL)
#     df[RES_FOR_STOCK] = (
#         df[NOT_SOLD] / df[NOT_SOLD_TOTAL]
#         * df[[NOT_SOLD_TOTAL, OVERSTOCK]].min(axis=1)
#     ).round(0)
#     df = utils.void_to(df, RES_FOR_STOCK, 0)
#     df[RES_FOR_STOCK_PURCH] = (
#         df[NOT_SOLD_PURCH] / df[NOT_SOLD_PURCH_TOTAL]
#         * df[[NOT_SOLD_PURCH_TOTAL, OVERSTOCK]].min(axis=1)
#     ).round(0)
#     df = utils.void_to(df, RES_FOR_STOCK_PURCH, 0)
#     return df


# def merge_directory(df):
#     """Добавляем данные из справочника"""
#     df_dir = get_data(TABLE_DIRECTORY)[[EAN, MSU, BASE_PRICE]]
#     df = df.merge(df_dir, on=EAN, how='left')
#     df = utils.void_to(df, MSU, 0)
#     df = utils.void_to(df, BASE_PRICE, 0)
#     return df


def main():
    df = gen_plan_sales_rsv(merge_reserve(merge_sales(get_factors())))
    df = not_sold(merge_remains(df))
    save_to_excel(REPORT_DIR + REPORT_TRACKING_NOT_SOLD, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
