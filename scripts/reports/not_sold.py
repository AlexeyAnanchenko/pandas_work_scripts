"""
Отчёт по выявлению недогруженных объёмов по факторам

"""
import utils
utils.path_append()

import numpy as np
import pandas as pd

from service import save_to_excel, get_data, print_complete
from settings import REPORT_NOT_SOLD_PAST, REPORT_DIR, TABLE_FACTORS, PAST
from settings import CURRENT, FACTOR_PERIOD, PURPOSE_PROMO, INACTIVE_PURPOSE
from settings import SALES_FACTOR_PERIOD, LINK, LINK_HOLDING, WHS, NAME_HOLDING
from settings import EAN, PRODUCT, LEVEL_3, PLAN_NFE, ADJUSTMENT_PBI, FACT_NFE
from settings import DESCRIPTION, USER, SALES_CURRENT_FOR_PAST, TABLE_REMAINS
from settings import RSV_FACTOR_PERIOD_TOTAL, TABLE_PURCHASES, FREE_REST
from settings import ARCHIVE_DIR, FULL_REST, SOFT_HARD_RSV, QUOTA, OVERSTOCK
from settings import TABLE_SALES, TABLE_DIRECTORY, MSU, BASE_PRICE
from settings import REPORT_NOT_SOLD_CURRENT


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
    """Получаем начальные данные для отчёта"""
    df = get_data(TABLE_FACTORS)
    df = df[
        (df[FACTOR_PERIOD].isin([PAST, CURRENT]))
        & (df[PURPOSE_PROMO] != INACTIVE_PURPOSE)
    ]
    df[LINK_HOLDING_PERIOD] = df[LINK_HOLDING] + df[FACTOR_PERIOD]
    group_df = df.groupby([
        LINK_HOLDING_PERIOD, LINK, LINK_HOLDING, FACTOR_PERIOD, WHS,
        NAME_HOLDING, EAN, PRODUCT, LEVEL_3
    ]).agg({
        PLAN_NFE: 'max',
        ADJUSTMENT_PBI: 'max',
        FACT_NFE: 'max',
        SALES_FACTOR_PERIOD: 'max',
        SALES_CURRENT_FOR_PAST: 'max',
        RSV_FACTOR_PERIOD_TOTAL: 'max'
    }).reset_index()
    df = df[[
        LINK_HOLDING_PERIOD, DESCRIPTION, USER
    ]].drop_duplicates(subset=LINK_HOLDING_PERIOD)
    group_df = group_df.merge(df, on=LINK_HOLDING_PERIOD, how='left')
    return group_df


def gen_plan_sales_rsv(df):
    """Формируем итоговый столбец плана и столбцы продаж"""
    df[PLAN] = np.maximum(
        df[PLAN_NFE],
        (df[ADJUSTMENT_PBI] * 0.7)
    ).round(0)
    df[SALES_TOTAL] = df[SALES_FACTOR_PERIOD] + df[SALES_CURRENT_FOR_PAST]
    df[SALES] = df[[PLAN, SALES_TOTAL]].min(axis=1)
    df[RSV] = df[RSV_FACTOR_PERIOD_TOTAL]
    return df


def get_archive_remains():
    """Возвращает архивный файл с остатками"""
    from glob import glob
    from datetime import date
    from calendar import monthrange

    archive_rem = None
    files = glob(ARCHIVE_DIR + '/**/*', recursive=True)
    files = [file for file in files if TABLE_REMAINS in file]
    today = date.today()
    num_days = monthrange(today.year, today.month)[1]
    flag = True
    count = 1
    while flag and count <= num_days:
        first_day = today.replace(day=count)
        first_day_string = first_day.strftime('%d.%m.%Y')
        for file in files:
            if first_day_string + ' ' + TABLE_REMAINS in file:
                flag = False
                archive_rem = file
        count += 1
    if archive_rem is not None:
        return pd.ExcelFile(archive_rem).parse()
    return get_data(TABLE_REMAINS)


def merge_remains_purchases(df):
    """Подтягивает остатки и закупки для расчёта"""
    df_rem = get_data(TABLE_REMAINS)[[
        LINK, FREE_REST, SOFT_HARD_RSV, QUOTA, OVERSTOCK
    ]]
    df_rem = df_rem.rename(columns={FREE_REST: FREE_REST_CURRENT})
    df_purch, col_purch = get_data(TABLE_PURCHASES)
    df_purch = df_purch[[LINK, col_purch['pntm'], col_purch['last']]]
    df_purch = df_purch.rename(columns={
        col_purch['pntm']: PURCH_PAST,
        col_purch['last']: PURCH_CURRENT
    })
    df = df.merge(df_rem, on=LINK, how='left')
    df = df.merge(df_purch, on=LINK, how='left')

    df_rem_archive = get_archive_remains()[[LINK, FULL_REST]]
    df = df.merge(df_rem_archive, on=LINK, how='left')
    sales, col_sales = get_data(TABLE_SALES)
    df = df.merge(sales[[LINK, col_sales['last_sale']]], on=LINK, how='left')

    list_col = [
        FREE_REST_CURRENT, PURCH_PAST, PURCH_CURRENT,
        FULL_REST, SOFT_HARD_RSV, QUOTA, OVERSTOCK
    ]
    for col in list_col:
        df = utils.void_to(df, col, 0)

    df[FREE_REST_PAST] = np.minimum(
        np.maximum(
            (df[FULL_REST] - df[col_sales['last_sale']]
             - df[SOFT_HARD_RSV] - df[QUOTA]),
            0
        ),
        df[FREE_REST_CURRENT]
    )
    df = utils.void_to(df, FREE_REST_PAST, 0)
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


def not_sold(df, purch, rest):
    """Рассчёт показателей недогруза"""
    df[LINK_PERIOD] = df[LINK] + df[FACTOR_PERIOD]
    df[PLAN_MINUS_SALES] = np.maximum(df[PLAN] - df[SALES], 0)
    df[RSV_BY_PLAN] = df[[RSV, PLAN_MINUS_SALES]].min(axis=1)
    df = merge_group_col(df, RSV_BY_PLAN, RSV_BY_PLAN_TOTAL)
    df[RSV_BY_PLAN_PURCH] = (
        df[RSV_BY_PLAN] / df[RSV_BY_PLAN_TOTAL]
        * df[[purch, RSV_BY_PLAN_TOTAL]].min(axis=1)
    ).round(0)
    df = utils.void_to(df, RSV_BY_PLAN_PURCH, 0)

    df[PLAN_MINUS_SALES_RSV] = np.maximum(df[PLAN] - df[SALES] - df[RSV], 0)
    df = merge_group_col(df, PLAN_MINUS_SALES_RSV, PLAN_MINUS_SALES_RSV_TOTAL)
    df[FREE_REST_BY_PLAN] = (
        df[PLAN_MINUS_SALES_RSV] / df[PLAN_MINUS_SALES_RSV_TOTAL]
        * df[[PLAN_MINUS_SALES_RSV_TOTAL, rest]].min(axis=1)
    ).round(0)
    df = utils.void_to(df, FREE_REST_BY_PLAN, 0)
    df = merge_group_col(df, FREE_REST_BY_PLAN, FREE_REST_BY_PLAN_TOTAL)
    df[FREE_REST_BY_PLAN_PURCH] = (
        df[FREE_REST_BY_PLAN] / df[FREE_REST_BY_PLAN_TOTAL]
        * df[[FREE_REST_BY_PLAN_TOTAL, purch]].min(axis=1)
    ).round(0)
    df = utils.void_to(df, FREE_REST_BY_PLAN_PURCH, 0)

    df[NOT_SOLD] = df[RSV_BY_PLAN] + df[FREE_REST_BY_PLAN]
    df[NOT_SOLD_PURCH] = df[RSV_BY_PLAN_PURCH] + df[FREE_REST_BY_PLAN_PURCH]
    return df


def responsibility_for_overstock(df):
    """Расчёт вклад в пересток по заказчикам"""
    df = merge_group_col(df, NOT_SOLD, NOT_SOLD_TOTAL)
    df = merge_group_col(df, NOT_SOLD_PURCH, NOT_SOLD_PURCH_TOTAL)
    df[RES_FOR_STOCK] = (
        df[NOT_SOLD] / df[NOT_SOLD_TOTAL]
        * df[[NOT_SOLD_TOTAL, OVERSTOCK]].min(axis=1)
    ).round(0)
    df = utils.void_to(df, RES_FOR_STOCK, 0)
    df[RES_FOR_STOCK_PURCH] = (
        df[NOT_SOLD_PURCH] / df[NOT_SOLD_PURCH_TOTAL]
        * df[[NOT_SOLD_PURCH_TOTAL, OVERSTOCK]].min(axis=1)
    ).round(0)
    df = utils.void_to(df, RES_FOR_STOCK_PURCH, 0)
    return df


def merge_directory(df):
    """Добавляем данные из справочника"""
    df_dir = get_data(TABLE_DIRECTORY)[[EAN, MSU, BASE_PRICE]]
    df = df.merge(df_dir, on=EAN, how='left')
    df = utils.void_to(df, MSU, 0)
    df = utils.void_to(df, BASE_PRICE, 0)
    return df


def get_current_df(df):
    """Отрабатывает скрипт по факторам текущего периода"""
    df = df[df[FACTOR_PERIOD] == CURRENT].copy()
    df = not_sold(df, PURCH_CURRENT, FREE_REST_CURRENT)
    df = merge_directory(responsibility_for_overstock(df))
    df = df[[
        LINK, LINK_HOLDING, WHS, NAME_HOLDING, EAN, PRODUCT, LEVEL_3,
        PLAN_NFE, ADJUSTMENT_PBI, FACT_NFE, DESCRIPTION, USER, PLAN,
        SALES_TOTAL, SALES, RSV, FREE_REST_CURRENT, PURCH_CURRENT,
        PLAN_MINUS_SALES, RSV_BY_PLAN, RSV_BY_PLAN_PURCH, PLAN_MINUS_SALES_RSV,
        FREE_REST_BY_PLAN, FREE_REST_BY_PLAN_PURCH, NOT_SOLD, NOT_SOLD_PURCH,
        OVERSTOCK, RES_FOR_STOCK, RES_FOR_STOCK_PURCH, MSU, BASE_PRICE
    ]]
    return df


def get_past_df(df):
    """Отрабатывает скрипт по факторам прошедшего периода"""
    df = df[df[FACTOR_PERIOD] == PAST].copy()
    df = not_sold(df, PURCH_PAST, FREE_REST_PAST)
    df = merge_directory(responsibility_for_overstock(df))
    df = df[[
        LINK, LINK_HOLDING, WHS, NAME_HOLDING, EAN, PRODUCT, LEVEL_3,
        PLAN_NFE, ADJUSTMENT_PBI, FACT_NFE, DESCRIPTION, USER, PLAN,
        SALES_TOTAL, SALES, RSV, FREE_REST_PAST, PURCH_PAST,
        PLAN_MINUS_SALES, RSV_BY_PLAN, RSV_BY_PLAN_PURCH, PLAN_MINUS_SALES_RSV,
        FREE_REST_BY_PLAN, FREE_REST_BY_PLAN_PURCH, NOT_SOLD, NOT_SOLD_PURCH,
        OVERSTOCK, RES_FOR_STOCK, RES_FOR_STOCK_PURCH, MSU, BASE_PRICE
    ]]
    return df


def main():
    df = merge_remains_purchases(gen_plan_sales_rsv(get_factors()))
    save_to_excel(REPORT_DIR + REPORT_NOT_SOLD_PAST, get_past_df(df))
    save_to_excel(REPORT_DIR + REPORT_NOT_SOLD_CURRENT, get_current_df(df))
    print_complete(__file__)


if __name__ == "__main__":
    main()
