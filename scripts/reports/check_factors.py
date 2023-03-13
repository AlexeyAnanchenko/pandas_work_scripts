"""
Отчёт по проверке факторов согласованных и новых

"""
import utils
utils.path_append()

import pandas as pd
import numpy as np

from hidden_settings import WHS_ELBRUS
from service import save_to_excel, get_data, print_complete
from settings import REPORT_CHECK_FACTORS, REPORT_DIR, TABLE_FACTORS, PRODUCT
from settings import FACTOR_PERIOD, CURRENT, PURPOSE_PROMO, INACTIVE_PURPOSE
from settings import LINK, FUTURE, NAME_HOLDING, RSV_FACTOR_PERIOD_CURRENT
from settings import SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD, FACTOR_NUM, WHS
from settings import LINK_HOLDING, RSV_FACTOR_PERIOD_FUTURE, CUTS_PBI, MSU
from settings import ADJUSTMENT_PBI, FACT_NFE, NAME_TRAD, TABLE_EXCLUDE
from settings import FACTOR_STATUS, ACTIVE_STATUS, SALES_PBI, RESERVES_PBI
from settings import TABLE_ASSORTMENT, TABLE_DIRECTORY, EAN, MATRIX, MATRIX_LY
from settings import ELB_PRICE, BASE_PRICE, PLAN_NFE, ALL_CLIENTS, FULL_REST
from settings import AVG_FACTOR_PERIOD, TABLE_REMAINS, FREE_REST, TRANZIT
from settings import OVERSTOCK, AVG_FACTOR_PERIOD_WHS, SOFT_HARD_RSV, QUOTA
from settings import TABLE_SALES, DATE_EXPIRATION, DATE_START, REF_FACTOR
from settings import DESCRIPTION, USER, FACTOR, DATE_CREATION, LEVEL_3
from settings import CANCEL_STATUS


LINK_HOLDING_PERIOD = 'Сцепка Период-Склад-Холдинг-ШК'
LINK_FACTOR = 'Сцепка Номер фактора-Склад-Штрихкод'
LINK_PERIOD = 'Сцепка Склад-ШК-Период'
CHECK_FACT = 'ПРОВЕРКА ФАКТА ТЕКУЩЕГО ПЕРИОДА'
CHECK_DUPL = 'ПРОВЕРКА НАЛИЧИЯ ДУБЛИКАТОВ'
ACTIVE_LOC = 'Активный ассортимент для локации'
PRICE_LOC = 'Цена для локации (GIV/NIV), руб.'
L_YUG = 'Лоджистик-Юг ELBR (5553395)'
YES = 'Да'
NO = 'Нет'
PLAN_MSU = 'План, msu'
PLAN_PRICE = 'План, руб.'
AVARAGE_MSU = 'Средние продажи, msu'
AVARAGE_PRICE = 'Средние продажи, руб.'
RISK = 'Риски в мес.'
PLAN_NFE_TOTAL = 'Общий план в NFE на Склад-ШК'
RISK_TOTAL = 'Риски по общему плану, в мес.'
REST_TRANZIT = 'Общие остатки + Транзит, шт'
DAYS_IN_FACTOR = 'Дней в факторе'
REMAINING_DAYS = 'Остаток дней в факторе'
DAYS_PASSED = 'Прошлой дней в факторе'
NEW_PLAN = 'Новый план, шт'
NEW_RISK = 'Новые риски в мес.'
NEW_RISK_PLAN = 'Новый план с учётом рисков, в шт'
NEW_ADJUSTMENT = 'Новая корректировка прогноза, шт'
NEW_RISK_ADJUST = 'Новая корректировка прогноза с учётом рисков, шт'

df_exclude = get_data(TABLE_EXCLUDE)


def get_factors():
    """Получаем факторы для проверки"""
    df = get_data(TABLE_FACTORS)
    df = df[
        (df[FACTOR_PERIOD].isin([CURRENT, FUTURE]))
        & (df[PURPOSE_PROMO] != INACTIVE_PURPOSE)
        & (df[FACTOR_STATUS] != CANCEL_STATUS)
    ]
    df[RSV_FACTOR_PERIOD_FUTURE] = (df[RSV_FACTOR_PERIOD]
                                    - df[RSV_FACTOR_PERIOD_CURRENT])
    df.insert(0, LINK_FACTOR, df[FACTOR_NUM].map(str) + df[LINK])
    df.insert(0, LINK_HOLDING_PERIOD, df[FACTOR_PERIOD] + df[LINK_HOLDING])
    df.insert(0, LINK_PERIOD, df[FACTOR_PERIOD] + df[LINK])
    df = df.drop(labels=[LINK_HOLDING, PURPOSE_PROMO], axis=1)
    return df


def check_fact(df):
    """Проверяем факт продаж по заявке"""
    df.loc[
        (df[SALES_PBI] + df[RESERVES_PBI] < (df[SALES_FACTOR_PERIOD]
                                             + df[RSV_FACTOR_PERIOD_CURRENT]))
        & (df[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df[FACTOR_PERIOD] == CURRENT)
        & (~df[NAME_HOLDING].isin([NAME_TRAD, ALL_CLIENTS]))
        & (~df[LINK_FACTOR].isin(df_exclude[CHECK_FACT].to_list())),
        CHECK_FACT
    ] = 'Факт ниже продаж и резервов!'
    df.loc[
        (df[FACT_NFE] - df[CUTS_PBI] < (df[SALES_FACTOR_PERIOD]
                                        + df[RSV_FACTOR_PERIOD_CURRENT]))
        & (~df[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df[FACTOR_PERIOD] == CURRENT)
        & (~df[NAME_HOLDING].isin([NAME_TRAD, ALL_CLIENTS]))
        & (~df[LINK_FACTOR].isin(df_exclude[CHECK_FACT].to_list())),
        CHECK_FACT
    ] = 'Факт ниже продаж и резервов!'
    return df


def check_duplicates(df):
    """Проверяем дубликаты по периоду"""
    duplicate_df = df[
        ~df[LINK_FACTOR].isin(df_exclude[CHECK_DUPL].to_list())
    ].copy()
    duplicate_df = duplicate_df[duplicate_df.duplicated(
        [LINK_HOLDING_PERIOD], keep='first'
    )][LINK_HOLDING_PERIOD].to_list()
    df.loc[df[LINK_HOLDING_PERIOD].isin(duplicate_df), CHECK_DUPL] = 'Дубликат'
    return df


def merge_assort_and_dir(df):
    """Подтягиваем ассортимент и данные из справочника по ШК"""
    assort = get_data(TABLE_ASSORTMENT)[[LINK]]
    assort[ACTIVE_LOC] = YES
    df = df.merge(assort, on=LINK, how='left')

    columns = [MATRIX, MATRIX_LY, MSU, ELB_PRICE, BASE_PRICE]
    direct = get_data(TABLE_DIRECTORY)[[EAN] + columns]
    df = df.merge(direct, on=EAN, how='left')
    df.loc[df[ACTIVE_LOC].isnull(), ACTIVE_LOC] = NO
    df.loc[
        (df[NAME_HOLDING] == L_YUG) & (df[MATRIX_LY] == NO), ACTIVE_LOC
    ] = NO
    df.loc[
        (df[NAME_HOLDING] != L_YUG)
        & (df[MATRIX] == NO)
        & (df[WHS].isin(WHS_ELBRUS.keys())),
        ACTIVE_LOC
    ] = NO
    idx = df[df[WHS].isin(WHS_ELBRUS.keys())].index
    df.loc[idx, PRICE_LOC] = df.loc[idx, ELB_PRICE]
    idx = df[~df[WHS].isin(WHS_ELBRUS.keys())].index
    df.loc[idx, PRICE_LOC] = df.loc[idx, BASE_PRICE]

    df[PLAN_MSU] = (df[PLAN_NFE] * df[MSU]).round(4)
    df[PLAN_PRICE] = (df[PLAN_NFE] * df[PRICE_LOC]).round(2)
    df[AVARAGE_MSU] = (df[AVG_FACTOR_PERIOD] * df[MSU]).round(4)
    df[AVARAGE_PRICE] = (df[AVG_FACTOR_PERIOD] * df[PRICE_LOC]).round(2)
    df = df.drop(labels=columns + [PRICE_LOC], axis=1)
    return df


def merge_remains(df):
    """Подтягиваем остатки и резервы"""
    df = df.drop(columns=[QUOTA], axis=1)
    columns = [
        LINK, FULL_REST, SOFT_HARD_RSV, QUOTA, FREE_REST, TRANZIT, OVERSTOCK
    ]
    remainds = get_data(TABLE_REMAINS)[columns]
    df = df.merge(remainds, on=LINK, how='left')
    columns.remove(LINK)
    for col in columns:
        df = utils.void_to(df, col, 0)
    df[REST_TRANZIT] = df[FULL_REST] + df[TRANZIT]
    return df


def risk_calculation(df):
    """Считаем риски для склада"""
    df[RISK] = (df[PLAN_NFE] / df[AVG_FACTOR_PERIOD_WHS]).round(1)
    df.loc[df[AVG_FACTOR_PERIOD_WHS] == 0, RISK] = 9999
    df_group = df.groupby([LINK_PERIOD]).agg(
        {PLAN_NFE: 'sum'}
    ).reset_index()
    df_group = df_group.rename(columns={PLAN_NFE: PLAN_NFE_TOTAL})
    df = df.merge(df_group, on=LINK_PERIOD, how='left')
    df[RISK_TOTAL] = (df[PLAN_NFE_TOTAL] / df[AVG_FACTOR_PERIOD_WHS]).round(1)
    df.loc[df[AVG_FACTOR_PERIOD_WHS] == 0, RISK_TOTAL] = 9999
    return df


def merge_sales(df):
    """Подтягиваем тотал продажи по складу"""
    df_sales, col_sales = get_data(TABLE_SALES)
    df_sales = df_sales[[LINK, col_sales['last_cut'], col_sales['last_sale']]]
    df = df.merge(df_sales, on=LINK, how='left')
    return df


def calc_new_forecast(df):
    """Рассчитываем корректировку прогноза по трекингу"""
    curent_dt = pd.Timestamp(pd.Timestamp.now().date())
    risk_limit = 1
    add_purchase_days = 7

    df[DAYS_IN_FACTOR] = (df[DATE_EXPIRATION] - df[DATE_START]).dt.days
    df[REMAINING_DAYS] = (df[DATE_EXPIRATION] - curent_dt).dt.days
    df.loc[df[REMAINING_DAYS] < 0, REMAINING_DAYS] = 0
    df.loc[
        df[REMAINING_DAYS] > df[DAYS_IN_FACTOR], REMAINING_DAYS
    ] = df[DAYS_IN_FACTOR]
    df[DAYS_PASSED] = df[DAYS_IN_FACTOR] - df[REMAINING_DAYS]
    df[NEW_PLAN] = (np.maximum(
        (df[SALES_PBI] / df[DAYS_PASSED] * df[DAYS_IN_FACTOR])
        + (
            (
                df[CUTS_PBI] / df[DAYS_PASSED]
                * df[DAYS_IN_FACTOR] - df[CUTS_PBI]
            )
            / 2
        )
        + df[CUTS_PBI],
        (df[AVG_FACTOR_PERIOD] / df[DAYS_IN_FACTOR] * df[REMAINING_DAYS])
        + df[SALES_PBI] + df[CUTS_PBI]
    ) + df[RESERVES_PBI]).round(0)
    idx = df[df[NEW_PLAN].isnull()].index
    df.loc[idx, NEW_PLAN] = df.loc[idx, PLAN_NFE]
    idx = df[df[REMAINING_DAYS] == 0].index
    df.loc[idx, NEW_PLAN] = df.loc[idx, FACT_NFE]
    df[NEW_ADJUSTMENT] = df[[PLAN_NFE, ADJUSTMENT_PBI, NEW_PLAN]].max(axis=1)

    df[NEW_RISK] = (df[NEW_PLAN] / df[AVG_FACTOR_PERIOD_WHS]).round(1)
    df.loc[df[AVG_FACTOR_PERIOD_WHS] == 0, NEW_RISK] = 9999
    df[NEW_RISK_PLAN] = df[NEW_PLAN]
    idx = df[
        (df[REMAINING_DAYS] > add_purchase_days) & (df[NEW_RISK] > risk_limit)
    ].index
    df.loc[idx, NEW_RISK_PLAN] = (df[NEW_PLAN] / df[DAYS_IN_FACTOR]
                                  * (df[DAYS_PASSED]
                                     + add_purchase_days)).round(0)
    df[NEW_RISK_ADJUST] = df[[
        PLAN_NFE, ADJUSTMENT_PBI, NEW_RISK_PLAN
    ]].max(axis=1)
    return df


def reindex_and_sort(df):
    df_sales, col_sales = get_data(TABLE_SALES)
    df = df[[
        LINK_FACTOR, REF_FACTOR, DESCRIPTION, USER, FACTOR, FACTOR_PERIOD,
        FACTOR_NUM, FACTOR_STATUS, DATE_START, DATE_EXPIRATION, DATE_CREATION,
        WHS, NAME_HOLDING, EAN, LEVEL_3, PRODUCT, CHECK_FACT, CHECK_DUPL,
        ACTIVE_LOC, PLAN_MSU, PLAN_PRICE, ADJUSTMENT_PBI, PLAN_NFE, SALES_PBI,
        RESERVES_PBI, CUTS_PBI, FACT_NFE, SALES_FACTOR_PERIOD,
        RSV_FACTOR_PERIOD_CURRENT, RSV_FACTOR_PERIOD_FUTURE, RSV_FACTOR_PERIOD,
        AVARAGE_MSU, AVARAGE_PRICE, AVG_FACTOR_PERIOD, AVG_FACTOR_PERIOD_WHS,
        RISK, PLAN_NFE_TOTAL, RISK_TOTAL, col_sales['last_cut'],
        col_sales['last_sale'], SOFT_HARD_RSV, QUOTA, FULL_REST, REST_TRANZIT,
        FREE_REST, OVERSTOCK, TRANZIT, DAYS_IN_FACTOR, DAYS_PASSED,
        REMAINING_DAYS, NEW_PLAN, NEW_ADJUSTMENT, NEW_RISK, NEW_RISK_PLAN,
        NEW_RISK_ADJUST
    ]].sort_values(
        by=[DATE_CREATION, FACTOR_NUM, RISK], ascending=[False, False, False]
    )
    return df


def main():
    df = merge_assort_and_dir(check_duplicates(check_fact(get_factors())))
    df = calc_new_forecast(merge_sales(risk_calculation(merge_remains(df))))
    df = reindex_and_sort(df)
    save_to_excel(REPORT_DIR + REPORT_CHECK_FACTORS, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
