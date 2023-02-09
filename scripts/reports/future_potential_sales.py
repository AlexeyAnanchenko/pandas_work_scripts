"""
Формирует отчёт по максимальным потенциальным продажам текущего месяца

"""
import numpy as np
import pandas as pd

import utils
utils.path_append()

from service import save_to_excel, get_data, print_complete
from service import get_mult_clients_dict
from settings import FUTURE_ELB_PS, REPORT_DIR, TABLE_FACTORS, FUTURE_BASE_PS
from settings import FACTOR_PERIOD, FACTOR_STATUS, PURPOSE_PROMO, FUTURE
from settings import LINK, LINK_HOLDING, FACTOR, DATE_START, DATE_EXPIRATION
from settings import WHS, NAME_HOLDING, EAN, PRODUCT, LEVEL_3, FREE_REST
from settings import USER, PLAN_NFE, SALES_FACTOR_PERIOD, MSU, DATE_CREATION
from settings import RSV_FACTOR_PERIOD, TABLE_RESERVE, SOFT_HARD_RSV_FUTURE
from settings import SOFT_HARD_RSV, NAME_TRAD, ALL_CLIENTS, TABLE_DIRECTORY
from settings import ELB_PRICE, TABLE_REMAINS, TRANZIT_NEXT, ACTIVE_STATUS
from settings import FUTURE_REPORT_PS, INACTIVE_PURPOSE, BASE_PRICE
from settings import RSV_FACTOR_PERIOD_CURRENT, TABLE_LINES, LINES
from hidden_settings import WHS_POTENCTIAL_SALES, elbrus


FACTOR_SALES = 'Продажи'
ALIDI_MOVING = 'Alidi Межфилиальные продажи (80000000)'
PLAN_MINUS_FACT = 'План - Факт, шт'
TOTAL_PLAN_MINUS_FACT = 'Сумма план - факт по Склад-EAN, шт'
MAX_DEMAND = 'Максимальная потребность, шт'
RESULT_REST = 'Остаток + транзит, шт'
DIST_FREE_REST = 'Свободный остаток к распределению, шт'
DIST_RESULT_REST = 'Остаток + транзит к распределению, шт'
LINK_DATE = 'Сцепка с датами'
FULL_REST_CUSTOMER = 'Остаток + Транзит под заказчика, шт'
FREE_REST_CUSTOMER = 'Свободный остаток под заказчика, шт'
TRANZIT_CUSTOMER = 'Транзит под заказчика, шт'
MAX_REQUIREMENTS = 'Максимальная потребность с учётом стока, шт'
TO_ORDER = 'К дозаказу, шт'
TERRITORY = 'Территория'
RSV_FACTOR_PERIOD_FUTURE = 'Резервы клиента(-ов) будущего месяца'

# ACTIVE_STATUS = ACTIVE_STATUS.append('Не согласован(а)')


def get_factors():
    """Получаем список активных факторов"""
    df = get_data(TABLE_FACTORS)
    df = df[
        (df[FACTOR_PERIOD] == FUTURE)
        & (df[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df[PURPOSE_PROMO] != INACTIVE_PURPOSE)
    ]
    df = df[[
        LINK, LINK_HOLDING, FACTOR, DATE_START, DATE_CREATION,
        DATE_EXPIRATION, WHS, NAME_HOLDING, EAN, USER, PLAN_NFE,
        SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD, RSV_FACTOR_PERIOD_CURRENT
    ]]
    df[RSV_FACTOR_PERIOD_FUTURE] = (df[RSV_FACTOR_PERIOD]
                                    - df[RSV_FACTOR_PERIOD_CURRENT])
    return df


def correction_by_milt_clients(df):
    """Корректировка df из-за наличия мульти-клиентов"""
    all_clients = [NAME_TRAD, ALL_CLIENTS]
    idx = df[df[NAME_HOLDING].isin(all_clients)].index
    df.loc[idx, PLAN_NFE] = df.loc[idx, PLAN_MINUS_FACT]
    df.loc[idx, RSV_FACTOR_PERIOD] = 0
    df.loc[idx, RSV_FACTOR_PERIOD_FUTURE] = 0

    mult_clients = get_mult_clients_dict(df, NAME_HOLDING)
    for clients, values in mult_clients.items():
        list_link = list(df.loc[df[NAME_HOLDING] == clients, LINK].tolist())
        idx = df.loc[
            (df[LINK].isin(list_link)) & (df[NAME_HOLDING].isin(values))
        ].index
        df = df.drop(idx)
        idx = df.loc[df[NAME_HOLDING].isin(values)].index
        df_mult = df.loc[idx]
        df = df.drop(idx)
        replace_dict = {}
        for val in values:
            replace_dict[val] = clients
        df_mult = df_mult.replace({NAME_HOLDING: replace_dict})
        static_col = [LINK, WHS, NAME_HOLDING, EAN]
        agg_col = {
            PLAN_NFE: 'sum',
            SALES_FACTOR_PERIOD: 'sum',
            RSV_FACTOR_PERIOD: 'sum',
            RSV_FACTOR_PERIOD_FUTURE: 'sum',
            PLAN_MINUS_FACT: 'sum'
        }
        df_mult = df_mult.groupby(static_col).agg(agg_col).reset_index()
        df_mult.insert(
            0, LINK_HOLDING,
            df_mult[WHS] + df_mult[NAME_HOLDING] + df_mult[EAN].map(str)
        )
        df = pd.concat([df, df_mult], ignore_index=True)
    return df


def merge_sales_rsv(df):
    """Присоединяем продажи и резервы"""
    rsv = get_data(TABLE_RESERVE)
    merge_col = [LINK, LINK_HOLDING, WHS, NAME_HOLDING, EAN]
    data_list = [
        [rsv, SOFT_HARD_RSV, RSV_FACTOR_PERIOD],
        [rsv, SOFT_HARD_RSV_FUTURE, RSV_FACTOR_PERIOD_FUTURE]
    ]

    for data in data_list:
        merge_df = data[0]
        added_col = data[1]
        result_col = data[2]
        merge_df = merge_df[merge_df[added_col] != 0]
        df = df.merge(
            merge_df[merge_col + [added_col]],
            on=merge_col, how='outer'
        )
        idx = df[df[result_col].isnull()].index
        df.loc[idx, result_col] = df.loc[idx, added_col]
        df.drop(added_col, axis=1, inplace=True)

    col_null = [
        PLAN_NFE, SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD,
        RSV_FACTOR_PERIOD_FUTURE
    ]
    for col in col_null:
        df.loc[df[col].isnull(), col] = 0

    df = df[df[NAME_HOLDING] != ALIDI_MOVING]
    df[PLAN_MINUS_FACT] = (df[PLAN_NFE]
                           - df[SALES_FACTOR_PERIOD]
                           - df[RSV_FACTOR_PERIOD])
    df.loc[df[PLAN_MINUS_FACT] < 0, PLAN_MINUS_FACT] = 0

    df = correction_by_milt_clients(df)

    df[MAX_DEMAND] = np.maximum(
        df[PLAN_NFE],
        df[SALES_FACTOR_PERIOD] + df[RSV_FACTOR_PERIOD]
    )
    df = df[df[MAX_DEMAND] != 0]
    return df


def merge_directory(df):
    """Присоединяем данные из справочника по ШК"""
    direct = get_data(TABLE_DIRECTORY)[[
        EAN, PRODUCT, LEVEL_3, MSU, ELB_PRICE, BASE_PRICE
    ]]
    df = df.merge(direct, on=EAN, how='left')[[
        LINK, LINK_HOLDING, FACTOR, DATE_CREATION, DATE_START, DATE_EXPIRATION,
        WHS, NAME_HOLDING, EAN, PRODUCT, LEVEL_3, USER, MSU, ELB_PRICE,
        BASE_PRICE, PLAN_NFE, SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD,
        RSV_FACTOR_PERIOD_FUTURE, MAX_DEMAND, PLAN_MINUS_FACT
    ]].sort_values(
        by=[DATE_CREATION, DATE_START, FACTOR, WHS, NAME_HOLDING],
        ascending=[True, True, False, True, True])
    return df


def merge_remains(df):
    """Присоединяем остатки и транзиты"""
    remains = get_data(TABLE_REMAINS)[[LINK, FREE_REST, TRANZIT_NEXT]]
    df = df.merge(remains, on=LINK, how='left')
    df.loc[df[TRANZIT_NEXT].isnull(), TRANZIT_NEXT] = 0
    df.loc[df[FREE_REST].isnull(), FREE_REST] = 0
    df[RESULT_REST] = df[FREE_REST] + df[TRANZIT_NEXT]
    return df


def process_distribution(df, col_dis, result_col):
    """Процесс распределения количества по строкам"""
    result_df = pd.DataFrame(columns=[LINK_DATE, result_col])
    start_df = df.copy()

    flag = True
    while flag:
        mid_df = start_df.copy().drop_duplicates(subset=LINK)
        mid_df[result_col] = np.minimum(
            mid_df[PLAN_MINUS_FACT],
            mid_df[col_dis]
        )
        add_col = 'new'
        mid_df[add_col] = np.maximum(
            mid_df[col_dis] - mid_df[PLAN_MINUS_FACT], 0
        )
        mid_df = mid_df.drop(labels=[col_dis], axis=1)
        mid_df[col_dis] = mid_df[add_col]
        mid_df = mid_df.drop(labels=[add_col], axis=1)
        start_df = start_df.drop(labels=[col_dis], axis=1)
        start_df = start_df.merge(mid_df[[LINK, col_dis]], on=LINK, how='left')
        result_df = pd.concat(
            [result_df, mid_df[[LINK_DATE, result_col]]],
            ignore_index=True
        )
        row_drop = mid_df[LINK_DATE].to_list()
        start_df = start_df[~start_df.isin(row_drop).any(axis=1)]
        if start_df.shape[0] == 0:
            flag = False

    return df.merge(result_df, on=LINK_DATE, how='left')


def add_distribute(df, dist_col, result_col):
    """Добавляет колонку с распределением по заказчикам"""
    mid_df = df[
        (df[PLAN_MINUS_FACT] != 0)
        & (df[dist_col] != 0)
        & (df[TOTAL_PLAN_MINUS_FACT] > df[dist_col])
    ][[LINK_DATE, LINK, PLAN_MINUS_FACT, dist_col]]
    mid_df = process_distribution(mid_df, dist_col, result_col)
    df = df.merge(
        mid_df[[LINK_DATE, result_col]],
        on=LINK_DATE, how='left'
    )
    if df[result_col].isna().all():
        df = df.drop(labels=[result_col], axis=1)
        df[result_col] = np.minimum(df[PLAN_MINUS_FACT], df[dist_col])
    else:
        idx = df[df[result_col].isnull()].index
        df.loc[idx, result_col] = np.minimum(
            df.loc[idx, PLAN_MINUS_FACT], df.loc[idx, dist_col]
        )
    return df


def distribute_remainder(df):
    """Общая функция по распределению остатка и транзита на заказчиков"""
    dif_df = df.groupby([LINK]).agg({PLAN_MINUS_FACT: 'sum'}).rename(
        columns={PLAN_MINUS_FACT: TOTAL_PLAN_MINUS_FACT}
    )
    df = df.merge(dif_df, on=LINK, how='left')
    df[DIST_RESULT_REST] = np.minimum(
        df[RESULT_REST],
        df[TOTAL_PLAN_MINUS_FACT]
    )
    df[DIST_FREE_REST] = np.minimum(df[FREE_REST], df[TOTAL_PLAN_MINUS_FACT])
    df.insert(
        0, LINK_DATE,
        df[LINK_HOLDING] + df[DATE_CREATION].map(str) + df[DATE_START].map(str)
    )
    df = add_distribute(df, DIST_RESULT_REST, FULL_REST_CUSTOMER)
    df = add_distribute(df, DIST_FREE_REST, FREE_REST_CUSTOMER)
    df[TRANZIT_CUSTOMER] = df[FULL_REST_CUSTOMER] - df[FREE_REST_CUSTOMER]
    df[MAX_REQUIREMENTS] = (df[SALES_FACTOR_PERIOD]
                            + df[RSV_FACTOR_PERIOD] + df[FULL_REST_CUSTOMER])
    df[TO_ORDER] = df[PLAN_MINUS_FACT] - df[FULL_REST_CUSTOMER]
    df.insert(3, TERRITORY, df[WHS])
    df = df.replace({TERRITORY: WHS_POTENCTIAL_SALES})
    return df


def add_lines(df):
    df_lines = get_data(TABLE_LINES)
    df = df.merge(df_lines, on=EAN, how='left')
    df = utils.void_to(df, LINES, 'БЕЗ ЛИНЕЙКИ')
    return df


def get_elbrus_factors(df):
    elb_df = df[df[TERRITORY].isin([elbrus])]
    elb_df = elb_df.drop(
        labels=[
            LINK_DATE, TERRITORY, TOTAL_PLAN_MINUS_FACT, DIST_RESULT_REST,
            DIST_FREE_REST
        ],
        axis=1
    )
    elb_df = add_lines(elb_df)
    return elb_df


def get_base_factors(df):
    base_df = df[df[TERRITORY] != elbrus]
    base_df = base_df.drop(
        labels=[
            LINK_DATE, TERRITORY, TOTAL_PLAN_MINUS_FACT, DIST_RESULT_REST,
            DIST_FREE_REST
        ],
        axis=1
    )
    list_holding = base_df[base_df[PLAN_NFE] > 0][NAME_HOLDING].tolist()
    base_df = base_df[base_df[NAME_HOLDING].isin(list_holding)]
    return base_df


def main():
    df = merge_remains(merge_directory(merge_sales_rsv(get_factors())))
    df = distribute_remainder(df)
    save_to_excel(REPORT_DIR + FUTURE_REPORT_PS, df)
    save_to_excel(
        REPORT_DIR + FUTURE_ELB_PS,
        get_elbrus_factors(df)
    )
    save_to_excel(
        REPORT_DIR + FUTURE_BASE_PS,
        get_base_factors(df)
    )
    print_complete(__file__)


if __name__ == "__main__":
    main()
