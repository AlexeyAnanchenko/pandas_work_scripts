"""
Формирует отчёт по максимальным потенциальным продажам текущего месяца

"""
import numpy as np
import pandas as pd

import utils
utils.path_append()

from service import save_to_excel, get_data, print_complete
from settings import REPORT_POTENTIAL_SALES, REPORT_DIR, TABLE_FACTORS
from settings import FACTOR_PERIOD, FACTOR_STATUS, CURRENT, PURPOSE_PROMO
from settings import LINK, LINK_HOLDING, FACTOR, DATE_START, DATE_EXPIRATION
from settings import WHS, NAME_HOLDING, EAN, PRODUCT, LEVEL_3, FREE_REST
from settings import USER, PLAN_NFE, SALES_FACTOR_PERIOD, MSU, DATE_CREATION
from settings import RSV_FACTOR_PERIOD, TABLE_SALES_HOLDINGS, TABLE_RESERVE
from settings import SOFT_HARD_RSV, NAME_TRAD, ALL_CLIENTS, TABLE_DIRECTORY
from settings import ELB_PRICE, TABLE_REMAINS, TRANZIT, REPORT_ELBRUS_FACTORS
from hidden_settings import WHS_POTENCTIAL_SALES, elbrus


ACTIVE_STATUS = [
    'Полностью согласован(а)',
    'Завершен(а)',
    'Частично согласован(а)'
]
INACTIVE_PURPOSE = 'Минимизация потерь'
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


def get_factors():
    df = get_data(TABLE_FACTORS)
    df = df[
        (df[FACTOR_PERIOD] == CURRENT)
        & (df[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df[PURPOSE_PROMO] != INACTIVE_PURPOSE)
    ]
    df = df[[
        LINK, LINK_HOLDING, FACTOR, DATE_START, DATE_CREATION,
        DATE_EXPIRATION, WHS, NAME_HOLDING, EAN, USER, PLAN_NFE,
        SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD
    ]]
    return df


def merge_sales_rsv(df):
    sales, col = get_data(TABLE_SALES_HOLDINGS)
    rsv = get_data(TABLE_RESERVE)
    merge_col = [LINK, LINK_HOLDING, WHS, NAME_HOLDING, EAN]
    data_list = [
        [sales, col['last_sale'], SALES_FACTOR_PERIOD],
        [rsv, SOFT_HARD_RSV, RSV_FACTOR_PERIOD]
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

    col_null = [PLAN_NFE, SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD]
    for col in col_null:
        df.loc[df[col].isnull(), col] = 0

    df = df[df[NAME_HOLDING] != ALIDI_MOVING]
    mult_clients = [NAME_TRAD, ALL_CLIENTS]
    clients = list(set(df[NAME_HOLDING].to_list()))
    mult_clients.extend([i for i in clients if '), ' in str(i)])
    df[PLAN_MINUS_FACT] = (df[PLAN_NFE]
                           - df[SALES_FACTOR_PERIOD]
                           - df[RSV_FACTOR_PERIOD])
    df.loc[df[PLAN_MINUS_FACT] < 0, PLAN_MINUS_FACT] = 0
    idx = df[df[NAME_HOLDING].isin(mult_clients)].index
    df.loc[idx, PLAN_NFE] = df.loc[idx, PLAN_MINUS_FACT]
    df.loc[idx, SALES_FACTOR_PERIOD] = 0
    df.loc[idx, RSV_FACTOR_PERIOD] = 0
    df[MAX_DEMAND] = np.maximum(
        df[PLAN_NFE],
        df[SALES_FACTOR_PERIOD] + df[RSV_FACTOR_PERIOD]
    )
    df.loc[idx, RSV_FACTOR_PERIOD] = 0
    df = df[df[MAX_DEMAND] != 0]
    return df


def merge_directory(df):
    direct = get_data(TABLE_DIRECTORY)[[EAN, PRODUCT, LEVEL_3, MSU, ELB_PRICE]]
    df = df.merge(direct, on=EAN, how='left')[[
        LINK, LINK_HOLDING, FACTOR, DATE_CREATION, DATE_START, DATE_EXPIRATION,
        WHS, NAME_HOLDING, EAN, PRODUCT, LEVEL_3, USER, MSU,
        ELB_PRICE, PLAN_NFE, SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD,
        MAX_DEMAND, PLAN_MINUS_FACT
    ]].sort_values(
        by=[DATE_CREATION, DATE_START, FACTOR, WHS, NAME_HOLDING],
        ascending=[True, True, False, True, True])
    df.loc[df[ELB_PRICE].isnull(), ELB_PRICE] = 0
    return df


def merge_remains(df):
    remains = get_data(TABLE_REMAINS)[[LINK, FREE_REST, TRANZIT]]
    df = df.merge(remains, on=LINK, how='left')
    df.loc[df[TRANZIT].isnull(), TRANZIT] = 0
    df.loc[df[FREE_REST].isnull(), FREE_REST] = 0
    df[RESULT_REST] = df[FREE_REST] + df[TRANZIT]
    return df


def process_distribution(df, col_dis, result_col):
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
    idx = df[df[result_col].isnull()].index
    df.loc[idx, result_col] = np.minimum(
        df.loc[idx, PLAN_MINUS_FACT], df.loc[idx, dist_col]
    )
    return df


def distribute_remainder(df):
    dif_df = df.groupby([LINK]).agg({PLAN_MINUS_FACT: 'sum'}).rename(
        columns={PLAN_MINUS_FACT: TOTAL_PLAN_MINUS_FACT}
    )
    df = df.merge(dif_df, on=LINK, how='left')
    df.loc[df[TRANZIT].isnull(), TOTAL_PLAN_MINUS_FACT] = 0
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


def main():
    df = merge_remains(merge_directory(merge_sales_rsv(get_factors())))
    df = distribute_remainder(df)
    save_to_excel(REPORT_DIR + REPORT_POTENTIAL_SALES, df)
    elb_df = df[df[TERRITORY].isin([elbrus])]
    elb_df = elb_df.drop(
        labels=[
            LINK_DATE, TERRITORY, TOTAL_PLAN_MINUS_FACT, DIST_RESULT_REST,
            DIST_FREE_REST
        ],
        axis=1
    )
    save_to_excel(REPORT_DIR + REPORT_ELBRUS_FACTORS, elb_df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
