"""
Формирует отчёт по максимальным потенциальным продажам текущего месяца

"""
import numpy as np

import utils
utils.path_append()

from service import save_to_excel, get_data, print_complete
from settings import REPORT_POTENTIAL_SALES, REPORT_DIR, TABLE_FACTORS
from settings import FACTOR_PERIOD, FACTOR_STATUS, CURRENT, PURPOSE_PROMO
from settings import LINK, LINK_HOLDING, FACTOR, DATE_START, DATE_EXPIRATION
from settings import WHS, NAME_HOLDING, EAN, PRODUCT, LEVEL_3, DESCRIPTION
from settings import USER, PLAN_NFE, SALES_FACTOR_PERIOD, MSU, DATE_CREATION
from settings import RSV_FACTOR_PERIOD, TABLE_SALES_HOLDINGS, TABLE_RESERVE
from settings import SOFT_HARD_RSV, NAME_TRAD, ALL_CLIENTS, TABLE_DIRECTORY
from settings import ELB_PRICE, TABLE_REMAINS, TRANZIT, FREE_REST


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
MAX_DEMAND = 'Максимальная потребность'
RESULT_REST = 'Остаток + транзит, шт'
DIST_FREE_REST = 'Свободный остаток к распределению, шт'
DIST_RESULT_REST = 'Остаток + транзит к распределению, шт'
LINK_DATE = 'Сцепка с датами'


def get_factors():
    df = get_data(TABLE_FACTORS)
    df = df[
        (df[FACTOR_PERIOD] == CURRENT)
        & (df[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df[PURPOSE_PROMO] != INACTIVE_PURPOSE)
    ]
    df = df[[
        LINK, LINK_HOLDING, FACTOR, DATE_START, DATE_CREATION,
        DATE_EXPIRATION, WHS, NAME_HOLDING, EAN, DESCRIPTION,
        USER, PLAN_NFE, SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD
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
        WHS, NAME_HOLDING, EAN, PRODUCT, LEVEL_3, DESCRIPTION, USER, MSU,
        ELB_PRICE, PLAN_NFE, SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD,
        MAX_DEMAND, PLAN_MINUS_FACT
    ]].sort_values(
        by=[DATE_CREATION, DATE_START, FACTOR, WHS, NAME_HOLDING],
        ascending=[True, True, False, True, True])
    return df


def merge_remains(df):
    remains = get_data(TABLE_REMAINS)[[LINK, FREE_REST, TRANZIT]]
    df = df.merge(remains, on=LINK, how='left')
    df.loc[df[TRANZIT].isnull(), TRANZIT] = 0
    df.loc[df[FREE_REST].isnull(), FREE_REST] = 0
    df[RESULT_REST] = df[FREE_REST] + df[TRANZIT]
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
    free_rest_df = df[
        (df[PLAN_MINUS_FACT] != 0)
        & (df[DIST_RESULT_REST] != 0)
        & (df[TOTAL_PLAN_MINUS_FACT] > df[DIST_RESULT_REST])
    ]
    # комментарий к строке выше, по остальным берём меньшее из (план - факт и остаток)
    save_to_excel(REPORT_DIR + 'ОТЧЁТИК.xlsx', free_rest_df)
    return df


def main():
    df = merge_remains(merge_directory(merge_sales_rsv(get_factors())))
    df = distribute_remainder(df)
    save_to_excel(REPORT_DIR + REPORT_POTENTIAL_SALES, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
