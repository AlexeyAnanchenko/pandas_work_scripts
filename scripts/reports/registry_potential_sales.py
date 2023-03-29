"""
Формирует отчёт по максимальным потенциальным продажам текущего месяца

"""
import numpy as np
import pandas as pd

import utils
utils.path_append()

from service import save_to_excel, get_data, print_complete
from service import get_mult_clients_dict
from settings import REPORT_POTENTIAL_SALES, REPORT_DIR, TABLE_FACTORS
from settings import FACTOR_PERIOD, FACTOR_STATUS, CURRENT, PURPOSE_PROMO
from settings import LINK, LINK_HOLDING, FACTOR, DATE_START, DATE_EXPIRATION
from settings import WHS, NAME_HOLDING, EAN, PRODUCT, LEVEL_3, AVAILABLE_REST
from settings import USER, PLAN_NFE, SALES_FACTOR_PERIOD, MSU, DATE_CREATION
from settings import RSV_FACTOR_PERIOD, TABLE_SALES_HOLDINGS, TABLE_RESERVE
from settings import SOFT_HARD_RSV, NAME_TRAD, ALL_CLIENTS, TABLE_DIRECTORY
from settings import ELB_PRICE, TABLE_REMAINS, TRANZIT_CURRENT, ACTIVE_STATUS
from settings import REPORT_ELBRUS_FACTORS, INACTIVE_PURPOSE, BASE_PRICE
from settings import REPORT_BASE_FACTORS, TABLE_REGISTRY_FACTORS, DATE_REGISTRY
from settings import QUANT_REGISTRY, FACTOR_NUM, RSV_FACTOR_PERIOD_CURRENT
from settings import SOFT_HARD_RSV_CURRENT, TABLE_LINES, LINES, SALES_BY_DATE
from settings import SOFT_RSV_BY_DATE, HARD_RSV_BY_DATE
from hidden_settings import WHS_ELBRUS, elbrus


FACTOR_SALES = 'Продажи'
ALIDI_MOVING = 'Alidi Межфилиальные продажи (80000000)'
PLAN_MINUS_FACT = 'План - Факт, шт'
TOTAL_PLAN_MINUS_FACT = 'Сумма план - факт по Склад-EAN, шт'
MAX_DEMAND = 'Максимальная потребность, шт'
RESULT_REST = 'Остаток + транзит, шт'
DIST_AVAILABLE_REST = 'Доступный остаток к распределению, шт'
DIST_RESULT_REST = 'Остаток + транзит к распределению, шт'
LINK_DATE = 'Сцепка с датами'
FULL_REST_CUSTOMER = 'Остаток + Транзит под заказчика, шт'
FREE_REST_CUSTOMER = 'Свободный остаток под заказчика, шт'
TRANZIT_CUSTOMER = 'Транзит под заказчика, шт'
MAX_REQUIREMENTS = 'Максимальная потребность с учётом стока, шт'
TO_ORDER = 'К дозаказу, шт'
TERRITORY = 'Территория'
PLAN_MINUS_FACT_ROW = 'План - Факт построчно, шт'


def get_factors():
    """Получаем список активных факторов"""
    df = get_data(TABLE_FACTORS)
    df = df[
        (df[FACTOR_PERIOD] == CURRENT)
        & (df[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df[PURPOSE_PROMO] != INACTIVE_PURPOSE)
    ]
    df = df[[
        LINK, LINK_HOLDING, FACTOR, DATE_START, DATE_CREATION,
        DATE_EXPIRATION, WHS, NAME_HOLDING, EAN, USER, PLAN_NFE,
        SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD, RSV_FACTOR_PERIOD_CURRENT,
        SALES_BY_DATE, SOFT_RSV_BY_DATE, HARD_RSV_BY_DATE
    ]]
    df[PLAN_MINUS_FACT] = (df[PLAN_NFE] - df[SALES_BY_DATE]
                           - df[SOFT_RSV_BY_DATE] - df[HARD_RSV_BY_DATE])
    df.loc[df[PLAN_MINUS_FACT] < 0, PLAN_MINUS_FACT] = 0
    df = df.drop(
        labels=[SALES_BY_DATE, SOFT_RSV_BY_DATE, HARD_RSV_BY_DATE], axis=1
    )
    return df


def correction_by_milt_clients(df):
    """Корректировка df из-за наличия мульти-клиентов"""
    all_clients = [NAME_TRAD, ALL_CLIENTS]
    idx = df[df[NAME_HOLDING].isin(all_clients)].index
    df.loc[idx, PLAN_NFE] = df.loc[idx, PLAN_MINUS_FACT]
    df.loc[idx, SALES_FACTOR_PERIOD] = 0
    df.loc[idx, RSV_FACTOR_PERIOD] = 0
    df.loc[idx, RSV_FACTOR_PERIOD_CURRENT] = 0

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
            RSV_FACTOR_PERIOD_CURRENT: 'sum',
            PLAN_MINUS_FACT: 'sum',
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
    sales, col = get_data(TABLE_SALES_HOLDINGS)
    rsv = get_data(TABLE_RESERVE)
    merge_col = [LINK, LINK_HOLDING, WHS, NAME_HOLDING, EAN]
    data_list = [
        [sales, col['last_sale'], SALES_FACTOR_PERIOD],
        [rsv, SOFT_HARD_RSV, RSV_FACTOR_PERIOD],
        [rsv, SOFT_HARD_RSV_CURRENT, RSV_FACTOR_PERIOD_CURRENT]
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
        RSV_FACTOR_PERIOD_CURRENT
    ]
    for col in col_null:
        df.loc[df[col].isnull(), col] = 0

    df = df[df[NAME_HOLDING] != ALIDI_MOVING]

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
        RSV_FACTOR_PERIOD_CURRENT, MAX_DEMAND, PLAN_MINUS_FACT
    ]]
    return df


def merge_remains(df):
    """Присоединяем остатки и транзиты"""
    remains = get_data(TABLE_REMAINS)[[LINK, AVAILABLE_REST, TRANZIT_CURRENT]]
    df = df.merge(remains, on=LINK, how='left')
    df.loc[df[TRANZIT_CURRENT].isnull(), TRANZIT_CURRENT] = 0
    df.loc[df[AVAILABLE_REST].isnull(), AVAILABLE_REST] = 0
    df[RESULT_REST] = df[AVAILABLE_REST] + df[TRANZIT_CURRENT]
    return df


def process_distribution(df, col_dis, result_col, uniq_col, link_dis):
    """Процесс распределения количества по строкам"""
    result_df = pd.DataFrame(columns=[LINK_DATE, result_col])
    start_df = df.copy()

    flag = True
    count = 0
    while flag:
        mid_df = start_df.copy().drop_duplicates(subset=link_dis)
        mid_df[result_col] = np.minimum(
            mid_df[uniq_col],
            mid_df[col_dis]
        )
        add_col = 'new'
        mid_df[add_col] = np.maximum(
            mid_df[col_dis] - mid_df[uniq_col], 0
        )
        mid_df = mid_df.drop(labels=[col_dis], axis=1)
        mid_df[col_dis] = mid_df[add_col]
        mid_df = mid_df.drop(labels=[add_col], axis=1)
        start_df = start_df.drop(labels=[col_dis], axis=1)
        start_df = start_df.merge(
            mid_df[[link_dis, col_dis]], on=link_dis, how='left'
        )
        result_df = pd.concat(
            [result_df, mid_df[[LINK_DATE, result_col]]],
            ignore_index=True
        )
        row_drop = mid_df[LINK_DATE].to_list()
        start_df = start_df[~start_df.isin(row_drop).any(axis=1)]
        count += 1
        if start_df.shape[0] == 0:
            flag = False

    return df.merge(result_df, on=LINK_DATE, how='left')


def get_registry_factors(df):
    """Функци возвращает реестр по факторам с распределённым 'План - Факт'"""
    df_reg = get_data(TABLE_REGISTRY_FACTORS)
    df_reg = df_reg[df_reg[FACTOR_PERIOD] == CURRENT]
    df_reg.insert(
        0, LINK_HOLDING,
        df_reg[WHS] + df_reg[NAME_HOLDING] + df_reg[EAN].map(str)
    )
    df_reg = df_reg.merge(
        df[[LINK_HOLDING, PLAN_MINUS_FACT]], on=LINK_HOLDING, how='left'
    )
    df_reg.loc[df_reg[PLAN_MINUS_FACT].isnull(), PLAN_MINUS_FACT] = 0
    df_reg.insert(
        0, LINK_DATE, df_reg[LINK_HOLDING] + df_reg[DATE_REGISTRY].map(str)
    )
    df_reg = process_distribution(
        df_reg, PLAN_MINUS_FACT, PLAN_MINUS_FACT_ROW,
        QUANT_REGISTRY, LINK_HOLDING
    )
    df_reg = df_reg.drop(labels=[PLAN_MINUS_FACT], axis=1)
    df_reg = df_reg.sort_values(
        by=[DATE_REGISTRY, DATE_START, FACTOR_NUM],
        ascending=[True, True, True])
    return df_reg


def add_distribute(df, dist_col, result_col):
    """Добавляет колонку с распределением по заказчикам"""
    mid_df = df[
        (df[PLAN_MINUS_FACT_ROW] != 0)
        & (df[dist_col] != 0)
        & (df[TOTAL_PLAN_MINUS_FACT] > df[dist_col])
    ][[LINK_DATE, LINK, PLAN_MINUS_FACT_ROW, dist_col]]
    mid_df = process_distribution(
        mid_df, dist_col, result_col, PLAN_MINUS_FACT_ROW, LINK
    )
    df = df.merge(
        mid_df[[LINK_DATE, result_col]],
        on=LINK_DATE, how='left'
    )
    idx = df[df[result_col].isnull()].index
    df.loc[idx, result_col] = np.minimum(
        df.loc[idx, PLAN_MINUS_FACT_ROW], df.loc[idx, dist_col]
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
    df[DIST_AVAILABLE_REST] = np.minimum(
        df[AVAILABLE_REST], df[TOTAL_PLAN_MINUS_FACT]
    )
    df_reg = get_registry_factors(df)
    df_reg = df_reg.merge(
        df[[
            LINK, DIST_RESULT_REST, DIST_AVAILABLE_REST, TOTAL_PLAN_MINUS_FACT
        ]].drop_duplicates(subset=LINK),
        on=LINK, how='left'
    )
    df_reg = add_distribute(df_reg, DIST_RESULT_REST, FULL_REST_CUSTOMER)
    df_reg = add_distribute(df_reg, DIST_AVAILABLE_REST, FREE_REST_CUSTOMER)
    df_reg = df_reg.groupby([LINK_HOLDING]).agg(
        {FULL_REST_CUSTOMER: 'sum', FREE_REST_CUSTOMER: 'sum'}
    ).reset_index()
    df = df.merge(df_reg, on=LINK_HOLDING, how='left')
    df.loc[df[FULL_REST_CUSTOMER].isnull(), FULL_REST_CUSTOMER] = 0
    df.loc[df[FREE_REST_CUSTOMER].isnull(), FREE_REST_CUSTOMER] = 0
    df[TRANZIT_CUSTOMER] = df[FULL_REST_CUSTOMER] - df[FREE_REST_CUSTOMER]
    df[MAX_REQUIREMENTS] = (df[SALES_FACTOR_PERIOD]
                            + df[RSV_FACTOR_PERIOD] + df[FULL_REST_CUSTOMER])
    df[TO_ORDER] = df[PLAN_MINUS_FACT] - df[FULL_REST_CUSTOMER]
    df.insert(3, TERRITORY, df[WHS])
    df = df.replace({TERRITORY: WHS_ELBRUS})
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
            TERRITORY, TOTAL_PLAN_MINUS_FACT,
            DIST_RESULT_REST, DIST_AVAILABLE_REST
        ],
        axis=1
    )
    elb_df = add_lines(elb_df)
    return elb_df


def get_base_factors(df):
    base_df = df[df[TERRITORY] != elbrus]
    base_df = base_df.drop(
        labels=[
            TERRITORY, TOTAL_PLAN_MINUS_FACT,
            DIST_RESULT_REST, DIST_AVAILABLE_REST
        ],
        axis=1
    )
    list_holding = base_df[base_df[PLAN_NFE] > 0][NAME_HOLDING].tolist()
    base_df = base_df[base_df[NAME_HOLDING].isin(list_holding)]
    return base_df


def main():
    df = merge_remains(merge_directory(merge_sales_rsv(get_factors())))
    df = distribute_remainder(df)
    save_to_excel(
        REPORT_DIR + REPORT_POTENTIAL_SALES.replace('.', ' по реестру.'), df
    )
    save_to_excel(
        REPORT_DIR + REPORT_ELBRUS_FACTORS.replace('.', ' по реестру.'),
        get_elbrus_factors(df)
    )
    save_to_excel(
        REPORT_DIR + REPORT_BASE_FACTORS.replace('.', ' по реестру.'),
        get_base_factors(df)
    )
    print_complete(__file__)


if __name__ == "__main__":
    main()
