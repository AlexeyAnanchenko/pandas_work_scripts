"""
Отчёт по проверке факторов согласованных и новых

"""
import utils
utils.path_append()

import numpy as np
import pandas as pd
from datetime import date, timedelta, datetime

from hidden_settings import WHS_ELBRUS
from service import save_to_excel, get_data, print_complete
from settings import TABLE_ASSORTMENT, REPORT_ORDER_FORM, REPORT_DIR
from settings import WHS, EAN, TABLE_DIRECTORY, PRODUCT, LEVEL_3, MSU
from settings import PIC_IN_BOX, PIC_IN_LAYER, PIC_IN_PALLET, MATRIX
from settings import MATRIX_LY, MIN_ORDER, TABLE_REMAINS, FULL_REST, LINK
from settings import SOFT_RSV, QUOTA_WITH_REST, QUOTA_WITHOUT_REST, WORD_YES
from settings import FREE_REST, TRANZIT, OVERSTOCK, HARD_RSV, EXPECTED_DATE
from settings import TABLE_RSV_BY_DATE, EXCLUDE_STRING, QUOTA, TABLE_FACTORS
from settings import FACTOR_STATUS, FACTOR, REF_FACTOR, FACTOR_NUM, DATE_START
from settings import ACTIVE_STATUS, PURPOSE_PROMO, INACTIVE_PURPOSE, PLAN_NFE
from settings import DATE_EXPIRATION, DATE_CREATION, DESCRIPTION, USER
from settings import SALES_BY_DATE, CUTS_BY_DATE, NAME_HOLDING, TABLE_MHL
from settings import TABLE_ORDER_FACTORS, TABLE_SALES_BY_DATE, DATE_SALES
from settings import HARD_RSV_BY_DATE, SOFT_RSV_BY_DATE, QUOTA_BY_DATE


LOG_LEVERAGE = 7
SOFT_RSV_FAR = f'Мягкие резервы, шт (свыше {LOG_LEVERAGE} дней по лог. плечу)'
DEMAIND_BY_FACTOR = 'Потребность по факторам, шт'
EAN_MHL = 'EAN'
CLASSIF_N = 'классификация'
MHL = 'MHL'
LIST_A = 'Лист А'
MHL_VALUE = 'MHL'
LIST_A_VALUE = 'Лист А'
SALES_WEEK = 'Продажи '
CUTS_WEEK = 'Урезания '
WEEKS = 8
AVG_SALES = 'СДО по продажам, шт'
AVG_CUTS = 'СДО по урезаниям, шт'

today = date.today() - timedelta(days=1)
log_days = pd.to_datetime(today + timedelta(days=LOG_LEVERAGE + 1))

list_sales_week = []
list_cuts_week = []
dict_week = {
    SALES_WEEK: list_sales_week,
    CUTS_WEEK: list_cuts_week
}
for name, list_cols in dict_week.items():
    for week in range(1, WEEKS + 1):
        list_cols.append(name + f'-{week}' + ' неделя')

DICT_COLS_WEEK = {
    SALES_BY_DATE: list_sales_week,
    CUTS_BY_DATE: list_cuts_week
}


def get_assortment():
    """Получаем ассортимент доступный к заказу"""
    df = get_data(TABLE_ASSORTMENT)
    col_directory = [
        EAN, PRODUCT, LEVEL_3, MSU, PIC_IN_BOX, PIC_IN_LAYER,
        PIC_IN_PALLET, MIN_ORDER, MATRIX, MATRIX_LY
    ]
    temp_product = '1'
    temp_pic_in_box = '2'
    temp_pic_in_layer = '3'
    temp_pic_in_pallet = '4'
    temp_min_order = '5'
    dict_dir = {
        PRODUCT: temp_product,
        PIC_IN_BOX: temp_pic_in_box,
        PIC_IN_LAYER: temp_pic_in_layer,
        PIC_IN_PALLET: temp_pic_in_pallet,
        MIN_ORDER: temp_min_order
    }
    df_dir = get_data(TABLE_DIRECTORY)[col_directory]
    df_dir = df_dir.rename(columns=dict_dir)
    df = df.merge(df_dir, on=EAN, how='left')
    for col, temp in dict_dir.items():
        idx = df[df[col].isnull()].index
        df.loc[idx, col] = df.loc[idx, temp]
    df = df.drop(columns=dict_dir.values(), axis=1)

    idx = df[
        (df[WHS].isin(WHS_ELBRUS.keys()))
        & (df[MATRIX] == 'Нет')
        & (df[MATRIX_LY] == 'Нет')
    ].index
    df = df.drop(index=idx)
    df = df.drop(columns=[MATRIX, MATRIX_LY], axis=1)
    return df


def add_mhl_list_a(df):
    df_mhl = get_data(TABLE_MHL)[[EAN_MHL, CLASSIF_N]].rename(columns={
        EAN_MHL: EAN
    })
    df_mhl_copy = df_mhl[df_mhl[CLASSIF_N] == MHL_VALUE].copy()
    df = df.merge(df_mhl_copy, on=EAN, how='left').rename(columns={
        CLASSIF_N: MHL
    })
    df_mhl = df_mhl[df_mhl[CLASSIF_N] == LIST_A_VALUE]
    df = df.merge(df_mhl, on=EAN, how='left').rename(columns={
        CLASSIF_N: LIST_A
    })
    return df


def merge_remains(df):
    col_rem = [
        LINK, FULL_REST, HARD_RSV, SOFT_RSV, QUOTA_WITH_REST,
        QUOTA_WITHOUT_REST, FREE_REST, TRANZIT, OVERSTOCK
    ]
    df_rem = get_data(TABLE_REMAINS)[col_rem]
    df = df.merge(df_rem, on=LINK, how='left')
    for col in col_rem:
        df = utils.void_to(df, col, 0)

    df_rsv = get_data(TABLE_RSV_BY_DATE)[[LINK, SOFT_RSV, EXPECTED_DATE]]
    df_rsv = df_rsv[df_rsv[EXPECTED_DATE] >= log_days].groupby([LINK]).agg({
        SOFT_RSV: 'sum'
    }).reset_index().rename(columns={SOFT_RSV: SOFT_RSV_FAR})
    df = df.merge(df_rsv, on=LINK, how='left')
    df = utils.void_to(df, SOFT_RSV_FAR, 0)
    return df


def correct_by_exclude_rsv(df):
    df_rsv = get_data(TABLE_RSV_BY_DATE)
    soft_rsv_ex = 'Мягкий резерв для исключения'
    soft_rsv_ex_log = 'Мягкий резерв для исключения свыше 7 дней'
    soft_rsv_item = 'Мягкий резерв временный'
    soft_rsv_item_log = 'Мягкий резерв временный свыше 7 дней'
    quota_without_rest_item = 'Квота с товаром временная'
    quota_full_new = 'Квота без товара временная'

    df_rsv_total = df_rsv[df_rsv[EXCLUDE_STRING] == 'ДА'].groupby([LINK])[[
        SOFT_RSV, QUOTA
    ]].sum().reset_index().rename(columns={SOFT_RSV: soft_rsv_ex})

    if df_rsv_total.empty:
        return df

    df = df.merge(df_rsv_total, on=LINK, how='left')

    df_rsv = df_rsv[
        (df_rsv[EXCLUDE_STRING] == 'ДА') & (df_rsv[EXPECTED_DATE] >= log_days)
    ].groupby([LINK])[[SOFT_RSV]].sum().reset_index().rename(
        columns={SOFT_RSV: soft_rsv_ex_log}
    )
    df = df.merge(df_rsv, on=LINK, how='left')

    df[soft_rsv_item] = np.maximum(df[SOFT_RSV] - df[soft_rsv_ex], 0)
    df[soft_rsv_item_log] = np.maximum(
        df[SOFT_RSV_FAR] - df[soft_rsv_ex_log], 0
    )
    df[quota_without_rest_item] = df[QUOTA_WITHOUT_REST] - df[QUOTA]
    df[quota_full_new] = np.maximum(
        df[quota_without_rest_item] + df[QUOTA_WITH_REST], 0
    )

    idx = df[df[soft_rsv_ex] > 0].index
    df.loc[idx, SOFT_RSV] = df.loc[idx, soft_rsv_item]
    idx = df[df[soft_rsv_ex_log] > 0].index
    df.loc[idx, SOFT_RSV_FAR] = df.loc[idx, soft_rsv_item_log]
    idx = df[(df[QUOTA] > 0) & (df[quota_without_rest_item] >= 0)].index
    df.loc[idx, QUOTA_WITHOUT_REST] = df.loc[idx, quota_without_rest_item]
    idx = df[(df[QUOTA] > 0) & (df[quota_without_rest_item] < 0)].index
    df.loc[idx, QUOTA_WITHOUT_REST] = 0
    df.loc[idx, QUOTA_WITH_REST] = df.loc[idx, quota_full_new]
    idx = df[df[FULL_REST] >= 0].index
    df.loc[idx, FREE_REST] = np.maximum(
        df[FULL_REST] - df[HARD_RSV] - df[SOFT_RSV] - df[QUOTA_WITH_REST], 0
    )
    df = df.drop(columns=[
        QUOTA, soft_rsv_ex, soft_rsv_ex_log, soft_rsv_item,
        soft_rsv_item_log, quota_without_rest_item, quota_full_new
    ], axis=1)
    return df


def merge_forecast(df):
    df_fct = get_data(TABLE_FACTORS)
    df_fct = df_fct[
        (df_fct[DATE_EXPIRATION] >= log_days)
        & (df_fct[DATE_START] <= log_days)
        & (df_fct[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df_fct[PURPOSE_PROMO] != INACTIVE_PURPOSE)
    ]
    df_fct = df_fct[[
        LINK, FACTOR, REF_FACTOR, FACTOR_NUM, FACTOR_STATUS, DATE_CREATION,
        DATE_START, DATE_EXPIRATION, WHS, NAME_HOLDING, EAN, PRODUCT,
        DESCRIPTION, USER, PLAN_NFE, SALES_BY_DATE, CUTS_BY_DATE,
        HARD_RSV_BY_DATE, SOFT_RSV_BY_DATE, QUOTA_BY_DATE
    ]]
    df_fct[DEMAIND_BY_FACTOR] = np.maximum(
        (df_fct[PLAN_NFE] - df_fct[SALES_BY_DATE] - df_fct[CUTS_BY_DATE]
         - df_fct[HARD_RSV_BY_DATE] - df_fct[SOFT_RSV_BY_DATE]
         - df_fct[QUOTA_BY_DATE]),
        0)
    save_to_excel(REPORT_DIR + TABLE_ORDER_FACTORS, df_fct)
    df_fct = df_fct.groupby([LINK])[[DEMAIND_BY_FACTOR]].sum().reset_index()
    df = df.merge(df_fct, on=LINK, how='left')
    df = utils.void_to(df, DEMAIND_BY_FACTOR, 0)
    return df


def merge_sales_history(df):
    df_sales = get_data(TABLE_SALES_BY_DATE)[[
        LINK, EXCLUDE_STRING, SALES_BY_DATE, CUTS_BY_DATE, DATE_SALES
    ]]
    df_sales = df_sales[df_sales[EXCLUDE_STRING] != WORD_YES]
    for col_by_date, cols_by_week in DICT_COLS_WEEK.items():
        begin_day = datetime.today().date() - timedelta(days=1)
        for week in range(WEEKS):
            start_day = pd.to_datetime(str(begin_day))
            end_day = pd.to_datetime(start_day - timedelta(days=6))
            df_merge = df_sales[
                (df_sales[DATE_SALES] <= start_day)
                & (df_sales[DATE_SALES] >= end_day)
            ].copy().groupby([LINK])[[col_by_date]].sum().reset_index()
            df_merge = df_merge.rename(
                columns={col_by_date: cols_by_week[week]}
            )
            df = df.merge(df_merge, on=LINK, how='left')
            df = utils.void_to(df, cols_by_week[week], 0)

            begin_day = begin_day - timedelta(days=7)
    days_in_week = 7
    df[AVG_SALES] = (df[list_sales_week].sum(axis=1)
                     / WEEKS / days_in_week).round(1)
    df[AVG_CUTS] = (df[list_cuts_week].sum(axis=1)
                    / WEEKS / days_in_week).round(1)
    return df


def main():
    df = merge_remains(add_mhl_list_a(get_assortment()))
    df = merge_sales_history(merge_forecast(correct_by_exclude_rsv((df))))
    save_to_excel(REPORT_DIR + REPORT_ORDER_FORM, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
