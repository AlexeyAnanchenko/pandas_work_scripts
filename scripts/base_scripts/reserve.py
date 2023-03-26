"""
Скрипт подготавливает файл с резервами в удобном формате

"""
import utils
utils.path_append()

import os
import datetime as dt
import pandas as pd


from hidden_settings import WAREHOUSE_RESERVE
from service import get_filtered_df, save_to_excel, print_complete, get_data
from settings import CODES, HOLDING, NAME_HOLDING, LINK_HOLDING, LINK, WHS, EAN
from settings import SOFT_RSV, HARD_RSV, SOFT_HARD_RSV, QUOTA, PRODUCT, CURRENT
from settings import QUOTA_BY_AVAILABLE, AVAILABLE_REST, TOTAL_RSV, DATE_RSV
from settings import TABLE_RESERVE, SOURCE_DIR, RESULT_DIR, TABLE_HOLDINGS
from settings import FUTURE, TABLE_RESERVE_CURRENT, TABLE_RESERVE_FUTURE
from settings import SOFT_HARD_RSV_CURRENT, SOFT_HARD_RSV_FUTURE, EXPECTED_DATE
from settings import TABLE_RSV_BY_DATE, EXCLUDE_STRING, TABLE_EXCEPTIONS
from settings import EX_RSV, PG_PROGRAMM, WORD_YES


SOURCE_FILE = '1275 - Резервы и резервы-квоты по холдингам.xlsx'
EMPTY_ROWS = 2
HOLDING_LOC = 'Код холдинга'
RSV_HOLDING = 'Наименование'
WHS_LOC = 'Склад'
EAN_LOC = 'EAN'
PRODUCT_NAME = 'Наименование товара'
SOFT_RSV_LOC = 'Мягкий резерв'
HARD_RSV_LOC = 'Жесткий резерв'
QUOTA_RSV = 'Резерв квота остаток'
BY_LINK = '_всего, по ШК-Склад'
AVAILABLE = 'Доступность на складе'
EXPECTED_DATE_LOC = 'Дата ожидаемой доставки'
RESERVE_FOR = 'Дата резерва По'
LINK_DATE = 'Сцепка Дата-Склад-Наименование холдинга-ШК'


def get_reserve():
    excel = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    df = get_filtered_df(
        excel, WAREHOUSE_RESERVE, WHS_LOC, skiprows=EMPTY_ROWS
    )
    holdings = pd.ExcelFile(RESULT_DIR + TABLE_HOLDINGS).parse()
    df = pd.merge(
        df.rename(columns={HOLDING_LOC: CODES}),
        holdings, on=CODES, how='left'
    )
    idx = df[df[HOLDING].isnull()].index
    df.loc[idx, HOLDING] = df.loc[idx, CODES]
    df.loc[idx, NAME_HOLDING] = df.loc[idx, RSV_HOLDING]
    df.loc[df[AVAILABLE] < 0, AVAILABLE] = 0
    return df


def reserve_by_date(df):
    static_col = [
        WHS_LOC, NAME_HOLDING, EAN_LOC,
        PRODUCT_NAME, EXPECTED_DATE_LOC
    ]
    num_col = [SOFT_RSV_LOC, HARD_RSV_LOC, QUOTA_RSV]
    df = df[static_col + num_col].copy().groupby(
        static_col, dropna=False
    )[num_col].sum().reset_index()
    df[LINK_DATE] = (df[EXPECTED_DATE_LOC].map(str) + df[WHS_LOC]
                     + df[NAME_HOLDING] + df[EAN_LOC].map(str))
    df[LINK] = df[WHS_LOC] + df[EAN_LOC].map(str)

    df[EXCLUDE_STRING] = ''
    df_except = get_data(TABLE_EXCEPTIONS)
    ex_list = list(set(df_except[EX_RSV].to_list()))
    idx = df[df[LINK_DATE].isin(ex_list)].index
    df.loc[idx, EXCLUDE_STRING] = WORD_YES

    df_holdings = get_data(TABLE_HOLDINGS)[[
        NAME_HOLDING, PG_PROGRAMM
    ]].drop_duplicates(subset=[NAME_HOLDING])
    df = df.merge(df_holdings, on=NAME_HOLDING, how='left')
    static_col = [
        WHS_LOC, NAME_HOLDING, PG_PROGRAMM, EAN_LOC,
        PRODUCT_NAME, EXPECTED_DATE_LOC
    ]

    df = df[static_col + num_col + [EXCLUDE_STRING, LINK_DATE, LINK]]
    df = df.rename(columns={
        WHS_LOC: WHS,
        EAN_LOC: EAN,
        PRODUCT_NAME: PRODUCT,
        SOFT_RSV_LOC: SOFT_RSV,
        HARD_RSV_LOC: HARD_RSV,
        QUOTA_RSV: QUOTA,
        EXPECTED_DATE_LOC: EXPECTED_DATE
    })
    return df


def table_processing(df, period=None):
    idx_date = df.loc[df[EXPECTED_DATE_LOC].isnull()].index
    df.loc[idx_date, EXPECTED_DATE_LOC] = df.loc[idx_date, RESERVE_FOR]
    today = dt.date.today() - dt.timedelta(days=1)
    next_month = today.month + 1 if today.month < 12 else 1
    next_month_fday = pd.to_datetime(dt.date(today.year, next_month, 1))
    if period == CURRENT:
        df = df[df[EXPECTED_DATE_LOC] < next_month_fday]
    elif period == FUTURE:
        df = df[df[EXPECTED_DATE_LOC] >= next_month_fday]

    group_df = df.groupby([
        WHS_LOC, HOLDING, NAME_HOLDING, EAN_LOC, PRODUCT_NAME
    ]).agg({
        SOFT_RSV_LOC: 'sum',
        HARD_RSV_LOC: 'sum',
        QUOTA_RSV: 'sum',
        AVAILABLE: 'max',
        EXPECTED_DATE_LOC: 'max'
    }).reset_index()
    group_df.insert(0, LINK, group_df[WHS_LOC] + group_df[EAN_LOC].map(str))
    group_df.insert(
        0, LINK_HOLDING,
        group_df[WHS_LOC] + group_df[NAME_HOLDING] + group_df[EAN_LOC].map(str)
    )
    group_df = group_df.merge(
        group_df.groupby([LINK]).agg({QUOTA_RSV: 'sum'}),
        on=LINK,
        how='left',
        suffixes=('', BY_LINK)
    )
    group_df.insert(
        len(group_df.axes[1]),
        QUOTA_BY_AVAILABLE,
        (group_df[QUOTA_RSV] / group_df[QUOTA_RSV + BY_LINK]) * group_df[
            QUOTA_RSV + BY_LINK
        ].where(
            group_df[QUOTA_RSV + BY_LINK] < group_df[AVAILABLE],
            other=group_df[AVAILABLE]
        )
    )
    group_df.loc[group_df[QUOTA_BY_AVAILABLE].isnull(), QUOTA_BY_AVAILABLE] = 0
    group_df.drop(QUOTA_RSV + BY_LINK, axis=1, inplace=True)
    group_df[SOFT_HARD_RSV] = group_df[SOFT_RSV_LOC] + group_df[HARD_RSV_LOC]
    group_df[TOTAL_RSV] = (
        group_df[SOFT_HARD_RSV] + group_df[QUOTA_BY_AVAILABLE]
    )

    group_df = group_df.rename(columns={
        WHS_LOC: WHS,
        EAN_LOC: EAN,
        PRODUCT_NAME: PRODUCT,
        SOFT_RSV_LOC: SOFT_RSV,
        HARD_RSV_LOC: HARD_RSV,
        QUOTA_RSV: QUOTA,
        AVAILABLE: AVAILABLE_REST,
        EXPECTED_DATE_LOC: DATE_RSV
    })
    return group_df.round()


def merge_period_rsv(df):
    rsv_current = get_data(TABLE_RESERVE_CURRENT).rename(columns={
        SOFT_HARD_RSV: SOFT_HARD_RSV_CURRENT
    })[[LINK_HOLDING, SOFT_HARD_RSV_CURRENT]]
    rsv_future = get_data(TABLE_RESERVE_FUTURE).rename(columns={
        SOFT_HARD_RSV: SOFT_HARD_RSV_FUTURE
    })[[LINK_HOLDING, SOFT_HARD_RSV_FUTURE]]
    df = df.merge(rsv_current, on=LINK_HOLDING, how='left')
    df = df.merge(rsv_future, on=LINK_HOLDING, how='left')
    df = utils.void_to(df, SOFT_HARD_RSV_CURRENT, 0)
    df = utils.void_to(df, SOFT_HARD_RSV_FUTURE, 0)
    return df


def main():
    # if os.environ.get('SRS_DOWNLOAD') is None:
    #     from update_data import update_reserve
    #     update_reserve(SOURCE_FILE)
    df = get_reserve()
    df_by_date = reserve_by_date(df)
    save_to_excel(RESULT_DIR + TABLE_RSV_BY_DATE, df_by_date)
    df_current = table_processing(df, CURRENT)
    save_to_excel(RESULT_DIR + TABLE_RESERVE_CURRENT, df_current)
    df_future = table_processing(df, FUTURE)
    save_to_excel(RESULT_DIR + TABLE_RESERVE_FUTURE, df_future)
    df = merge_period_rsv(table_processing(df))
    save_to_excel(RESULT_DIR + TABLE_RESERVE, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
