"""
Скрипт подготавливает файл с остатками в удобном формате

"""
import utils
utils.path_append()

import os
import pandas as pd
import numpy as np
from datetime import datetime

from hidden_settings import WAREHOUSE_REMAIN
from service import get_filtered_df, save_to_excel, get_data, print_complete
from settings import LINK, WHS, EAN, PRODUCT, TARGET_STOCK, OVERSTOCK, TRANZIT
from settings import FULL_REST, AVAILABLE_REST, SOFT_HARD_RSV, FREE_REST, QUOTA
from settings import TABLE_REMAINS, TABLE_SALES, SOURCE_DIR, RESULT_DIR
from settings import TABLE_DIRECTORY, MSU, FULL_REST_MSU, OVERSTOCK_MSU
from settings import TRANZIT_CURRENT, TRANZIT_NEXT, QUOTA_WITHOUT_REST
from settings import QUOTA_WITH_REST, SOFT_RSV, HARD_RSV


SOURCE_FILE = '1082 - Доступность товара по складам (PG).xlsx'
FULL_REST_LOC = 'Полное наличие  (уч.ЕИ) '
AVAILABLE_REST_LOC = 'Доступно (уч.ЕИ) '
RESERVE = 'Резерв (уч.ЕИ) '
HARD_RSV_LOC = 'Жесткий резерв (уч. ЕИ )'
QUOTA_LOC = 'Остаток невыбранного резерва (уч.ЕИ)'
FREE_REST_LOC = 'Свободный остаток (уч.ЕИ)'
WHS_LOC = 'Склад'
EAN_LOC = 'EAN'
NAME = 'Наименование'
TRANZIT_LOC = 'Транзит штук'
DATE_TRANZIT = 'Дата планируемой доставки'
MONTH = 'Месяц транзита'
SUBSTRING = 'Транзит'


def create_remains():
    xl = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    filter_df = get_filtered_df(xl, WAREHOUSE_REMAIN, WHS_LOC)
    filter_df[SOFT_RSV] = filter_df[RESERVE] - filter_df[HARD_RSV_LOC]
    filter_df = filter_df.rename(columns={
        WHS_LOC: WHS,
        EAN_LOC: EAN,
        NAME: PRODUCT,
        FULL_REST_LOC: FULL_REST,
        RESERVE: SOFT_HARD_RSV,
        HARD_RSV_LOC: HARD_RSV,
        AVAILABLE_REST_LOC: AVAILABLE_REST,
        QUOTA_LOC: QUOTA,
        FREE_REST_LOC: FREE_REST
    })
    yug_df = filter_df.groupby([WHS, EAN]).agg({
        FULL_REST: 'sum',
        SOFT_RSV: 'sum',
        HARD_RSV: 'sum',
        SOFT_HARD_RSV: 'sum',
        AVAILABLE_REST: 'sum',
        QUOTA: 'max'
    }).reset_index()
    yug_df.insert(0, LINK, yug_df[WHS] + yug_df[EAN].map(str))
    yug_df.loc[yug_df[AVAILABLE_REST] < 0, AVAILABLE_REST] = 0
    yug_df[QUOTA_WITH_REST] = yug_df[[AVAILABLE_REST, QUOTA]].min(axis=1)
    yug_df[QUOTA_WITHOUT_REST] = yug_df[QUOTA] - yug_df[QUOTA_WITH_REST]
    yug_df[FREE_REST] = yug_df[AVAILABLE_REST] - yug_df[QUOTA]
    yug_df.loc[yug_df[FREE_REST] < 0, FREE_REST] = 0
    return yug_df


def add_transit_directory(df):
    list_dir = os.listdir(SOURCE_DIR)
    static_col = [LINK, WHS, EAN]
    transit_file = None
    for file in list_dir:
        if SUBSTRING in file:
            date_now = datetime.now().strftime("%d.%m")
            if date_now in file:
                transit_file = file
    if transit_file is not None:
        xl = pd.ExcelFile(SOURCE_DIR + transit_file)
        tz_df = get_filtered_df(xl, WAREHOUSE_REMAIN, WHS_LOC)
        tz_df = tz_df.rename(columns={
            TRANZIT_LOC: TRANZIT,
            EAN_LOC: EAN,
            WHS_LOC: WHS
        })
        tz_df[LINK] = tz_df[WHS] + tz_df[EAN].map(str)

        tz_df[MONTH] = tz_df[DATE_TRANZIT].map(
            pd.to_datetime
        ).dt.strftime('%B')
        tz_df = tz_df.pivot_table(
            values=TRANZIT, index=static_col,
            columns=[MONTH], aggfunc=np.sum
        ).reset_index()
        current_month = datetime.now().strftime("%B")
        next_month = utils.get_next_month(current_month)
        tz_df_col = tz_df.columns

        if current_month not in tz_df_col:
            tz_df[current_month] = 0
        if next_month not in tz_df_col:
            tz_df[next_month] = 0
        tz_df = tz_df.rename(columns={
            current_month: TRANZIT_CURRENT, next_month: TRANZIT_NEXT
        })
        tz_df = utils.void_to(tz_df, TRANZIT_CURRENT, 0)
        tz_df = utils.void_to(tz_df, TRANZIT_NEXT, 0)
        tz_df[TRANZIT] = tz_df[TRANZIT_CURRENT] + tz_df[TRANZIT_NEXT]
        df = df.merge(tz_df, on=static_col, how='outer')
    else:
        df[TRANZIT_CURRENT] = 0
        df[TRANZIT_NEXT] = 0
        df[TRANZIT] = 0

    direct = get_data(TABLE_DIRECTORY)[[EAN, PRODUCT]]
    df = df.merge(direct, on=EAN, how='left')
    dig_col = [
        FULL_REST, HARD_RSV, SOFT_RSV, SOFT_HARD_RSV, AVAILABLE_REST, QUOTA,
        QUOTA_WITH_REST, QUOTA_WITHOUT_REST, FREE_REST, TRANZIT_CURRENT,
        TRANZIT_NEXT, TRANZIT
    ]
    static_col.append(PRODUCT)
    df = df.reindex(columns=static_col + dig_col)
    for col in dig_col:
        df.loc[df[col].isnull(), col] = 0
    return df


def added_overstock(df):
    sales, sales_col = get_data(TABLE_SALES)
    avg_cut_sale = sales_col['avg_cut_sale']
    df = df.merge(sales[[LINK, avg_cut_sale]], on=LINK, how='left')
    df[OVERSTOCK] = (df[FULL_REST] - df[avg_cut_sale] * TARGET_STOCK).round()
    idx = df[df[OVERSTOCK].isnull()].index
    df.loc[idx, OVERSTOCK] = df.loc[idx, FULL_REST]
    df.loc[df[OVERSTOCK] < 0, OVERSTOCK] = 0
    df.loc[df[df[avg_cut_sale].isnull()].index, avg_cut_sale] = 0
    return df


def conversion_msu(df):
    direct = get_data(TABLE_DIRECTORY)
    df = df.merge(direct[[EAN, MSU]], on=EAN, how='left')
    df[FULL_REST_MSU] = df[FULL_REST] * df[MSU]
    df[OVERSTOCK_MSU] = df[OVERSTOCK] * df[MSU]
    df.drop(MSU, axis=1, inplace=True)
    return df


def main():
    if os.environ.get('SRS_DOWNLOAD') is None:
        from update_data import update_remains
        update_remains(SOURCE_FILE)
    result_df = add_transit_directory(create_remains())
    result_df = conversion_msu(added_overstock(result_df))
    save_to_excel(RESULT_DIR + TABLE_REMAINS, result_df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
