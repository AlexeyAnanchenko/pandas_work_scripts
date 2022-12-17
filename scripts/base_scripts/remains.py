"""
Скрипт подготавливает файл с остатками в удобном формате

"""
import utils
utils.path_append()

import pandas as pd

from hidden_settings import WAREHOUSE_REMAIN
from service import get_filtered_df, save_to_excel, get_data, print_complete
from settings import LINK, WHS, EAN, PRODUCT, TARGET_STOCK, OVERSTOCK
from settings import FULL_REST, AVAILABLE_REST, SOFT_HARD_RSV, FREE_REST, QUOTA
from settings import TABLE_REMAINS, TABLE_SALES, SOURCE_DIR, RESULT_DIR
from settings import TABLE_DIRECTORY, MSU, FULL_REST_MSU, OVERSTOCK_MSU
from update_data import update_remains


SOURCE_FILE = '1082 - Доступность товара по складам (PG).xlsx'
FULL_REST_LOC = 'Полное наличие  (уч.ЕИ) '
AVAILABLE_REST_LOC = 'Доступно (уч.ЕИ) '
RESERVE = 'Резерв (уч.ЕИ) '
QUOTA_LOC = 'Остаток невыбранного резерва (уч.ЕИ)'
FREE_REST_LOC = 'Свободный остаток (уч.ЕИ)'
WHS_LOC = 'Склад'
EAN_LOC = 'EAN'
NAME = 'Наименование'


def create_remains():
    xl = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    filter_df = get_filtered_df(xl, WAREHOUSE_REMAIN, WHS_LOC)
    filter_df = filter_df.rename(columns={
        WHS_LOC: WHS,
        EAN_LOC: EAN,
        NAME: PRODUCT,
        FULL_REST_LOC: FULL_REST,
        RESERVE: SOFT_HARD_RSV,
        AVAILABLE_REST_LOC: AVAILABLE_REST,
        QUOTA_LOC: QUOTA,
        FREE_REST_LOC: FREE_REST
    })
    yug_df = filter_df.groupby([WHS, EAN]).agg({
        FULL_REST: 'sum',
        SOFT_HARD_RSV: 'sum',
        AVAILABLE_REST: 'sum',
        QUOTA: 'max'
    }).reset_index()
    yug_df.insert(0, LINK, yug_df[WHS] + yug_df[EAN].map(str))
    yug_df[FREE_REST] = yug_df[AVAILABLE_REST] - yug_df[QUOTA]
    yug_df.loc[yug_df[FREE_REST] < 0, FREE_REST] = 0
    yug_df.loc[yug_df[AVAILABLE_REST] < 0, AVAILABLE_REST] = 0
    yug_df = yug_df.merge(
        filter_df[[EAN, PRODUCT]].drop_duplicates(subset=[EAN]),
        on=EAN,
        how='left',
    )
    yug_df = yug_df.reindex(columns=[
        LINK, WHS, EAN, PRODUCT,
        FULL_REST, SOFT_HARD_RSV, AVAILABLE_REST, QUOTA, FREE_REST
    ])
    return yug_df


def added_overstock(df):
    sales, sales_col = get_data(TABLE_SALES)
    avg_cut_sale = sales_col['avg_cut_sale']
    df = df.merge(sales[[LINK, avg_cut_sale]], on=LINK, how='left')
    df[OVERSTOCK] = (df[FULL_REST] - df[avg_cut_sale] * TARGET_STOCK).round()
    idx = df[df[OVERSTOCK].isnull()].index
    df.loc[idx, OVERSTOCK] = df.loc[idx, FULL_REST]
    df.loc[df[OVERSTOCK] < 0, OVERSTOCK] = 0
    df.drop(labels=list(df[df[FULL_REST] == 0].index), axis=0, inplace=True)
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
    update_remains(SOURCE_FILE)
    result_df = conversion_msu(added_overstock(create_remains()))
    save_to_excel(RESULT_DIR + TABLE_REMAINS, result_df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
