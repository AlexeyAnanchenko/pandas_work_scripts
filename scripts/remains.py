"""
Скрипт подготавливает файл с остатками в удобном формате

"""

import pandas as pd

from service import get_filtered_df, save_to_excel, get_data
from service import LINK, WHS, EAN, PRODUCT, TARGET_STOCK, OVERSTOCK
from service import FULL_REST, AVAILABLE_REST, SOFT_HARD_RSV, FREE_REST, QUOTA
from service import TABLE_REMAINS, TABLE_SALES, SOURCE_DIR, RESULT_DIR


SOURCE_FILE = '1082 - Доступность товара по складам (PG).xlsx'
FULL_REST_LOC = 'Полное наличие  (уч.ЕИ) '
AVAILABLE_REST_LOC = 'Доступно (уч.ЕИ) '
RESERVE = 'Резерв (уч.ЕИ) '
QUOTA_LOC = 'Остаток невыбранного резерва (уч.ЕИ)'
FREE_REST_LOC = 'Свободный остаток (уч.ЕИ)'
WHS_LOC = 'Склад'
EAN_LOC = 'EAN'
NAME = 'Наименование'

WAREHOUSE = {
    '800WHDIS': 'Краснодар',
    '803WHDIS': 'Пятигорск',
    '815WHDIS': 'Волгоград',
    '800WHELB': 'Краснодар-ELB',
    '803WHELB': 'Пятигорск-ELB'
}


def main():
    xl = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    filter_df = get_filtered_df(xl, WAREHOUSE, WHS_LOC)
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

    sales, sales_col = get_data(TABLE_SALES)
    avg_cut_sale = sales_col['avg_cut_sale']
    yug_df = yug_df.merge(sales[[LINK, avg_cut_sale]], on=LINK, how='left')
    yug_df[OVERSTOCK] = (
        yug_df[FULL_REST] - yug_df[avg_cut_sale] * TARGET_STOCK
    ).round()
    yug_df.drop(avg_cut_sale, axis=1, inplace=True)
    idx = yug_df[yug_df[OVERSTOCK].isnull()].index
    yug_df.loc[idx, OVERSTOCK] = yug_df.loc[idx, FULL_REST]
    yug_df.loc[yug_df[OVERSTOCK] < 0, OVERSTOCK] = 0

    save_to_excel(RESULT_DIR + TABLE_REMAINS, yug_df)


if __name__ == "__main__":
    main()
