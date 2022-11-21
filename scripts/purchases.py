"""
Скрипт подготавливает файл с закупками в удобном формате

"""

import pandas as pd
from service import get_filtered_df, save_to_excel


WHS = 'Бизнес-единица'
WHS_RENAME = 'Склад'
EAN = 'EAN'
PRODUCT_NAME = 'Наименование товара'
LINK = 'Сцепка'
NUM_MONTHS = 3
EMPTY_ROWS = 10
WAREHOUSES = {
    'Склад DIS г. Краснодар': 'Краснодар',
    'Склад DIS г. Пятигорск': 'Пятигорск',
    'Склад DIS г. Волгоград': 'Волгоград',
    'Склад Эльбрус г. Краснодар': 'Краснодар-ELB',
    'Склад Эльбрус г. Пятигорск': 'Пятигорск-ELB'
}


def main():
    xl = pd.ExcelFile('../Исходники/Куб_Закупки.xlsx')
    df = get_filtered_df(xl, WAREHOUSES, WHS, skiprows=EMPTY_ROWS)
    df.insert(0, LINK, df[WHS] + df[EAN].map(int).map(str))
    df = df.rename(columns={WHS: WHS_RENAME})
    columns = list(df.columns)
    int_col = {}

    for i in range(NUM_MONTHS):
        int_col[columns[-NUM_MONTHS + i]] = 'sum'

    group_df = df.groupby([
        LINK, WHS_RENAME, EAN
    ]).agg(int_col).reset_index()
    group_df = group_df.merge(
        df[[LINK, PRODUCT_NAME]].drop_duplicates(subset=[LINK]),
        on=LINK, how='left'
    )
    reindex_col = [LINK, WHS_RENAME, EAN, PRODUCT_NAME]
    reindex_col.extend([i for i in int_col.keys()])
    group_df = group_df.reindex(columns=reindex_col)
    save_to_excel('../Результаты/Закупки.xlsx', group_df)


if __name__ == "__main__":
    main()
