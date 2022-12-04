"""
Скрипт подготавливает файл с закупками в удобном формате

"""

import pandas as pd
from service import get_filtered_df, save_to_excel, BASE_DIR
from service import LINK, WHS, EAN, PRODUCT, NUM_MONTHS


WHS_LOC = 'Бизнес-единица'
EAN_LOC = 'EAN'
PRODUCT_NAME = 'Наименование товара'
EMPTY_ROWS = 10
WAREHOUSES = {
    'Склад DIS г. Краснодар': 'Краснодар',
    'Склад DIS г. Пятигорск': 'Пятигорск',
    'Склад DIS г. Волгоград': 'Волгоград',
    'Склад Эльбрус г. Краснодар': 'Краснодар-ELB',
    'Склад Эльбрус г. Пятигорск': 'Пятигорск-ELB'
}


def main():
    xl = pd.ExcelFile(f'{BASE_DIR}/Исходники/Куб_Закупки.xlsx')
    df = get_filtered_df(xl, WAREHOUSES, WHS_LOC, skiprows=EMPTY_ROWS)
    df.insert(0, LINK, df[WHS_LOC] + df[EAN_LOC].map(int).map(str))
    df = df.rename(columns={WHS_LOC: WHS, EAN_LOC: EAN})
    columns = list(df.columns)
    int_col = {}

    for i in range(NUM_MONTHS):
        int_col[columns[-NUM_MONTHS + i]] = 'sum'

    group_df = df.groupby([
        LINK, WHS, EAN
    ]).agg(int_col).reset_index()
    group_df = group_df.merge(
        df[[LINK, PRODUCT_NAME]].rename(
            columns={PRODUCT_NAME: PRODUCT}
        ).drop_duplicates(subset=[LINK]),
        on=LINK, how='left'
    )
    reindex_col = [LINK, WHS, EAN, PRODUCT]
    reindex_col.extend([i for i in int_col.keys()])
    group_df = group_df.reindex(columns=reindex_col)
    save_to_excel(f'{BASE_DIR}/Результаты/Закупки.xlsx', group_df)


if __name__ == "__main__":
    main()
