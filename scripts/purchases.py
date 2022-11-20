"""
Скрипт подготавливает файл с закупками в удобном формате

"""

import pandas as pd
from service import get_filtered_df


WHS = 'Бизнес-единица'
WHS_RENAME = 'Склад'
EAN = 'EAN'
PRODUCT_NAME = 'Наименование товара'
LINK = 'Сцепка'
NUM_MONTH = 3
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
    group_df = df.groupby([
        LINK, WHS_RENAME, EAN
    ]).agg({
        columns[-NUM_MONTH]: 'sum',
        columns[-NUM_MONTH + 1]: 'sum',
        columns[-NUM_MONTH + 2]: 'sum'
    }).reset_index()
    group_df = group_df.merge(
        df[[LINK, PRODUCT_NAME]].drop_duplicates(subset=[LINK]),
        on=LINK, how='left'
    )
    group_df = group_df.reindex(columns=[
        LINK, WHS_RENAME, EAN, PRODUCT_NAME,
        columns[-NUM_MONTH], columns[-NUM_MONTH + 1], columns[-NUM_MONTH + 2]
    ])
    group_df.to_excel('../Результаты/Закупки.xlsx', index=False)


if __name__ == "__main__":
    main()
