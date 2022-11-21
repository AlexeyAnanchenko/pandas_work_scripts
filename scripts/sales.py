"""
Скрипт подготавливает файл с резервами в удобном формате

"""

import pandas as pd

from service import get_filtered_df, save_to_excel


EMPTY_ROWS = 13
NUM_MONTHS = 3
REASON_FOR_CUTS = 3
WHS = 'Склад'
EAN = 'EAN'
M_HOLDING = 'Основной холдинг'
SALES_DIRECT = 'Направление продаж'
PRODUCT_NAME = 'Наименование товара'
LINK_WHS_NAME = 'Сцепка Склад-Наименование холдинга-ШК'
LINK_WHS = 'Сцепка Склад-ШК'
CUTS = 'Урезания '
SALES = 'Продажи '

WAREHOUSES = {
    'Склад DIS г. Краснодар (800WHDIS)': 'Краснодар',
    'Склад DIS г. Пятигорск (803WHDIS)': 'Пятигорск',
    'Склад DIS г. Волгоград (815WHDIS)': 'Волгоград',
    'Склад Эльбрус г. Краснодар (800WHELB)': 'Краснодар-ELB',
    'Склад Эльбрус г. Пятигорск (803WHELB)': 'Пятигорск-ELB'
}


def sales_by_client():
    excel = pd.ExcelFile('../Исходники/Продажи общие.xlsx')
    full_df = get_filtered_df(excel, WAREHOUSES, WHS, skiprows=EMPTY_ROWS)
    full_df.insert(
        0, LINK_WHS_NAME,
        full_df[WHS] + full_df[M_HOLDING] + full_df[EAN].map(int).map(str)
    )
    col_month = list(full_df.columns)[-(NUM_MONTHS * REASON_FOR_CUTS):]
    correct_month = col_month[:NUM_MONTHS]

    for i in range(len(correct_month)):
        full_df[CUTS + correct_month[i]] = full_df[[
            col_month[i],
            col_month[-NUM_MONTHS + i]
        ]].sum(axis=1)

    for i in range(len(correct_month)):
        full_df[SALES + correct_month[i]] = full_df[col_month[i + NUM_MONTHS]]

    full_df.drop(col_month, axis=1, inplace=True)
    
    cuts_all = CUTS + 'за {} месяца (-ев)'.format(NUM_MONTHS)
    sales_all = SALES + 'за {} месяца (-ев)'.format(NUM_MONTHS)
    return full_df


def main():
    df = sales_by_client()
    save_to_excel('../Результаты/Продажи по клиентам и складам.xlsx', df[:10])


if __name__ == "__main__":
    main()
