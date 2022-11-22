"""
Скрипт подготавливает файл с продажами в удобном формате

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
CUTS_SALES = 'Урезания + продажи '
AVARAGE = 'Средние '
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
    # Схлопываем причины урезаний, корректно именуем столбцы и удаляем лишнее
    col_month = list(full_df.columns)[-(NUM_MONTHS * REASON_FOR_CUTS):]
    correct_month = col_month[:NUM_MONTHS]

    for i in range(NUM_MONTHS):
        full_df[CUTS + correct_month[i]] = full_df[[
            col_month[i],
            col_month[-NUM_MONTHS + i]
        ]].sum(axis=1)

    for i in range(NUM_MONTHS):
        full_df[SALES + correct_month[i]] = full_df[col_month[i + NUM_MONTHS]]

    full_df.drop(col_month, axis=1, inplace=True)
    # группируем данные по EAN
    col_month_new = {}

    for i in list(full_df.columns)[-(NUM_MONTHS * (REASON_FOR_CUTS - 1)):]:
        col_month_new[i] = 'sum'

    correct_seq = [LINK_WHS_NAME, WHS, M_HOLDING, EAN]
    group_df = full_df.groupby(correct_seq).agg(col_month_new).reset_index()
    # подтягиваем наименование товара
    group_df = group_df.merge(
        full_df[[EAN, PRODUCT_NAME]].drop_duplicates(subset=[EAN]),
        on=EAN, how='left'
    )
    correct_seq += [PRODUCT_NAME] + list(col_month_new.keys())
    group_df = group_df.reindex(columns=correct_seq)
    # добавляем новые столбцы
    cut_col = list(col_month_new.keys())[:NUM_MONTHS]
    sales_col = list(col_month_new.keys())[NUM_MONTHS:]
    cuts_sales_col = [CUTS_SALES + i for i in correct_month]

    for i in range(NUM_MONTHS):
        group_df[cuts_sales_col[i]] = group_df[[
            cut_col[i], sales_col[i]
        ]].sum(axis=1)

    def sum_all_col(reason_col, prefix):
        name_col = prefix + 'за {} месяца (-ев)'.format(NUM_MONTHS)
        group_df[name_col] = group_df[[col for col in reason_col]].sum(axis=1)
        return name_col

    cuts_all = sum_all_col(cut_col, CUTS)
    sales_all = sum_all_col(sales_col, SALES)
    cuts_sales_all = sum_all_col(cuts_sales_col, CUTS_SALES)

    for i in [cuts_all, sales_all, cuts_sales_all]:
        group_df[AVARAGE + i] = group_df[i] / NUM_MONTHS

    return group_df


def main():
    df = sales_by_client()
    save_to_excel('../Результаты/Продажи по клиентам и складам.xlsx', df)


if __name__ == "__main__":
    main()
