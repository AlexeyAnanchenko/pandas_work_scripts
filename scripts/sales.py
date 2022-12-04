"""
Скрипт подготавливает файл с продажами в удобном формате

"""

import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta

from service import get_filtered_df, save_to_excel, BASE_DIR
from service import WHS, EAN, NAME_HOLDING, PRODUCT, NUM_MONTHS
from service import LINK, LINK_HOLDING, CUTS, SALES, CUTS_SALES, AVARAGE


EMPTY_ROWS = 13
REASON_FOR_CUTS = 3
WHS_LOC = 'Склад'
EAN_LOC = 'EAN'
M_HOLDING = 'Основной холдинг'
PRODUCT_NAME = 'Наименование товара'
WAREHOUSES = {
    'Склад DIS г. Краснодар (800WHDIS)': 'Краснодар',
    'Склад DIS г. Пятигорск (803WHDIS)': 'Пятигорск',
    'Склад DIS г. Волгоград (815WHDIS)': 'Волгоград',
    'Склад Эльбрус г. Краснодар (800WHELB)': 'Краснодар-ELB',
    'Склад Эльбрус г. Пятигорск (803WHELB)': 'Пятигорск-ELB'
}


def sales_by_client():
    """Формирование фрейма данных продаж в разрезе клиент-склад-шк"""

    excel = pd.ExcelFile(f'{BASE_DIR}/Исходники/Продажи общие.xlsx')
    full_df = get_filtered_df(excel, WAREHOUSES, WHS, skiprows=EMPTY_ROWS)
    full_df = full_df.rename(columns={
        EAN_LOC: EAN,
        WHS_LOC: WHS,
        M_HOLDING: NAME_HOLDING,
        PRODUCT_NAME: PRODUCT
    })
    full_df.insert(
        0, LINK_HOLDING,
        full_df[WHS] + full_df[NAME_HOLDING] + full_df[EAN].map(int).map(str)
    )

    # Суммируем урезания, корректно именуем столбцы и удаляем лишнее
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

    correct_seq = [LINK_HOLDING, WHS, NAME_HOLDING, EAN]
    group_df = full_df.groupby(correct_seq).agg(col_month_new).reset_index()

    # подтягиваем наименование товара
    group_df = group_df.merge(
        full_df[[EAN, PRODUCT]].drop_duplicates(subset=[EAN]),
        on=EAN, how='left'
    )
    correct_seq += [PRODUCT] + list(col_month_new.keys())
    group_df = group_df.reindex(columns=correct_seq)

    # добавляем столбцы продажи + урезания
    cut_col = list(col_month_new.keys())[:NUM_MONTHS]
    sales_col = list(col_month_new.keys())[NUM_MONTHS:]
    cuts_sales_col = [CUTS_SALES + i for i in correct_month]

    for i in range(NUM_MONTHS):
        group_df[cuts_sales_col[i]] = group_df[[
            cut_col[i], sales_col[i]
        ]].sum(axis=1)

    # формируем столбцы с суммой за все месяца
    def get_sum_all_columns(reason_col, prefix):
        name_col = prefix + 'за {} месяца (-ев)'.format(NUM_MONTHS)
        group_df[name_col] = group_df[[col for col in reason_col]].sum(axis=1)
        return name_col

    sum_all_columns = [
        get_sum_all_columns(cut_col, CUTS),
        get_sum_all_columns(sales_col, SALES),
        get_sum_all_columns(cuts_sales_col, CUTS_SALES)
    ]

    # определяем делитель (кол-во дней) для расчёта среднего
    first_day_current_month = date(date.today().year, date.today().month, 1)
    delta = relativedelta(months=(3 - 1))
    divide_for_avarage = (
        date.today() - (first_day_current_month - delta)
    ).days
    avarage_days_in_month = 365 / 12

    # создаём финальные столбцы и удаляем суммы за все месяца
    avarage_col = []

    for name_col in sum_all_columns:
        new_name = AVARAGE + name_col
        group_df[new_name] = (group_df[name_col]
                              / divide_for_avarage
                              * avarage_days_in_month)
        avarage_col.append(new_name)

    group_df.drop(sum_all_columns, axis=1, inplace=True)
    numeric_col = [
        cut_col + sales_col + cuts_sales_col + avarage_col
    ]
    return group_df, numeric_col


def sales_by_warehouses(dataframe, numeric_columns):
    """Формирование фрейма данных продаж в разрезе склад-шк"""

    dataframe.insert(
        0, LINK,
        dataframe[WHS] + dataframe[EAN].map(int).map(str)
    )
    sequence_col = [LINK, WHS, EAN, PRODUCT]
    numeric_col_dict = {}

    for col_list in numeric_columns:
        for col in col_list:
            numeric_col_dict[col] = 'sum'

    df = dataframe.groupby(sequence_col).agg(numeric_col_dict).reset_index()
    return df


def main():
    df, numeric_columns = sales_by_client()
    save_to_excel(
        f'{BASE_DIR}/Результаты/Продажи по клиентам и складам.xlsx', df
    )
    df = sales_by_warehouses(df, numeric_columns)
    save_to_excel(f'{BASE_DIR}/Результаты/Продажи по складам.xlsx', df)


if __name__ == "__main__":
    main()
