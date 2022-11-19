"""
Скрипт подготавливает файл с остатками в удобном формате

"""

import pandas as pd


FULL_REST = 'Полное наличие  (уч.ЕИ) '
AVIALABLE_REST = 'Доступно (уч.ЕИ) '
RESERVE = 'Мягкие + жёсткие резервы (уч.ЕД)'
QUOTA = 'Остаток невыбранного резерва (уч.ЕИ)'
FREE_REST = 'Свободный остаток (уч.ЕИ)'
WHS = 'Склад'
EAN = 'EAN'
NAME = 'Наименование'
LINK = 'Сцепка'
WAREHOUSE = {
    '800WHDIS': 'Краснодар',
    '803WHDIS': 'Пятигорск',
    '815WHDIS': 'Волгоград',
    '800WHELB': 'Краснодар-ELB',
    '803WHELB': 'Пятигорск-ELB'
}


def get_filtered_df(excel, dict_warehouses, skiprows=0):
    full_df = excel.parse(skiprows=skiprows)
    filter_df = full_df[full_df[WHS].isin(list(dict_warehouses.keys()))]
    return filter_df.replace({WHS: dict_warehouses})


def main():
    xl = pd.ExcelFile(
        'Исходники/1082 - Доступность товара по складам (PG).xlsx'
    )
    filter_df = get_filtered_df(xl, WAREHOUSE)
    filter_df.insert(
        len(filter_df.axes[1]),
        RESERVE,
        filter_df[FULL_REST] - filter_df[AVIALABLE_REST]
    )

    yug_df = filter_df.groupby([WHS, EAN]).agg({
        FULL_REST: 'sum',
        RESERVE: 'sum',
        AVIALABLE_REST: 'sum',
        QUOTA: 'max'
    }).reset_index()
    yug_df.insert(0, LINK, yug_df[WHS] + yug_df[EAN].map(str))
    yug_df.insert(
        len(yug_df.axes[1]),
        FREE_REST,
        yug_df[AVIALABLE_REST] - yug_df[QUOTA]
    )
    yug_df.loc[yug_df[FREE_REST] < 0, FREE_REST] = 0
    yug_df.loc[yug_df[AVIALABLE_REST] < 0, AVIALABLE_REST] = 0
    yug_df = yug_df.merge(
        filter_df[[EAN, NAME]].drop_duplicates(subset=[EAN]),
        on='EAN',
        how='left',
    )
    yug_df.to_excel('Результаты/Остатки.xlsx', index=False)


if __name__ == "__main__":
    main()
