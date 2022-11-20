"""
Скрипт подготавливает файл с резервами в удобном формате

"""

import pandas as pd

from service import get_filtered_df, save_to_excel
from holdings import NAME_M_HOLDING, M_HOLDING, CODES


EMPTY_ROWS = 2
RSV_HOLDING = 'Наименование'
WHS = 'Склад'
EAN = 'EAN'
PRODUCT_NAME = 'Наименование товара'
SOFT_RSV = 'Мягкий резерв'
HARD_RSV = 'Жесткий резерв'
QUOTA_RSV = 'Резерв квота остаток'
BY_LINK = '_всего, по ШК-Склад'
QUOTA_BY_AVAILABLE = 'Резерв квота с учётом доступного стока'
AVAILABLE = 'Доступность на складе'
LINK_WHS = 'Сцепка Склад-ШК'
LINK_WHS_NAME = 'Сцепка Склад-Наименование холдинга-ШК'
SOFT_HARD_RSV = 'Жёсткий + мягкий резерв'
TOTAL_RSV = 'В резерве всего'
WAREHOUSE = {
    '    800WHDIS': 'Краснодар',
    '    803WHDIS': 'Пятигорск',
    '    815WHDIS': 'Волгоград',
    '    800WHELB': 'Краснодар-ELB',
    '    803WHELB': 'Пятигорск-ELB'
}


def main():
    excel = pd.ExcelFile(
        '../Исходники/1275 - Резервы и резервы-квоты по холдингам.xlsx'
    )
    df = get_filtered_df(excel, WAREHOUSE, WHS, skiprows=EMPTY_ROWS)
    holdings = pd.ExcelFile('../Результаты/Холдинги.xlsx').parse()
    df = df.merge(holdings, on=CODES, how='left')
    idx = df[df[M_HOLDING].isnull()].index
    df.loc[idx, M_HOLDING] = df.loc[idx, CODES]
    df.loc[idx, NAME_M_HOLDING] = df.loc[idx, RSV_HOLDING]
    df.loc[df[AVAILABLE] < 0, AVAILABLE] = 0
    group_df = df.groupby([
        WHS, M_HOLDING, NAME_M_HOLDING, EAN, PRODUCT_NAME
    ]).agg({
        SOFT_RSV: 'sum',
        HARD_RSV: 'sum',
        QUOTA_RSV: 'sum',
        AVAILABLE: 'max'
    }).reset_index()
    group_df.insert(0, LINK_WHS, group_df[WHS] + group_df[EAN].map(str))
    group_df.insert(
        0, LINK_WHS_NAME,
        group_df[WHS] + group_df[NAME_M_HOLDING] + group_df[EAN].map(str)
    )
    group_df = group_df.merge(
        group_df.groupby([LINK_WHS]).agg({QUOTA_RSV: 'sum'}),
        on=LINK_WHS,
        how='left',
        suffixes=('', BY_LINK)
    )
    group_df.insert(
        len(group_df.axes[1]),
        QUOTA_BY_AVAILABLE,
        (group_df[QUOTA_RSV] / group_df[QUOTA_RSV + BY_LINK])
        * group_df[QUOTA_RSV + BY_LINK].where(
            group_df[QUOTA_RSV + BY_LINK] < group_df[AVAILABLE],
            other=group_df[AVAILABLE]
        )
    )
    group_df.loc[group_df[QUOTA_BY_AVAILABLE].isnull(), QUOTA_BY_AVAILABLE] = 0
    group_df.drop(QUOTA_RSV + BY_LINK, axis=1, inplace=True)
    group_df.insert(
        len(group_df.axes[1]),
        SOFT_HARD_RSV,
        group_df[SOFT_RSV] + group_df[HARD_RSV]
    )
    group_df.insert(
        len(group_df.axes[1]),
        TOTAL_RSV,
        group_df[SOFT_HARD_RSV] + group_df[QUOTA_BY_AVAILABLE]
    )
    save_to_excel('../Результаты/Резервы.xlsx', group_df.round())


if __name__ == "__main__":
    main()
