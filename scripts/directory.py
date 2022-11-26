"""
Формирует файл соответствия - ШК-SU-MSU

"""

import re
import pandas as pd

from category import CATEGORY_BREND, CATEGORY_SUBSECTOR, CATEGORY_SUBBREND
from service import save_to_excel


MEGA = 1000
SKIPROWS = 2
ACTIVE_GCAS = 'Активный GCAS?(на дату формирования отчета)'
PRODUCT_NAME = 'Короткое наименование (рус)'
EAN = 'EAN штуки'
SUBBREND = 'СубБренд (англ)'  # 1 уровень
BREND = 'Бренд (англ)'  # 2 уровень
SUBSECTOR = 'СубСектор (англ)'  # 3 уровень
SU = 'SU фактор штуки'
MSU = 'MSU фактор штуки'
LEVEL_1 = '1-й уровень'
LEVEL_2 = '2-й уровень'
LEVEL_3 = '3-й уровень'
NO_DATA_VALUE = 'Нет данных'


def main():
    excel = pd.ExcelFile('../Исходники/Справочник.xlsx')
    df_full = excel.parse(skiprows=SKIPROWS)[[
        ACTIVE_GCAS, EAN, PRODUCT_NAME, SUBBREND, BREND, SUBSECTOR, SU
    ]].sort_index(ascending=False)
    df_full = df_full.sort_values(by=[ACTIVE_GCAS], kind='mergesort').drop(
        ACTIVE_GCAS, axis=1
    )
    df = pd.merge(
        df_full[[EAN]].drop_duplicates(),
        df_full.drop_duplicates(subset=[EAN]),
        on=EAN, how='left'
    )
    df.dropna(subset=[EAN], inplace=True)

    category_dict = {BREND: CATEGORY_BREND, SUBBREND: CATEGORY_SUBBREND}
    for col, vals in category_dict.items():
        df.loc[df[col].isin(list(vals.keys())), SUBSECTOR] = df[col]
        df = df.replace({SUBSECTOR: vals})

    df = df.replace({SUBSECTOR: CATEGORY_SUBSECTOR})

    for col in [SUBSECTOR, BREND, SUBBREND]:
        for val in list(set(df[col].tolist())):
            if bool(re.search('[а-яА-Я]', val)):
                df.loc[df[col] == val, col] = NO_DATA_VALUE

    df = df.rename(columns={
        SUBBREND: LEVEL_1, BREND: LEVEL_2, SUBSECTOR: LEVEL_3
    })
    df[MSU] = df[SU] / MEGA
    save_to_excel('../Результаты/Справочник_ШК.xlsx', df)


if __name__ == "__main__":
    main()
