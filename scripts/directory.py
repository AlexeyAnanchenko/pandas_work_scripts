"""
Формирует файл соответствия - ШК-SU-MSU

"""

import re
import pandas as pd

from category import CATEGORY_BREND, CATEGORY_SUBSECTOR, CATEGORY_SUBBREND
from service import PRODUCT, LEVEL_1, LEVEL_2, LEVEL_3, SU, MSU, EAN
from service import SOURCE_DIR, RESULT_DIR, TABLE_DIRECTORY, save_to_excel


SOURCE_FILE = 'Справочник.xlsx'
MEGA = 1000
SKIPROWS = 2
ACTIVE_GCAS = 'Активный GCAS?(на дату формирования отчета)'
PRODUCT_NAME = 'Короткое наименование (рус)'
EAN_LOC = 'EAN штуки'
SUBBREND = 'СубБренд (англ)'  # 1 уровень
BREND = 'Бренд (англ)'  # 2 уровень
SUBSECTOR = 'СубСектор (англ)'  # 3 уровень
SU_LOC = 'SU фактор штуки'
NO_DATA_VALUE = 'Нет данных'


def main():
    excel = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    df_full = excel.parse(skiprows=SKIPROWS)[[
        ACTIVE_GCAS, EAN_LOC, PRODUCT_NAME, SUBBREND, BREND, SUBSECTOR, SU_LOC
    ]].sort_index(ascending=False)
    df_full = df_full.sort_values(by=[ACTIVE_GCAS], kind='mergesort').drop(
        ACTIVE_GCAS, axis=1
    )
    df = pd.merge(
        df_full[[EAN_LOC]].drop_duplicates(),
        df_full.drop_duplicates(subset=[EAN_LOC]),
        on=EAN_LOC, how='left'
    )
    df.dropna(subset=[EAN_LOC], inplace=True)

    category_dict = {BREND: CATEGORY_BREND, SUBBREND: CATEGORY_SUBBREND}
    for col, vals in category_dict.items():
        df.loc[df[col].isin(list(vals.keys())), SUBSECTOR] = df[col]
        df = df.replace({SUBSECTOR: vals})

    df = df.replace({SUBSECTOR: CATEGORY_SUBSECTOR})

    for col in [SUBSECTOR, BREND, SUBBREND]:
        for val in list(set(df[col].tolist())):
            if bool(re.search('[а-яА-Я]', val)):
                df.loc[df[col] == val, col] = NO_DATA_VALUE

    df[MSU] = df[SU_LOC] / MEGA
    df = df.rename(columns={
        SUBBREND: LEVEL_1,
        BREND: LEVEL_2,
        SUBSECTOR: LEVEL_3,
        SU_LOC: SU,
        PRODUCT_NAME: PRODUCT,
        EAN_LOC: EAN
    })
    save_to_excel(RESULT_DIR + TABLE_DIRECTORY, df)


if __name__ == "__main__":
    main()
