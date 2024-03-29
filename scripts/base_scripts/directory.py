"""
Формирует файл соответствия - ШК-SU-MSU

"""
import utils
utils.path_append()

import re
import pandas as pd

from service import save_to_excel, get_data, print_complete
from hidden_settings import CATEGORY_BREND
from hidden_settings import CATEGORY_SUBSECTOR, CATEGORY_SUBBREND
from settings import PRODUCT, LEVEL_1, LEVEL_2, LEVEL_3, LEVEL_0
from settings import SOURCE_DIR, RESULT_DIR, TABLE_DIRECTORY, TABLE_PRICE
from settings import MATRIX, MATRIX_LY, BASE_PRICE, ELB_PRICE, SU, MSU, EAN
from settings import PIC_IN_BOX, PIC_IN_LAYER, PIC_IN_PALLET, MIN_ORDER


SOURCE_FILE = 'Справочник/Справочник.xlsx'
SOURCE_MATRIX = 'Справочник/Матрица Эльбрус.xlsx'
SOURCE_ELB = 'Справочник/Матрица ЛЮ.xlsx'
MEGA = 1000
SKIPROWS = 2
ACTIVE_GCAS = 'Активный GCAS?(на дату формирования отчета)'
PRODUCT_NAME = 'Короткое наименование (рус)'
EAN_LOC = 'EAN штуки'
DETAIL_SUBBREND = 'Детали суббренда (англ)'  # 0 уровень
SUBBREND = 'СубБренд (англ)'  # 1 уровень
BREND = 'Бренд (англ)'  # 2 уровень
SUBSECTOR = 'СубСектор (англ)'  # 3 уровень
SU_LOC = 'SU фактор штуки'
NO_DATA_VALUE = 'Нет данных'
BOXES_IN_PALLET_LOC = 'Количество коробок в паллете'
PICIES_IN_BOX_LOC = 'Количество штук в коробке'
BOXES_IN_LAYER_LOC = 'Количество коробок в ряду'
MIN_ORDER_LOC = ('Минимальная единица заказа для премии за '
                 'логистическую эффективность')
MIN_ORDER_DICT = {
    'CS': PIC_IN_BOX,
    'L': PIC_IN_LAYER,
    'P': PIC_IN_PALLET
}
# список категорий из колонки 'Бренд (англ)', которые нужно
# удалить из справочника
DELETE_LEVEL_2 = ['WELLAFLEX']


def get_category_msu():
    """Формирует справочник с msu и категориями"""
    excel = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    df_full = excel.parse(skiprows=SKIPROWS)[[
        ACTIVE_GCAS, EAN_LOC, PRODUCT_NAME, DETAIL_SUBBREND, SUBBREND,
        BREND, SUBSECTOR, SU_LOC, MIN_ORDER_LOC, PICIES_IN_BOX_LOC,
        BOXES_IN_LAYER_LOC, BOXES_IN_PALLET_LOC
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

    cols = [
        PICIES_IN_BOX_LOC, BOXES_IN_PALLET_LOC,
        BOXES_IN_LAYER_LOC, MIN_ORDER_LOC
    ]
    for col in cols:
        df = utils.void_to(df, col, 0)

    df[PIC_IN_LAYER] = df[PICIES_IN_BOX_LOC] * df[BOXES_IN_LAYER_LOC]
    df[PIC_IN_PALLET] = df[PICIES_IN_BOX_LOC] * df[BOXES_IN_PALLET_LOC]
    df[MSU] = df[SU_LOC] / MEGA
    df = df.rename(columns={
        DETAIL_SUBBREND: LEVEL_0,
        SUBBREND: LEVEL_1,
        BREND: LEVEL_2,
        SUBSECTOR: LEVEL_3,
        SU_LOC: SU,
        PRODUCT_NAME: PRODUCT,
        EAN_LOC: EAN,
        PICIES_IN_BOX_LOC: PIC_IN_BOX,
        MIN_ORDER_LOC: MIN_ORDER
    })
    df = df.drop(columns=[BOXES_IN_LAYER_LOC, BOXES_IN_PALLET_LOC], axis=1)

    for val, col in MIN_ORDER_DICT.items():
        idx = df[df[MIN_ORDER] == val].index
        df.loc[idx, MIN_ORDER] = df.loc[idx, col]

    return df


def added_matrix(dataframe):
    """Добавляет матрицы Эльбруса в справочник"""
    matrix_df = pd.ExcelFile(SOURCE_DIR + SOURCE_MATRIX).parse()
    matrix__df_LY = pd.ExcelFile(SOURCE_DIR + SOURCE_ELB).parse()
    matrix_df[MATRIX] = 'Да'
    matrix__df_LY[MATRIX_LY] = 'Да'
    dataframe = dataframe.merge(
        matrix_df.rename(columns={matrix_df.columns[0]: EAN}),
        on=EAN, how='left'
    )
    dataframe = dataframe.merge(
        matrix__df_LY.rename(columns={matrix__df_LY.columns[0]: EAN}),
        on=EAN, how='left'
    )
    return dataframe


def added_price(dataframe):
    """Добавляет цены в справочник"""
    price = get_data(TABLE_PRICE)[[EAN, BASE_PRICE, ELB_PRICE]]
    dataframe = dataframe.merge(price, on=EAN, how='left')
    return dataframe


def delete_category_level(df):
    """Удаляет категории и уровни, которые пока не нужны в справочнике"""
    for i in DELETE_LEVEL_2:
        df = df[df[LEVEL_2] != i]

    df = df.drop(labels=[LEVEL_0, LEVEL_1, LEVEL_2], axis=1)
    return df


def fill_empty_cells(df):
    """Заполняет пустые ячейки нулями"""
    numeric_col = [SU, MSU, BASE_PRICE, ELB_PRICE]

    for col in numeric_col:
        df = utils.void_to(df, col, 0)

    matrix_col = [MATRIX, MATRIX_LY]
    for col in matrix_col:
        df = utils.void_to(df, col, 'Нет')

    return df


def main():
    directory = added_price(added_matrix(get_category_msu()))
    directory = fill_empty_cells(delete_category_level(directory))
    save_to_excel(RESULT_DIR + TABLE_DIRECTORY, directory)
    print_complete(__file__)


if __name__ == "__main__":
    main()
