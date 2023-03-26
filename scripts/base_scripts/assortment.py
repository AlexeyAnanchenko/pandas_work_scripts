"""
Скрипт формирует единый файл с активным ассортиментом в разрезе складов МР ЮГ
из разных файлов

"""

import utils
utils.path_append()

import pandas as pd

from service import save_to_excel, print_complete
from hidden_settings import WAREHOUSE_ASSORT
from settings import SOURCE_DIR, RESULT_DIR, TABLE_ASSORTMENT, LINK, WHS, EAN
from settings import PIC_IN_BOX, PIC_IN_LAYER, PIC_IN_PALLET, PRODUCT
from settings import MIN_ORDER, CRITICAL_EAN, WORD_YES


PATH_ASSORTMENT = 'Ассортимент МР Юг/'
SHEET_CRIT = 'Критические коды'
SHEET_ASSORT = 'Ассортимент'
EAN_CRIT = 'EAN'
EAN_ASSORT = 'EAN Заказа'
PRODUCT_NAME = 'Название товара'
MIN_ORDER_LOC = 'Минимальная единица заказа'
PIC_IN_BOX_LOC = 'Штук в коробке'
BOXES_IN_LAYER = 'Кол-во коробок в ряду'
BOXES_IN_PALLET = 'Кол-во коробок в паллете'
MIN_ORDER_DICT = {
    'CASE': PIC_IN_BOX,
    'LAYER': PIC_IN_LAYER,
    'PALLET': PIC_IN_PALLET
}


def get_warehouse_data(file_path, warehouse):
    """Формирует активный ассортимент по заданному складу"""
    xl = pd.ExcelFile(file_path)
    df_critical_full = xl.parse(SHEET_CRIT)[[EAN_CRIT]]
    df_critical_full = df_critical_full.rename(columns={EAN_CRIT: EAN})
    df_critical_full[CRITICAL_EAN] = WORD_YES
    df_avialable_full = xl.parse(SHEET_ASSORT)[[
        EAN_ASSORT, PRODUCT_NAME, MIN_ORDER_LOC, PIC_IN_BOX_LOC,
        BOXES_IN_LAYER, BOXES_IN_PALLET
    ]]
    df_avialable_full = df_avialable_full.rename(columns={
        EAN_ASSORT: EAN, PRODUCT_NAME: PRODUCT, MIN_ORDER_LOC: MIN_ORDER,
        PIC_IN_BOX_LOC: PIC_IN_BOX
    })
    df_critical = df_critical_full[[EAN]]
    df_avialable = df_avialable_full[[EAN]]
    df_critical_full = df_critical_full.drop_duplicates(subset=EAN)
    df_avialable_full = df_avialable_full.drop_duplicates(subset=EAN)

    df = pd.concat([df_critical, df_avialable], ignore_index=True)
    df[WHS] = f'{warehouse}'
    df.insert(0, LINK, df[WHS] + df[EAN].map(str))
    df = df.drop_duplicates()
    df = df.merge(df_critical_full, on=EAN, how='left')
    df = df.merge(df_avialable_full, on=EAN, how='left')
    df[PIC_IN_LAYER] = df[PIC_IN_BOX] * df[BOXES_IN_LAYER]
    df[PIC_IN_PALLET] = df[PIC_IN_BOX] * df[BOXES_IN_PALLET]

    for val, col in MIN_ORDER_DICT.items():
        idx = df[df[MIN_ORDER] == val].index
        df.loc[idx, MIN_ORDER] = df.loc[idx, col]

    df = df[[
        LINK, WHS, EAN, CRITICAL_EAN, PRODUCT, PIC_IN_BOX,
        PIC_IN_LAYER, PIC_IN_PALLET, MIN_ORDER
    ]]
    return df


def main():
    path_to_folder = SOURCE_DIR + PATH_ASSORTMENT
    warehouse_df = []
    for warehouse, file in WAREHOUSE_ASSORT.items():
        warehouse_df.append(get_warehouse_data(
            path_to_folder + file,
            warehouse
        ))
    df_result = pd.concat(warehouse_df, ignore_index=True)
    save_to_excel(RESULT_DIR + TABLE_ASSORTMENT, df_result)
    print_complete(__file__)


if __name__ == "__main__":
    main()
