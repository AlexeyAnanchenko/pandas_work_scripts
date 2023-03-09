"""
Отчёт по проверке факторов согласованных и новых

"""
import utils
utils.path_append()

from hidden_settings import WHS_ELBRUS
from service import save_to_excel, get_data, print_complete
from settings import TABLE_ASSORTMENT, REPORT_ORDER_FORM, REPORT_DIR
from settings import WHS, EAN, TABLE_DIRECTORY, PRODUCT, LEVEL_3, MSU
from settings import PIC_IN_BOX, PIC_IN_LAYER, PIC_IN_PALLET, MATRIX
from settings import MATRIX_LY, MIN_ORDER

TEMP_PRODUCT = '1'
TEMP_PIC_IN_BOX = '2'
TEMP_PIC_IN_LAYER = '3'
TEMP_PIC_IN_PALLET = '4'
TEMP_MIN_ORDER = '5'


def get_assortment():
    """Получаем ассортимент доступный к заказу"""
    df = get_data(TABLE_ASSORTMENT)
    col_directory = [
        EAN, PRODUCT, LEVEL_3, MSU, PIC_IN_BOX, PIC_IN_LAYER,
        PIC_IN_PALLET, MIN_ORDER, MATRIX, MATRIX_LY
    ]
    dict_dir = {
        PRODUCT: TEMP_PRODUCT,
        PIC_IN_BOX: TEMP_PIC_IN_BOX,
        PIC_IN_LAYER: TEMP_PIC_IN_LAYER,
        PIC_IN_PALLET: TEMP_PIC_IN_PALLET,
        MIN_ORDER: TEMP_MIN_ORDER
    }
    df_dir = get_data(TABLE_DIRECTORY)[col_directory]
    df_dir = df_dir.rename(columns=dict_dir)
    df = df.merge(df_dir, on=EAN, how='left')
    for col, temp in dict_dir.items():
        idx = df[df[col].isnull()].index
        df.loc[idx, col] = df.loc[idx, temp]
    df = df.drop(columns=dict_dir.values(), axis=1)

    idx = df[
        (df[WHS].isin(WHS_ELBRUS.keys()))
        & (df[MATRIX] == 'Нет')
        & (df[MATRIX_LY] == 'Нет')
    ].index
    df = df.drop(index=idx)
    df = df.drop(columns=[MATRIX, MATRIX_LY], axis=1)
    return df


def main():
    df = get_assortment()
    save_to_excel(REPORT_DIR + REPORT_ORDER_FORM, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
