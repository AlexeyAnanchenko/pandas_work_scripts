"""
Отчёт по проверке факторов согласованных и новых

"""
import utils
utils.path_append()

from hidden_settings import WHS_ELBRUS
from service import save_to_excel, get_data, print_complete
from settings import TABLE_ASSORTMENT, REPORT_ORDER_FORM, REPORT_DIR, LINK
from settings import WHS, EAN, TABLE_DIRECTORY, PRODUCT, LEVEL_3, MSU
from settings import PICIES_IN_BOX, BOXES_IN_PALLET, PICIES_IN_PALLET, MATRIX
from settings import MATRIX_LY


def get_assortment():
    """Получаем ассортимент доступный к заказу"""
    df = get_data(TABLE_ASSORTMENT)[[LINK, WHS, EAN]]
    col_directory = [
        EAN, PRODUCT, LEVEL_3, MSU, PICIES_IN_BOX, BOXES_IN_PALLET,
        PICIES_IN_PALLET, MATRIX, MATRIX_LY
    ]
    df_dir = get_data(TABLE_DIRECTORY)[col_directory]
    df = df.merge(df_dir, on=EAN, how='left')
    idx = df[
        (df[WHS].isin(WHS_ELBRUS.keys()))
        & (df[MATRIX] == 'Нет')
        & (df[MATRIX_LY] == 'Нет')
    ].index
    df = df.drop(index=idx)
    df = df.drop(columns=[MATRIX, MATRIX_LY], axis=1)

    if df.isnull().values.any():
        print('ЕСТЬ ПУСТЫЕ ШК!!!')
    return df


def main():
    df = get_assortment()
    save_to_excel(REPORT_DIR + REPORT_ORDER_FORM, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
