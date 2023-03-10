"""
Отчёт по проверке факторов согласованных и новых

"""
import utils
utils.path_append()

import pandas as pd
from datetime import date, timedelta

from hidden_settings import WHS_ELBRUS
from service import save_to_excel, get_data, print_complete
from settings import TABLE_ASSORTMENT, REPORT_ORDER_FORM, REPORT_DIR
from settings import WHS, EAN, TABLE_DIRECTORY, PRODUCT, LEVEL_3, MSU
from settings import PIC_IN_BOX, PIC_IN_LAYER, PIC_IN_PALLET, MATRIX
from settings import MATRIX_LY, MIN_ORDER, TABLE_REMAINS, FULL_REST, LINK
from settings import SOFT_RSV, QUOTA_WITH_REST, QUOTA_WITHOUT_REST
from settings import FREE_REST, TRANZIT, OVERSTOCK, HARD_RSV, EXPECTED_DATE
from settings import TABLE_RSV_BY_DATE


SOFT_RSV_FAR = 'Мягкие резервы, шт (свыше 7 дней по лог. плечу)'


def get_assortment():
    """Получаем ассортимент доступный к заказу"""
    df = get_data(TABLE_ASSORTMENT)
    col_directory = [
        EAN, PRODUCT, LEVEL_3, MSU, PIC_IN_BOX, PIC_IN_LAYER,
        PIC_IN_PALLET, MIN_ORDER, MATRIX, MATRIX_LY
    ]
    temp_product = '1'
    temp_pic_in_box = '2'
    temp_pic_in_layer = '3'
    temp_pic_in_pallet = '4'
    temp_min_order = '5'
    dict_dir = {
        PRODUCT: temp_product,
        PIC_IN_BOX: temp_pic_in_box,
        PIC_IN_LAYER: temp_pic_in_layer,
        PIC_IN_PALLET: temp_pic_in_pallet,
        MIN_ORDER: temp_min_order
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


def merge_remains(df):
    col_rem = [
        LINK, FULL_REST, HARD_RSV, SOFT_RSV, QUOTA_WITH_REST,
        QUOTA_WITHOUT_REST, FREE_REST, TRANZIT, OVERSTOCK
    ]
    df_rem = get_data(TABLE_REMAINS)[col_rem]
    df = df.merge(df_rem, on=LINK, how='left')
    for col in col_rem:
        df = utils.void_to(df, col, 0)

    df_rsv = get_data(TABLE_RSV_BY_DATE)[[LINK, SOFT_RSV, EXPECTED_DATE]]
    today = date.today()
    result = pd.to_datetime(today + timedelta(days=8))
    df_rsv = df_rsv[df_rsv[EXPECTED_DATE] >= result].groupby([LINK]).agg({
        SOFT_RSV: 'sum'
    }).reset_index().rename(columns={SOFT_RSV: SOFT_RSV_FAR})
    df = df.merge(df_rsv, on=LINK, how='left')
    return df


def merge_forecast(df):
    
    return df


def main():
    df = merge_forecast(merge_remains(get_assortment()))
    save_to_excel(REPORT_DIR + REPORT_ORDER_FORM, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
