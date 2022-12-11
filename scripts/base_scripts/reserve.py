"""
Скрипт подготавливает файл с резервами в удобном формате

"""
from utils import path_import
path_import.path_append()

import pandas as pd
from os.path import basename

from hidden_settings import WAREHOUSE_RESERVE
from service import get_filtered_df, save_to_excel
from settings import CODES, HOLDING, NAME_HOLDING, LINK_HOLDING, LINK, WHS
from settings import SOFT_RSV, HARD_RSV, SOFT_HARD_RSV, QUOTA, PRODUCT
from settings import QUOTA_BY_AVAILABLE, AVAILABLE_REST, TOTAL_RSV, EAN
from settings import TABLE_RESERVE, SOURCE_DIR, RESULT_DIR, TABLE_HOLDINGS


SOURCE_FILE = '1275 - Резервы и резервы-квоты по холдингам.xlsx'
EMPTY_ROWS = 2
HOLDING_LOC = 'Код холдинга'
RSV_HOLDING = 'Наименование'
WHS_LOC = 'Склад'
EAN_LOC = 'EAN'
PRODUCT_NAME = 'Наименование товара'
SOFT_RSV_LOC = 'Мягкий резерв'
HARD_RSV_LOC = 'Жесткий резерв'
QUOTA_RSV = 'Резерв квота остаток'
BY_LINK = '_всего, по ШК-Склад'
AVAILABLE = 'Доступность на складе'


def main():
    excel = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    df = get_filtered_df(
        excel, WAREHOUSE_RESERVE, WHS_LOC, skiprows=EMPTY_ROWS
    )
    holdings = pd.ExcelFile(RESULT_DIR + TABLE_HOLDINGS).parse()
    df = pd.merge(
        df.rename(columns={HOLDING_LOC: CODES}),
        holdings, on=CODES, how='left'
    )
    idx = df[df[HOLDING].isnull()].index
    df.loc[idx, HOLDING] = df.loc[idx, CODES]
    df.loc[idx, NAME_HOLDING] = df.loc[idx, RSV_HOLDING]
    df.loc[df[AVAILABLE] < 0, AVAILABLE] = 0
    group_df = df.groupby([
        WHS_LOC, HOLDING, NAME_HOLDING, EAN_LOC, PRODUCT_NAME
    ]).agg({
        SOFT_RSV_LOC: 'sum',
        HARD_RSV_LOC: 'sum',
        QUOTA_RSV: 'sum',
        AVAILABLE: 'max'
    }).reset_index()
    group_df.insert(0, LINK, group_df[WHS_LOC] + group_df[EAN_LOC].map(str))
    group_df.insert(
        0, LINK_HOLDING,
        group_df[WHS_LOC] + group_df[NAME_HOLDING] + group_df[EAN_LOC].map(str)
    )
    group_df = group_df.merge(
        group_df.groupby([LINK]).agg({QUOTA_RSV: 'sum'}),
        on=LINK,
        how='left',
        suffixes=('', BY_LINK)
    )
    group_df.insert(
        len(group_df.axes[1]),
        QUOTA_BY_AVAILABLE,
        (group_df[QUOTA_RSV] / group_df[QUOTA_RSV + BY_LINK]) * group_df[
            QUOTA_RSV + BY_LINK
        ].where(
            group_df[QUOTA_RSV + BY_LINK] < group_df[AVAILABLE],
            other=group_df[AVAILABLE]
        )
    )
    group_df.loc[group_df[QUOTA_BY_AVAILABLE].isnull(), QUOTA_BY_AVAILABLE] = 0
    group_df.drop(QUOTA_RSV + BY_LINK, axis=1, inplace=True)
    group_df[SOFT_HARD_RSV] = group_df[SOFT_RSV_LOC] + group_df[HARD_RSV_LOC]
    group_df[TOTAL_RSV] = (
        group_df[SOFT_HARD_RSV] + group_df[QUOTA_BY_AVAILABLE]
    )

    group_df = group_df.rename(columns={
        WHS_LOC: WHS,
        EAN_LOC: EAN,
        PRODUCT_NAME: PRODUCT,
        SOFT_RSV_LOC: SOFT_RSV,
        HARD_RSV_LOC: HARD_RSV,
        QUOTA_RSV: QUOTA,
        AVAILABLE: AVAILABLE_REST
    })
    save_to_excel(RESULT_DIR + TABLE_RESERVE, group_df.round())
    print('Скрипт {} выполнен!'.format(basename(__file__)))


if __name__ == "__main__":
    main()
