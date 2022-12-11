"""
Скрипт подготавливает файл с закупками в удобном формате

"""
import utils
utils.path_append()

import pandas as pd

from hidden_settings import WAREHOUSE_PURCH
from service import get_filtered_df, save_to_excel, print_complete
from settings import SOURCE_DIR, RESULT_DIR, LINK, WHS, EAN
from settings import PRODUCT, NUM_MONTHS, TABLE_PURCHASES


SOURCE_FILE = 'Куб_Закупки.xlsx'
WHS_LOC = 'Бизнес-единица'
EAN_LOC = 'EAN'
PRODUCT_NAME = 'Наименование товара'
EMPTY_ROWS = 10


def main():
    xl = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    df = get_filtered_df(xl, WAREHOUSE_PURCH, WHS_LOC, skiprows=EMPTY_ROWS)
    df.insert(0, LINK, df[WHS_LOC] + df[EAN_LOC].map(int).map(str))
    df = df.rename(columns={WHS_LOC: WHS, EAN_LOC: EAN})
    columns = list(df.columns)
    int_col = {}

    for i in range(NUM_MONTHS):
        int_col[columns[-NUM_MONTHS + i]] = 'sum'

    group_df = df.groupby([
        LINK, WHS, EAN
    ]).agg(int_col).reset_index()
    group_df = group_df.merge(
        df[[LINK, PRODUCT_NAME]].rename(
            columns={PRODUCT_NAME: PRODUCT}
        ).drop_duplicates(subset=[LINK]),
        on=LINK, how='left'
    )
    reindex_col = [LINK, WHS, EAN, PRODUCT]
    reindex_col.extend([i for i in int_col.keys()])
    group_df = group_df.reindex(columns=reindex_col)
    save_to_excel(RESULT_DIR + TABLE_PURCHASES, group_df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
