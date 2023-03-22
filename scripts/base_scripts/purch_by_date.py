"""
Скрипт формирует закупки в разрезе дат

"""
import utils
utils.path_append()

import pandas as pd

from hidden_settings import WAREHOUSE_PURCH
from service import get_filtered_df, save_to_excel, print_complete, get_data
from settings import SOURCE_DIR, PURCH_BY_DATE, DATE_PURCH, LINK, PRODUCT
from settings import RESULT_DIR, WHS, EAN, TABLE_PURCH_BY_DATE, TABLE_DIRECTORY

SOURCE_FILE = 'Закупки по дням.xlsx'
EMPTY_ROWS = 8
WHS_LOC = 'Бизнес-единица'
EAN_LOC = 'EAN'
DATE_LOC = 'День'
PURCH_LOC = 'Закупки шт'


def get_source():
    """Получает исходные данные"""
    excel = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    df = get_filtered_df(
        excel, WAREHOUSE_PURCH, WHS_LOC, skiprows=EMPTY_ROWS
    )
    df = df.rename(columns={
        WHS_LOC: WHS,
        EAN_LOC: EAN,
        PURCH_LOC: PURCH_BY_DATE,
    })
    df.insert(0, LINK, df[WHS] + df[EAN].map(int).map(str))
    df = df[[LINK, WHS, EAN, DATE_LOC, PURCH_BY_DATE]]
    return df


def col_to_datetime(df):
    ru_en_month_dict = {
        'Январь': '01',
        'Февраль': '02',
        'Март': '03',
        'Апрель': '04',
        'Май': '05',
        'Июнь': '06',
        'Июль': '07',
        'Август': '08',
        'Сентябрь': '09',
        'Октябрь': '10',
        'Ноябрь': '11',
        'Декабрь': '12'
    }

    uniq_date_ru = list(set(df[DATE_LOC].tolist()))
    uniq_date_dict = {}
    for month, num_month in ru_en_month_dict.items():
        for date in uniq_date_ru:
            if month in date:
                uniq_date_dict[date] = date.replace(month, num_month)

    df = df.replace({DATE_LOC: uniq_date_dict})
    df[DATE_PURCH] = pd.to_datetime(df[DATE_LOC], format='%d %m %Y')
    df = df.drop(columns=[DATE_LOC], axis=1)
    df = df.sort_values(by=[DATE_PURCH, LINK], ascending=[False, True])
    return df


def merge_dir(df):
    df_dir = get_data(TABLE_DIRECTORY)[[EAN, PRODUCT]]
    df = df.merge(df_dir, on=EAN, how='left')
    df = df[[LINK, WHS, EAN, PRODUCT, DATE_PURCH, PURCH_BY_DATE]]
    return df


def main():
    df = merge_dir(col_to_datetime(get_source()))
    save_to_excel(RESULT_DIR + TABLE_PURCH_BY_DATE, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
