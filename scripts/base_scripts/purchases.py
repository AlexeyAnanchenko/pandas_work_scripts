"""
Скрипт формирует закупки в разрезе дат

"""
import utils
utils.path_append()

import pandas as pd
from datetime import date, timedelta, datetime

from hidden_settings import WAREHOUSE_PURCH
from service import get_filtered_df, save_to_excel, print_complete, get_data
from settings import SOURCE_DIR, PURCH_BY_DATE, DATE_PURCH, LINK, PRODUCT
from settings import RESULT_DIR, WHS, EAN, TABLE_PURCH_BY_DATE, TABLE_DIRECTORY
from settings import TABLE_PURCHASES, NUM_MONTHS

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
    """Меняем формат даты"""
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
        for date_loc in uniq_date_ru:
            if month in date_loc:
                uniq_date_dict[date_loc] = date_loc.replace(month, num_month)

    df = df.replace({DATE_LOC: uniq_date_dict})
    df[DATE_PURCH] = pd.to_datetime(df[DATE_LOC], format='%d %m %Y')
    df = df.drop(columns=[DATE_LOC], axis=1)
    df = df.sort_values(by=[DATE_PURCH, LINK], ascending=[False, True])
    return df


def merge_dir(df):
    """Добавляем название товара"""
    df_dir = get_data(TABLE_DIRECTORY)[[EAN, PRODUCT]]
    df = df.merge(df_dir, on=EAN, how='left')
    df = df[[LINK, WHS, EAN, PRODUCT, DATE_PURCH, PURCH_BY_DATE]]
    return df


def get_purch_by_month(df):
    """Получаем данные по месяцам"""
    df = get_data(TABLE_PURCH_BY_DATE)
    last_sale = (date.today() - timedelta(days=1))
    date_for_loop = last_sale
    df_main = df[[LINK, WHS, EAN, PRODUCT]].drop_duplicates()
    year = date_for_loop.year
    month = date_for_loop.month
    first_day = datetime(year, month - 2, 1)
    last_day = datetime(year, month - 1, 1) - timedelta(days=1)
    month_name = first_day.strftime('%B')
    for _ in range(NUM_MONTHS):
        df_iter = df[
            (df[DATE_PURCH] >= first_day) & (df[DATE_PURCH] <= last_day)
        ].copy().groupby([LINK])[PURCH_BY_DATE].sum().reset_index()
        df_main = df_main.merge(df_iter, on=LINK, how='left')
        df_main = utils.void_to(df_main, PURCH_BY_DATE, 0)
        df_main = df_main.rename(columns={
            PURCH_BY_DATE: PURCH_BY_DATE + f' {month_name}' + f' {year}'
        })
        year = first_day.year
        month = first_day.month
        first_day = datetime(year, month + 1, 1)
        last_day = datetime(year, month + 2, 1) - timedelta(days=1)
        month_name = first_day.strftime('%B')
    return df_main


def main():
    df = merge_dir(col_to_datetime(get_source()))
    save_to_excel(RESULT_DIR + TABLE_PURCH_BY_DATE, df)
    df = get_purch_by_month(df)
    save_to_excel(RESULT_DIR + TABLE_PURCHASES, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
