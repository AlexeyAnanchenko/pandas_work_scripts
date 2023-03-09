"""
Скрипт формирует продажи в разрезе дат

"""
import utils
utils.path_append()

import pandas as pd

from hidden_settings import WAREHOUSES_SALES
from service import get_filtered_df, save_to_excel, print_complete, get_data
from settings import SOURCE_DIR, SALES_BY_DATE, CUTS_BY_DATE, DATE_SALES, LINK
from settings import RESULT_DIR, WHS, EAN, NAME_HOLDING, TABLE_SALES_BY_DATE
from settings import TABLE_EXCEPTIONS, EX_LINK_DATE, EX_NAME_HOLDING, EX_LINK
from settings import EX_EAN, EXCLUDE_STRING, TABLE_DIRECTORY, PRODUCT


SOURCE_FILE = 'Продажи по дням.xlsx'
EMPTY_ROWS = 11
WHS_LOC = 'Склад'
NAME_HOLDING_LOC = 'Основной холдинг'
EAN_LOC = 'EAN'
DATE_LOC = 'День'
CUTS_ALIDI = 'ALIDI'
CUTS_SUPPLIER = 'Supplier'
SALES_LOC = 'Sales'
LINK_DATE = 'Сцепка Дата-Склад-Клиент-Штрихкод'


def get_source():
    """Получает исходные данные"""
    excel = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    df = get_filtered_df(
        excel, WAREHOUSES_SALES, WHS_LOC, skiprows=EMPTY_ROWS
    )
    df = df.rename(columns={
        WHS_LOC: WHS,
        EAN_LOC: EAN,
        NAME_HOLDING_LOC: NAME_HOLDING,
        SALES_LOC: SALES_BY_DATE,
    })

    for col in [CUTS_SUPPLIER, CUTS_ALIDI, SALES_BY_DATE]:
        df = utils.void_to(df, col, 0)

    df[CUTS_BY_DATE] = df[CUTS_ALIDI] + df[CUTS_SUPPLIER]
    df = df[[WHS, NAME_HOLDING, EAN, DATE_LOC, SALES_BY_DATE, CUTS_BY_DATE]]
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
    df[DATE_SALES] = pd.to_datetime(df[DATE_LOC], format='%d %m %Y')
    df = df.drop(columns=[DATE_LOC], axis=1)
    df.insert(
        0, LINK,
        df[WHS] + df[EAN].map(str)
    )
    df.insert(
        0, LINK_DATE,
        df[DATE_SALES].map(str) + df[WHS] + df[NAME_HOLDING] + df[EAN].map(str)
    )
    df = df.sort_values(by=[LINK_DATE], ascending=[True])
    return df


def add_exceptions(df):
    df[EXCLUDE_STRING] = ''
    df_except = get_data(TABLE_EXCEPTIONS)
    col_dict = {
        LINK_DATE: EX_LINK_DATE,
        NAME_HOLDING: EX_NAME_HOLDING,
        LINK: EX_LINK,
        EAN: EX_EAN
    }
    for col, ex_col in col_dict.items():
        ex_list = list(set(df_except[ex_col].to_list()))
        idx = df[df[col].isin(ex_list)].index
        df.loc[idx, EXCLUDE_STRING] = 'ДА'
    return df


def merge_directory(df):
    df_dir = get_data(TABLE_DIRECTORY)[[EAN, PRODUCT]]
    df = df.merge(df_dir, on=EAN, how='left')
    df = df[[
        WHS, NAME_HOLDING, EAN, PRODUCT, SALES_BY_DATE, CUTS_BY_DATE,
        DATE_SALES, EXCLUDE_STRING, LINK_DATE, LINK
    ]]
    return df


def main():
    df = merge_directory(add_exceptions(col_to_datetime(get_source())))
    save_to_excel(RESULT_DIR + TABLE_SALES_BY_DATE, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
