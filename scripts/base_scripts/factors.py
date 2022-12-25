"""
Формирует отчёт по факторам в удобном формате и с трекингом продаж

"""
import utils
utils.path_append()

import pandas as pd

from hidden_settings import WAREHOUSE_FACTORS
from service import save_to_excel, print_complete
from settings import SOURCE_DIR, RESULT_DIR, TABLE_FACTORS, WHS, FACTOR
from settings import FACTOR_NUM, REF_FACTOR
from update_data import update_factors_nfe, update_factors_pbi


SOURCE_FILE = 'NovoForecastServer_РезультатыПоиска.xlsx'
SOURCE_FILE_PB = 'Статистика факторов PG.xlsx'
WORKING_FACTORS = ['Акция', 'Предзаказ']
TOWN = 'Город'
TYPE_FACTOR = 'Тип'
ORDER_LOC = 'Сумма Заказ'
SALES_LOC = 'Сумма Продажи'
CUTS_LOC = 'Сумма Урезания'
RESERVES_LOC = 'Сумма Резервы'
HOLDING_LOC = 'Холдинг'
EAN_LOC = 'EAN'
LINK_FACTOR = 'Сцепка номер фактора-холдинг-EAN'
FACTOR_NUM_PB = 'Номер акции'


def filtered_factors(file=SOURCE_FILE, column=FACTOR, skip=0):
    """Получить отфильтрованный dataframe по нужным факторам"""
    df = pd.ExcelFile(SOURCE_DIR + file).parse(skiprows=skip)
    df = df[df[column].isin(WORKING_FACTORS)]
    df = df[df[TOWN].isin(WAREHOUSE_FACTORS)]
    df = df.rename(columns={TOWN: WHS})
    return df


def add_num_factors(df):
    """Функция добавляет номер фактора из строк"""
    cleared_rows = []
    ch_1 = '>'
    ch_2 = '<'
    for row in df[REF_FACTOR]:
        cleared_rows.append(row[(row.find(ch_1) + 1):row.find(ch_2, 1)])
    df[FACTOR_NUM] = cleared_rows
    return df


def add_pbi_col(df):
    """Добавляет столбцы из отчёта в PBI по факторам"""
    df_pb = filtered_factors(SOURCE_FILE_PB, TYPE_FACTOR, 2)
    df_pb.insert(
        0, LINK_FACTOR,
        (df_pb[FACTOR_NUM_PB].map(str) + df_pb[HOLDING_LOC]
         + df_pb[EAN_LOC].map(str))
    )
    df.insert(
        0, LINK_FACTOR,
        (df[FACTOR_NUM] + df[HOLDING_LOC] + df[EAN_LOC].map(str))
    )
    df = df.merge(
        df_pb[[LINK_FACTOR, ORDER_LOC, SALES_LOC, RESERVES_LOC, CUTS_LOC]],
        on=LINK_FACTOR, how='left'
    )
    return df


def main():
    update_factors_nfe(SOURCE_FILE)
    update_factors_pbi(SOURCE_FILE_PB)
    factors = add_pbi_col(add_num_factors(filtered_factors()))
    save_to_excel(RESULT_DIR + TABLE_FACTORS, factors)
    print_complete(__file__)


if __name__ == "__main__":
    main()
