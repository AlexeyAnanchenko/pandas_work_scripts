"""
Формирует отчёт по факторам в удобном формате и с трекингом продаж

"""
import utils
utils.path_append()

import pandas as pd

from hidden_settings import WAREHOUSE_FACTORS
from service import save_to_excel, print_complete
from settings import SOURCE_DIR, RESULT_DIR, TABLE_FACTORS, WHS, FACTOR
from settings import FACTOR_NUM, LINK_FACTOR


SOURCE_FILE = 'NovoForecastServer_РезультатыПоиска.xlsx'
WORKING_FACTORS = ['Акция', 'Предзаказ']
TOWN = 'Город'


def filtered_factors():
    """Получить отфильтрованный dataframe по нужным факторам"""
    df = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE).parse()
    df = df[df[FACTOR].isin(WORKING_FACTORS)]
    df = df[df[TOWN].isin(WAREHOUSE_FACTORS)]
    df = df.rename(columns={TOWN: WHS})
    return df


def take_num_factors(df):
    """Функция вытаскивает номер фактора из строк"""
    cleared_rows = []
    ch_1 = '>'
    ch_2 = '<'
    for row in df[LINK_FACTOR]:
        cleared_rows.append(row[(row.find(ch_1) + 1):row.find(ch_2, 1)])
    df[FACTOR_NUM] = cleared_rows
    return df


def main():
    factors = take_num_factors(filtered_factors())
    save_to_excel(RESULT_DIR + TABLE_FACTORS, factors)
    print_complete(__file__)


if __name__ == "__main__":
    main()
