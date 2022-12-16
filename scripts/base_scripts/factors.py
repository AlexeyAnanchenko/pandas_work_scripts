"""
Формирует отчёт по факторам в удобном формате и с трекингом продаж

"""
import utils
utils.path_append()

import pandas as pd

from hidden_settings import WAREHOUSE_FACTORS
from service import save_to_excel, print_complete
from settings import SOURCE_DIR, RESULT_DIR, TABLE_FACTORS


SOURCE_FILE = 'NovoForecastServer_РезультатыПоиска.xlsx'
WORKING_FACTORS = ['Акция', 'Предзаказ']
FACTOR = 'Фактор'
TOWN = 'Город'


def filtered_factors():
    """Получить отфильтрованный dataframe по нужным факторам"""
    df = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE).parse()
    df = df[df[FACTOR].isin(WORKING_FACTORS)]
    df = df[df[TOWN].isin(WAREHOUSE_FACTORS)]
    return df


# def get_num_factors(df):
#     """Функция вытаскивает номер фактора из строк"""



def main():
    factors = filtered_factors()
    save_to_excel(RESULT_DIR + TABLE_FACTORS, factors)
    print_complete(__file__)


if __name__ == "__main__":
    main()
