"""
Формирует отчёт по факторам в удобном формате и с трекингом продаж

"""
import utils
utils.path_append()

import pandas as pd

from service import save_to_excel, print_complete
from settings import SOURCE_DIR, RESULT_DIR, TABLE_FACTORS


SOURCE_FILE = 'NovoForecastServer_РезультатыПоиска.xlsx'
WORKING_FACTORS = ['Акция', 'Предзаказ']
FACTOR = 'Фактор'


def get_factors():
    df = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE).parse()
    df = df[df[FACTOR].isin(WORKING_FACTORS)]
    return df


def main():
    factors = get_factors()
    save_to_excel(RESULT_DIR + TABLE_FACTORS, factors)
    print_complete(__file__)


if __name__ == "__main__":
    main()