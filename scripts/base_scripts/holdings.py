"""
Скрипт создаёт реестр соответствий кодов точек доставки,
плательщиков, холдингов к основному холдингу

"""
from sys import path
from os.path import dirname, basename
path.append(dirname(dirname(__file__)))

import pandas as pd

from service import save_to_excel
from settings import CODES, HOLDING, NAME_HOLDING
from settings import SOURCE_DIR, RESULT_DIR, TABLE_HOLDINGS


SOURCE_FILE = 'Холдинги-Резервы.xlsx'
POINT = 'Точка доставки'
PAYMENT = 'Плательщик'
HOLDING_LOC = 'Холдинг'
M_HOLDING = 'Основной холдинг'
NAME_M_HOLDING = 'Наименование основного холдинга'


def main():
    excel = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    df_full = excel.parse('Точка доставки-Холдинг')
    df_full = df_full.rename(columns={
        NAME_M_HOLDING: NAME_HOLDING, M_HOLDING: HOLDING
    })
    holding_holding = df_full[[HOLDING, NAME_HOLDING]]
    holding_holding.insert(0, M_HOLDING, df_full[HOLDING])
    df = pd.concat([
        df_full[[POINT, HOLDING, NAME_HOLDING]].rename(
            columns={POINT: CODES}),
        df_full[[PAYMENT, HOLDING, NAME_HOLDING]].rename(
            columns={PAYMENT: CODES}),
        df_full[[HOLDING_LOC, HOLDING, NAME_HOLDING]].rename(
            columns={HOLDING_LOC: CODES}),
        holding_holding.rename(
            columns={M_HOLDING: CODES})],
        ignore_index=True
    ).drop_duplicates()
    result_data = {}
    for column in df:
        cleared_values = []
        ch1 = '('
        ch2 = ')'
        for values in df[column]:
            if ch1 in str(values):
                cleared_values.append(
                    values[(values.find(ch1) + 1):values.find(ch2)]
                )
            else:
                cleared_values.append('Удалить строку')
        result_data[column] = cleared_values
    result_data[NAME_HOLDING] = df[NAME_HOLDING].to_list()
    df_result = pd.DataFrame(result_data)
    save_to_excel(
        RESULT_DIR + TABLE_HOLDINGS,
        df_result[df_result[CODES] != 'Удалить строку']
    )
    print('Скрипт {} выполнен!'.format(basename(__file__)))


if __name__ == "__main__":
    main()
