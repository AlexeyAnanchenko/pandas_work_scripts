"""
Скрипт создаёт реестр соответствий кодов точек доставки,
плательщиков, холдингов к основному холдингу

"""
import utils
utils.path_append()

import pandas as pd

from service import save_to_excel, print_complete
from settings import CODES, HOLDING, NAME_HOLDING, PG_PROGRAMM
from settings import SOURCE_DIR, RESULT_DIR, TABLE_HOLDINGS


SOURCE_FILE = 'Холдинги-Резервы.xlsx'
POINT = 'Точка доставки'
PAYMENT = 'Плательщик'
HOLDING_LOC = 'Холдинг'
M_HOLDING = 'Основной холдинг'
NAME_M_HOLDING = 'Наименование основного холдинга'
PG_PROGRAMM_LOC = 'PG программа'
SUM = 'Итог'


def generate_holdings():
    excel = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    df_full = excel.parse('Точка доставки-Холдинг')
    df_full = df_full.rename(columns={
        NAME_M_HOLDING: NAME_HOLDING, M_HOLDING: HOLDING,
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
    df_result = df_result[df_result[CODES] != 'Удалить строку']
    return df_result


def merge_pg_programm(df):
    excel = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE)
    df_full = excel.parse('Точка доставки-Холдинг')
    df_full = df_full[[NAME_M_HOLDING, PG_PROGRAMM_LOC, SUM]].rename(columns={
        NAME_M_HOLDING: NAME_HOLDING, PG_PROGRAMM_LOC: PG_PROGRAMM
    })
    df_full = df_full.groupby([NAME_HOLDING, PG_PROGRAMM]).agg({
        SUM: 'sum'
    }).reset_index().sort_values(by=[SUM], ascending=[False]).drop_duplicates(
        subset=[NAME_HOLDING]
    ).drop(columns=[SUM], axis=1)
    df = df.merge(df_full, on=NAME_HOLDING, how='left')
    df = utils.void_to(df, PG_PROGRAMM, 'НЕТ ДАННЫХ')
    return df


def main():
    df = merge_pg_programm(generate_holdings())
    save_to_excel(RESULT_DIR + TABLE_HOLDINGS, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
