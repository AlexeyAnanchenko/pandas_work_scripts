"""
Скрипт создаёт реестр соответствий кодов точек доставки,
плательщиков, холдингов к основному холдингу

"""

import pandas as pd


POINT = 'Точка доставки'
PAYMENT = 'Плательщик'
HOLDING = 'Холдинг'
M_HOLDING = 'Основной холдинг'
CODES = 'Коды'


def main():
    excel = pd.ExcelFile('Исходники/Холдинги-Резервы.xlsx')
    df_full = excel.parse('Точка доставки-Холдинг')
    df = pd.concat([
        df_full[[POINT, M_HOLDING]].rename(columns={POINT: CODES}),
        df_full[[PAYMENT, M_HOLDING]].rename(columns={PAYMENT: CODES}),
        df_full[[HOLDING, M_HOLDING]].rename(columns={HOLDING: CODES})],
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
    df_result = pd.DataFrame(result_data)
    df_result[df_result[CODES] != 'Удалить строку'].to_excel(
        'Результаты/Холдинги.xlsx',
        index=False
    )


if __name__ == "__main__":
    main()
