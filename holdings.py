"""
Скрипт создаёт реестр соответствий кодов точек доставки,
плательщиков, холдингов к основному холдингу

"""

import pandas as pd


def main():
    excel = pd.ExcelFile('Исходники/Холдинги-Резервы.xlsx')
    df_full = excel.parse('Точка доставки-Холдинг')
    df = pd.concat([
        df_full[["Точка доставки", "Основной холдинг"]].rename(columns={"Точка доставки": "Коды"}),
        df_full[["Плательщик", "Основной холдинг"]].rename(columns={"Плательщик": "Коды"}),
        df_full[["Холдинг", "Основной холдинг"]].rename(columns={"Холдинг": "Коды"})],
        ignore_index=True
    ).drop_duplicates()
    result_data = {}
    for column in df:
        cleared_values = []
        ch1 = '('
        ch2 = ')'
        for values in df[column]:
            try:
                cleared_values.append(
                    values[(values.find(ch1) + 1):values.find(ch2)]
                )
            except Exception:
                print(values)
        result_data[column] = cleared_values
    pd.DataFrame(result_data).to_excel('Результаты/Холдинги.xlsx', index=False)


if __name__ == "__main__":
    main()
