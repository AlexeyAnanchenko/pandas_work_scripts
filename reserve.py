"""
Скрипт подготавливает файл с резервами в удобном формате

"""

import pandas as pd
from remains import get_filtered_df
from holdings import NAME_M_HOLDING, M_HOLDING, CODES


EMPTY_ROWS = 2
RESERVE_HOLDING = 'Наименование'

WAREHOUSE = {
    '    800WHDIS': 'Краснодар',
    '    803WHDIS': 'Пятигорск',
    '    815WHDIS': 'Волгоград',
    '    800WHELB': 'Краснодар-ELB',
    '    803WHELB': 'Пятигорск-ELB'
}


def main():
    excel = pd.ExcelFile(
        'Исходники/1275 - Резервы и резервы-квоты по холдингам.xlsx'
    )
    df = get_filtered_df(excel, WAREHOUSE, skiprows=EMPTY_ROWS)
    holdings = pd.ExcelFile('Результаты/Холдинги.xlsx').parse()
    df = df.merge(holdings, on=CODES, how='left')
    idx = df[df[M_HOLDING].isnull()].index
    df.loc[idx, M_HOLDING] = df.loc[idx, CODES]
    df.loc[idx, NAME_M_HOLDING] = df.loc[idx, RESERVE_HOLDING]
    df.to_excel('Результаты/Резервы.xlsx', index=False)


if __name__ == "__main__":
    main()
