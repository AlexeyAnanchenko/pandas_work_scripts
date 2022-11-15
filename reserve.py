"""
Скрипт подготавливает файл с резервами в удобном формате

"""

import pandas as pd
from remains import get_filtered_df


EMPTY_ROWS = 2

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
    print(df[['Код холдинга']])


if __name__ == "__main__":
    main()
