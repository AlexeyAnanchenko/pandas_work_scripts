"""
Скрипт формирует единый файл с активным ассортиментом в разрезе складов МР ЮГ
из разных файлов

"""

from utils import path_import
path_import.path_append()

import pandas as pd
from os.path import basename

from service import save_to_excel
from hidden_settings import WAREHOUSE_ASSORT
from settings import SOURCE_DIR, RESULT_DIR, TABLE_ASSORTMENT, LINK, WHS, EAN

PATH_ASSORTMENT = 'Ассортимент МР Юг/'
SHEET_CRIT = 'Критические коды'
SHEET_ASSORT = 'Ассортимент'
EAN_CRIT = 'EAN'
EAN_ASSORT = 'EAN Заказа'


def get_warehouse_data(file_path, warehouse):
    """Формирует активный ассортимент по заданному складу"""
    xl = pd.ExcelFile(file_path)
    df_critical = xl.parse(SHEET_CRIT)[[EAN_CRIT]]
    df_critical = df_critical.rename(columns={EAN_CRIT: EAN})
    df_avialable = xl.parse(SHEET_ASSORT)[[EAN_ASSORT]]
    df_avialable = df_avialable.rename(columns={EAN_ASSORT: EAN})

    df = pd.concat([df_critical, df_avialable], ignore_index=True)
    df[WHS] = f'{warehouse}'
    df.insert(0, LINK, df[WHS] + df[EAN].map(str))
    df = df.drop_duplicates()
    return df


def main():
    path_to_folder = SOURCE_DIR + PATH_ASSORTMENT
    warehouse_df = []
    for warehouse, file in WAREHOUSE_ASSORT.items():
        warehouse_df.append(get_warehouse_data(
            path_to_folder + file,
            warehouse
        ))
    df_result = pd.concat(warehouse_df, ignore_index=True)
    save_to_excel(RESULT_DIR + TABLE_ASSORTMENT, df_result)
    print('Скрипт {} выполнен!'.format(basename(__file__)))


if __name__ == "__main__":
    main()
