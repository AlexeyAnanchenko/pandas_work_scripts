"""
Скрипт формирует единый файл с активным ассортиментом в разрезе складов МР ЮГ
из разных файлов

"""

import pandas as pd

from service import save_to_excel, LINK, WHS, EAN


SHEET_CRIT = 'Критические коды'
SHEET_ASSORT = 'Ассортимент'
EAN_CRIT = 'EAN'
EAN_ASSORT = 'EAN Заказа'
WAREHOUSE_DICT = {
    'Краснодар': 'Raw_Assortment_ALIDI_KRASNODAR.xlsx',
    'Пятигорск': 'Raw_Assortment_ALIDI_PYATIGORSK.xlsx',
    'Волгоград': 'Raw_Assortment_ALIDI_VOLGOGRAD.xlsx',
    'Краснодар-ELB': 'Raw_Assortment_ALIDI_KRASNODAR_Elbrus.xlsx',
    'Пятигорск-ELB': 'Raw_Assortment_ALIDI_PYATIGORSK_Elbrus.xlsx'
}


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
    path_to_folder = '../Исходники/Ассортимент МР Юг/'
    warehouse_df = []
    for warehouse, file in WAREHOUSE_DICT.items():
        warehouse_df.append(get_warehouse_data(
            path_to_folder + file,
            warehouse
        ))
    df_result = pd.concat(warehouse_df, ignore_index=True)
    save_to_excel('../Результаты/Ассортимент.xlsx', df_result)


if __name__ == "__main__":
    main()
