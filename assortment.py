# Скрипт формирует единый файл с активным ассортиментом в разрезе складов МР ЮГ
# из разных файлов

import pandas as pd


CONFIG = {
    'Краснодар': 'Raw_Assortment_ALIDI_KRASNODAR.xlsx',
    'Пятигорск': 'Raw_Assortment_ALIDI_PYATIGORSK.xlsx',
    'Волгоград': 'Raw_Assortment_ALIDI_VOLGOGRAD.xlsx',
    'Краснодар-ELB': 'Raw_Assortment_ALIDI_KRASNODAR_Elbrus.xlsx',
    'Пятигорск-ELB': 'Raw_Assortment_ALIDI_PYATIGORSK_Elbrus.xlsx'
}


def get_warehouse_data(file_path, warehouse):
    """Формирует активный ассортимент по заданному складу"""
    xl = pd.ExcelFile(file_path)
    df_critical = xl.parse('Критические коды')[['EAN']]
    df_avialable = xl.parse('Ассортимент')[['EAN Заказа']]
    df_avialable = df_avialable.rename(columns={'EAN Заказа': 'EAN'})
    df = pd.concat([df_critical, df_avialable], ignore_index=True)
    df['Склад'] = f'{warehouse}'
    df.insert(0, 'Сцепка', df['Склад'] + df['EAN'].map(str))
    df = df.drop_duplicates()
    return df


def main():
    path_to_folder = 'Исходники/Ассортимент МР Юг/'
    warehouse_df = []
    for warehouse, file in CONFIG.items():
        warehouse_df.append(get_warehouse_data(
            path_to_folder + file,
            warehouse
        ))
    df_result = pd.concat(warehouse_df, ignore_index=True)
    df_result.to_excel('Результаты/Ассортимент.xlsx', index=False)


if __name__ == "__main__":
    main()
