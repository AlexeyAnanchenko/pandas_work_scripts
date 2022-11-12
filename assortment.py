# Скрипт формирует единый файл с активным ассортиментом в разрезе складов МР ЮГ
# из разных файлов

import pandas as pd


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


df_krd = get_warehouse_data(
    "./Исходники/Ассортимент МР Юг/Raw_Assortment_ALIDI_KRASNODAR.xlsx",
    'Краснодар'
)
df_pt = get_warehouse_data(
    "./Исходники/Ассортимент МР Юг/Raw_Assortment_ALIDI_PYATIGORSK.xlsx",
    'Пятигорск'
)
df_vlg = get_warehouse_data(
    "./Исходники/Ассортимент МР Юг/Raw_Assortment_ALIDI_VOLGOGRAD.xlsx",
    'Волгоград'
)
df_krd_elb = get_warehouse_data(
    "./Исходники/Ассортимент МР Юг/Raw_Assortment_ALIDI_KRASNODAR_Elbrus.xlsx",
    'Краснодар-ELB'
)
df_pt_elb = get_warehouse_data(
    './Исходники/Ассортимент МР Юг/'
    'Raw_Assortment_ALIDI_PYATIGORSK_Elbrus.xlsx',
    'Пятигорск-ELB'
)
df_result = pd.concat(
    [df_krd, df_pt, df_vlg, df_krd_elb, df_pt_elb],
    ignore_index=True
)
df_result.to_excel('Результаты/Ассортимент.xlsx', index=False)
