# Скрипт подготавливает файл с остатками в удобном формате

import pandas as pd


FULL_REST = 'Полное наличие  (уч.ЕИ) '
AVIALABLE_REST = 'Доступно (уч.ЕИ) '
RESERVE = 'Мягкие + жёсткие резервы (уч.ЕД)'
QUOTA = 'Остаток невыбранного резерва (уч.ЕИ)'
FREE_REST = 'Свободный остаток (уч.ЕИ)'

CONFIG = {
    '800WHDIS': 'Краснодар',
    '803WHDIS': 'Пятигорск',
    '815WHDIS': 'Волгоград',
    '800WHELB': 'Краснодар-ELB',
    '803WHELB': 'Пятигорск-ELB'
}

xl = pd.ExcelFile('Исходники/1082 - Доступность товара по складам (PG).xlsx')
full_df = xl.parse()
yug_df = full_df[full_df['Склад'].isin(list(CONFIG.keys()))]
yug_df = yug_df.replace({'Склад': CONFIG})
yug_df.insert(
    len(yug_df.axes[1]),
    RESERVE,
    yug_df[FULL_REST] - yug_df[AVIALABLE_REST]
)
yug_df = yug_df.groupby(['Склад', 'EAN']).agg({
  FULL_REST: 'sum',
  RESERVE: 'sum',
  AVIALABLE_REST: 'sum',
  QUOTA: 'max'
}).reset_index()
yug_df.insert(0, 'Сцепка', yug_df['Склад'] + yug_df['EAN'].map(str))
yug_df.insert(
    len(yug_df.axes[1]),
    FREE_REST,
    yug_df[AVIALABLE_REST] - yug_df[QUOTA]
)
yug_df.loc[yug_df[FREE_REST] < 0, FREE_REST] = 0
yug_df.to_excel('Результаты/Остатки.xlsx', index=False)
