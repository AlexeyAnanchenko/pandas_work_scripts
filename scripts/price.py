"""
Формирует файл соответствия - ШК-Цена GIV-Цена NIV

"""

import pandas as pd

from service import save_to_excel


SKIPROWS = 7
PRODUCT_NAME = 'Название продукта '
EAN = 'EAN Код штуки'
PRICE = '  Базовая цена за шт.  '
WHS_SRS = 'Склад'
EAN_SRS = 'Штрих код'
PRODUCT_NAME_SRS = 'Описание товара'
PRICE_SRS = 'Цена'
WAREHOUSE_SRS = [
    '    800WHDIS',
    '    800WHELB',
    '    803WHDIS',
    '    803WHELB',
    '    815WHDIS'
]


def main():
    # Формируем GIV
    price = pd.ExcelFile('../Исходники/Прайс/ALIDI NORD.xlsx')
    price_df = price.parse(skiprows=SKIPROWS)[[
        PRODUCT_NAME, EAN, PRICE
    ]]
    df = pd.merge(
        price_df[[EAN, PRICE]].drop_duplicates(),
        price_df[[EAN, PRODUCT_NAME]].drop_duplicates(subset=[EAN]),
        on=EAN, how='left'
    )
    df = df.reindex(columns=[EAN, PRODUCT_NAME, PRICE])

    srs_xl = pd.ExcelFile('../Исходники/Прайс/1344 Выгрузка цен из 4106.xlsx')
    srs_df = srs_xl.parse()
    srs_df = srs_df[srs_df[WHS_SRS].isin(WAREHOUSE_SRS)][[
        EAN_SRS, PRODUCT_NAME_SRS, PRICE_SRS
    ]].drop_duplicates(subset=[EAN_SRS])

    df = pd.concat(
        [df, srs_df.rename(columns={
            EAN_SRS: EAN,
            PRODUCT_NAME_SRS: PRODUCT_NAME,
            PRICE_SRS: PRICE
        })],
        ignore_index=True
    ).drop_duplicates(subset=[EAN])
    df.dropna(subset=[EAN], inplace=True)

    save_to_excel('../Результаты/Прайс.xlsx', df)


if __name__ == "__main__":
    main()
