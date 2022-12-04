"""
Формирует файл соответствия - ШК-Цена GIV-Цена NIV

"""

import pandas as pd

from service import save_to_excel, BASE_DIR
from service import EAN, PRODUCT, BASE_PRICE, ELB_PRICE


SKIPROWS = 7
PRODUCT_NAME = 'Название продукта '
EAN_LOC = 'EAN Код штуки'
PRICE = '  Базовая цена за шт.  '
WHS_SRS = 'Склад'
EAN_SRS = 'Штрих код'
PRODUCT_NAME_SRS = 'Описание товара'
PRICE_SRS = 'Цена'
EAN_ELB = 'Штрих-код штуки'
PRODUCT_NAME_ELB = 'Наименование продукта'
PRICE_ELB = 'Цена Прайса'
PRICE_ELB_ADD = ' Эльбрус, NIV с НДС'
SKIPROWS_ELB = 10
WAREHOUSE_SRS = [
    '    800WHDIS',
    '    800WHELB',
    '    803WHDIS',
    '    803WHELB',
    '    815WHDIS'
]


def main():
    # Формируем GIV
    price = pd.ExcelFile(f'{BASE_DIR}/Исходники/Прайс/ALIDI NORD.xlsx')
    price_df = price.parse(skiprows=SKIPROWS)[[
        PRODUCT_NAME, EAN_LOC, PRICE
    ]]
    df = pd.merge(
        price_df[[EAN_LOC, PRICE]].drop_duplicates(),
        price_df[[EAN_LOC, PRODUCT_NAME]].drop_duplicates(subset=[EAN_LOC]),
        on=EAN_LOC, how='left'
    )
    df = df.reindex(columns=[EAN_LOC, PRODUCT_NAME, PRICE])

    srs_xl = pd.ExcelFile(
        f'{BASE_DIR}/Исходники/Прайс/1344 Выгрузка цен из 4106.xlsx'
    )
    srs_df = srs_xl.parse()
    srs_df = srs_df[srs_df[WHS_SRS].isin(WAREHOUSE_SRS)][[
        EAN_SRS, PRODUCT_NAME_SRS, PRICE_SRS
    ]].drop_duplicates(subset=[EAN_SRS])

    df = pd.concat(
        [df, srs_df.rename(columns={
            EAN_SRS: EAN_LOC,
            PRODUCT_NAME_SRS: PRODUCT_NAME,
            PRICE_SRS: PRICE
        })],
        ignore_index=True
    )

    # Добавляем NIV Эльбрус
    elb_xl = pd.ExcelFile(f'{BASE_DIR}/Исходники/Прайс/Прайс Эльбрус.xlsx')
    elb_df = elb_xl.parse('Цены от 2 млн', skiprows=SKIPROWS_ELB)[
        [EAN_ELB, PRODUCT_NAME_ELB, PRICE_ELB]
    ]
    elb_df = elb_df.rename(columns={
            EAN_ELB: EAN_LOC,
            PRODUCT_NAME_ELB: PRODUCT_NAME,
            PRICE_ELB: PRICE_ELB + PRICE_ELB_ADD
    })

    df = pd.concat([df, elb_df[[EAN_LOC, PRODUCT_NAME]]], ignore_index=True)
    df = pd.merge(
        df,
        elb_df[[EAN_LOC, PRICE_ELB + PRICE_ELB_ADD]],
        on=EAN_LOC,
        how='left'
    ).drop_duplicates(subset=[EAN_LOC])
    df.dropna(subset=[EAN_LOC], inplace=True)
    df = df.rename(columns={
        EAN_LOC: EAN,
        PRODUCT_NAME: PRODUCT,
        PRICE: BASE_PRICE,
        PRICE_ELB + PRICE_ELB_ADD: ELB_PRICE
    })
    save_to_excel(f'{BASE_DIR}/Результаты/Прайс.xlsx', df)


if __name__ == "__main__":
    main()
