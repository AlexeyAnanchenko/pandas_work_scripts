"""
Отчёт по проверке факторов согласованных и новых

"""
import utils
utils.path_append()

from hidden_settings import WHS_ELBRUS
from service import save_to_excel, get_data, print_complete
from settings import REPORT_CHECK_FACTORS, REPORT_DIR_FINAL, TABLE_FACTORS
from settings import FACTOR_PERIOD, CURRENT, PURPOSE_PROMO, INACTIVE_PURPOSE
from settings import LINK, FUTURE, NAME_HOLDING, RSV_FACTOR_PERIOD_CURRENT
from settings import SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD, FACTOR_NUM, WHS
from settings import LINK_HOLDING, RSV_FACTOR_PERIOD_FUTURE, CUTS_PBI, MSU
from settings import ADJUSTMENT_PBI, FACT_NFE, NAME_TRAD, TABLE_EXCLUDE
from settings import FACTOR_STATUS, ACTIVE_STATUS, SALES_PBI, RESERVES_PBI
from settings import TABLE_ASSORTMENT, TABLE_DIRECTORY, EAN, MATRIX, MATRIX_LY
from settings import ELB_PRICE, BASE_PRICE, PLAN_NFE, ALL_CLIENTS
from settings import AVG_FACTOR_PERIOD


LINK_HOLDING_PERIOD = 'Сцепка Период-Склад-Холдинг-ШК'
LINK_FACTOR = 'Сцепка Номер фактора-Склад-Штрихкод'
CHECK_FACT = 'ПРОВЕРКА ФАКТА ТЕКУЩЕГО ПЕРИОДА'
CHECK_DUPL = 'ПРОВЕРКА НАЛИЧИЯ ДУБЛИКАТОВ'
CANCEL_STATUS = 'Отменен(а)'
ACTIVE_LOC = 'Активный ассортимент для локации'
PRICE_LOC = 'Цена для локации (GIV/NIV), руб.'
L_YUG = 'Лоджистик-Юг ELBR (5553395)'
YES = 'Да'
NO = 'Нет'
PLAN_MSU = 'План, msu'
PLAN_PRICE = 'План, руб.'
AVARAGE_MSU = 'Средние продажи, msu'
AVARAGE_PRICE = 'Средние продажи, руб.'

df_exclude = get_data(TABLE_EXCLUDE)


def get_factors():
    """Получаем факторы для проверки"""
    df = get_data(TABLE_FACTORS)
    df = df[
        (df[FACTOR_PERIOD].isin([CURRENT, FUTURE]))
        & (df[PURPOSE_PROMO] != INACTIVE_PURPOSE)
        & (df[FACTOR_STATUS] != CANCEL_STATUS)
    ]
    df[RSV_FACTOR_PERIOD_FUTURE] = (df[RSV_FACTOR_PERIOD]
                                    - df[RSV_FACTOR_PERIOD_CURRENT])
    df.insert(0, LINK_FACTOR, df[FACTOR_NUM].map(str) + df[LINK])
    df.insert(0, LINK_HOLDING_PERIOD, df[FACTOR_PERIOD] + df[LINK_HOLDING])
    df = df.drop(labels=[LINK_HOLDING, PURPOSE_PROMO, ADJUSTMENT_PBI], axis=1)
    return df


def check_fact(df):
    """Проверяем факт продаж по заявке"""
    df.loc[
        (df[SALES_PBI] + df[RESERVES_PBI] < (df[SALES_FACTOR_PERIOD]
                                             + df[RSV_FACTOR_PERIOD_CURRENT]))
        & (df[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df[FACTOR_PERIOD] == CURRENT)
        & (~df[NAME_HOLDING].isin([NAME_TRAD, ALL_CLIENTS]))
        & (~df[LINK_FACTOR].isin(df_exclude[CHECK_FACT].to_list())),
        CHECK_FACT
    ] = 'Факт ниже продаж и резервов!'
    df.loc[
        (df[FACT_NFE] - df[CUTS_PBI] < (df[SALES_FACTOR_PERIOD]
                                        + df[RSV_FACTOR_PERIOD_CURRENT]))
        & (~df[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df[FACTOR_PERIOD] == CURRENT)
        & (~df[NAME_HOLDING].isin([NAME_TRAD, ALL_CLIENTS]))
        & (~df[LINK_FACTOR].isin(df_exclude[CHECK_FACT].to_list())),
        CHECK_FACT
    ] = 'Факт ниже продаж и резервов!'
    return df


def check_duplicates(df):
    """Проверяем дубликаты по периоду"""
    duplicate_df = df[
        ~df[LINK_FACTOR].isin(df_exclude[CHECK_DUPL].to_list())
    ].copy()
    duplicate_df = duplicate_df[duplicate_df.duplicated(
        [LINK_HOLDING_PERIOD], keep='first'
    )][LINK_HOLDING_PERIOD].to_list()
    df.loc[df[LINK_HOLDING_PERIOD].isin(duplicate_df), CHECK_DUPL] = 'Дубликат'
    df = df.drop(labels=[LINK_HOLDING_PERIOD], axis=1)
    return df


def merge_assort_and_dir(df):
    """Подтягиваем ассортимент и и данные их справочника по ШК"""
    assort = get_data(TABLE_ASSORTMENT)[[LINK]]
    assort[ACTIVE_LOC] = YES
    df = df.merge(assort, on=LINK, how='left')

    columns = [MATRIX, MATRIX_LY, MSU, ELB_PRICE, BASE_PRICE]
    direct = get_data(TABLE_DIRECTORY)[[EAN] + columns]
    df = df.merge(direct, on=EAN, how='left')
    df.loc[df[ACTIVE_LOC].isnull(), ACTIVE_LOC] = NO
    df.loc[
        (df[NAME_HOLDING] == L_YUG) & (df[MATRIX_LY] == NO), ACTIVE_LOC
    ] = NO
    df.loc[
        (df[NAME_HOLDING] != L_YUG) & (df[MATRIX] == NO), ACTIVE_LOC
    ] = NO
    idx = df[df[WHS].isin(WHS_ELBRUS.keys())].index
    df.loc[idx, PRICE_LOC] = df.loc[idx, ELB_PRICE]
    idx = df[~df[WHS].isin(WHS_ELBRUS.keys())].index
    df.loc[idx, PRICE_LOC] = df.loc[idx, BASE_PRICE]

    df[PLAN_MSU] = df[PLAN_NFE] * df[MSU]
    df[PLAN_PRICE] = df[PLAN_NFE] * df[PRICE_LOC]
    df[AVARAGE_MSU] = df[AVG_FACTOR_PERIOD] * df[MSU]
    df[AVARAGE_PRICE] = df[AVG_FACTOR_PERIOD] * df[PRICE_LOC]
    df = df.drop(labels=columns + [PRICE_LOC], axis=1)
    return df


def main():
    df = merge_assort_and_dir(check_duplicates(check_fact(get_factors())))
    save_to_excel(REPORT_DIR_FINAL + REPORT_CHECK_FACTORS, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
