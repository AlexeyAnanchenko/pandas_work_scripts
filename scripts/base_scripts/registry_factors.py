"""Скрипт формирует реестр по факторам"""

import utils
utils.path_append()

import pandas as pd

from settings import RESULT_DIR, TABLE_REGISTRY_FACTORS, TABLE_FACTORS
from settings import LINK, FACTOR_NUM, FACTOR_PERIOD, CURRENT, FUTURE
from settings import FACTOR_STATUS, ACTIVE_STATUS, PURPOSE_PROMO, EAN
from settings import DATE_START, DATE_CREATION, DATE_EXPIRATION, WHS
from settings import INACTIVE_PURPOSE, NAME_HOLDING, PLAN_NFE, PRODUCT
from settings import DATE_REGISTRY, QUANT_REGISTRY
from service import get_data, save_to_excel, print_complete

LINK_FACTOR_NUM = 'Сцепка Номер фактора-Склад-ШК'
PLAN_IN_NFE = 'План в NFE, шт'
VARIANCE = 'Различие, шт'


def get_table():
    df_rgs = get_data(TABLE_REGISTRY_FACTORS)
    df_fct = get_data(TABLE_FACTORS)
    df_fct = df_fct[
        (df_fct[FACTOR_PERIOD].isin([CURRENT, FUTURE]))
        & (df_fct[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df_fct[PURPOSE_PROMO] != INACTIVE_PURPOSE)
    ]
    df_fct = df_fct[[
        LINK, FACTOR_PERIOD, FACTOR_NUM, DATE_CREATION, DATE_START,
        DATE_EXPIRATION, WHS, NAME_HOLDING, EAN, PRODUCT, PLAN_NFE
    ]]
    df_fct.insert(
        0, LINK_FACTOR_NUM,
        df_fct[FACTOR_NUM].map(str) + df_fct[WHS] + df_fct[EAN].map(str)
    )
    df_fct = df_fct.merge(
        df_rgs[[LINK_FACTOR_NUM, PLAN_IN_NFE]].drop_duplicates(
            subset=LINK_FACTOR_NUM
        ),
        on=LINK_FACTOR_NUM, how='left'
    )
    df_fct.loc[df_fct[PLAN_IN_NFE].isnull(), PLAN_IN_NFE] = 0
    df_fct[VARIANCE] = df_fct[PLAN_NFE] - df_fct[PLAN_IN_NFE]
    df_var = df_fct[df_fct[VARIANCE] != 0].copy()
    df_var[DATE_REGISTRY] = pd.Timestamp.today()
    idx = df_var[df_var[PLAN_IN_NFE] != 0].index
    df_var.loc[idx, DATE_CREATION] = df_var.loc[idx, DATE_REGISTRY]
    df_var[QUANT_REGISTRY] = df_var[VARIANCE]
    df_var = df_var.drop(labels=[PLAN_NFE, VARIANCE], axis=1)
    df_rgs = pd.concat([df_rgs, df_var], ignore_index=True)

    change_col = [FACTOR_PERIOD, DATE_START, DATE_EXPIRATION, PLAN_IN_NFE]
    df_rgs = df_rgs.drop(labels=change_col, axis=1)
    df_fct = df_fct.drop(labels=[PLAN_IN_NFE], axis=1)
    df_fct = df_fct.rename(columns={PLAN_NFE: PLAN_IN_NFE})
    change_col.append(LINK_FACTOR_NUM)
    df_rgs = df_rgs.merge(df_fct[change_col], on=LINK_FACTOR_NUM, how='left')
    df_rgs.dropna(subset=[FACTOR_PERIOD], inplace=True)
    return df_rgs


def main():
    df = get_table()
    save_to_excel(RESULT_DIR + TABLE_REGISTRY_FACTORS, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
