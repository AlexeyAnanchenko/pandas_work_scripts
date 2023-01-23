"""Скрипт формирует реестр по факторам"""

import utils
utils.path_append()

import os
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
LINK_DATE = 'Сцепка с датой регистриации'
COL_REPORT = [
    LINK_FACTOR_NUM, LINK, FACTOR_PERIOD, FACTOR_NUM, DATE_CREATION,
    DATE_START, DATE_EXPIRATION, WHS, NAME_HOLDING, EAN, PRODUCT,
    PLAN_IN_NFE, DATE_REGISTRY, QUANT_REGISTRY
]


def check_present_registry():
    list_dir = os.listdir(RESULT_DIR)
    if TABLE_REGISTRY_FACTORS in list_dir:
        return False
    return True


def create_registry():
    df = pd.DataFrame(columns=COL_REPORT)
    save_to_excel(RESULT_DIR + TABLE_REGISTRY_FACTORS, df)
    return


def fix_changes():
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
    df_var[DATE_REGISTRY] = pd.Timestamp(pd.Timestamp.today().date())
    df_var[QUANT_REGISTRY] = df_var[VARIANCE]
    df_var = df_var.drop(labels=[PLAN_NFE, VARIANCE], axis=1)
    df_rgs = pd.concat([df_rgs, df_var], ignore_index=True)

    change_col = [
        FACTOR_PERIOD, DATE_CREATION, DATE_START, DATE_EXPIRATION, PLAN_IN_NFE
    ]
    df_rgs = df_rgs.drop(labels=change_col, axis=1)
    df_fct = df_fct.drop(labels=[PLAN_IN_NFE], axis=1)
    df_fct = df_fct.rename(columns={PLAN_NFE: PLAN_IN_NFE})
    change_col.append(LINK_FACTOR_NUM)
    df_rgs = df_rgs.merge(df_fct[change_col], on=LINK_FACTOR_NUM, how='left')
    df_rgs.dropna(subset=[FACTOR_PERIOD], inplace=True)
    return df_rgs


def group_duplicates(df):
    df[LINK_DATE] = df[LINK_FACTOR_NUM] + df[DATE_REGISTRY].map(str)
    df_group = df[[LINK_DATE, QUANT_REGISTRY]].groupby([LINK_DATE]).agg(
        {QUANT_REGISTRY: 'sum'}
    ).reset_index()
    df = df.drop(labels=[QUANT_REGISTRY], axis=1)
    df = df.merge(df_group, on=LINK_DATE, how='left')
    df = df.drop(labels=[LINK_DATE], axis=1)
    df = df.reindex(columns=COL_REPORT).drop_duplicates()
    return df


def main():
    if check_present_registry():
        create_registry()
    df = group_duplicates(fix_changes())
    save_to_excel(RESULT_DIR + TABLE_REGISTRY_FACTORS, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
