"""Скрипт формирует реестр по факторам"""

import utils
utils.path_append()

import os
import re
import pandas as pd
import datetime as dt

from settings import RESULT_DIR, TABLE_REGISTRY_FACTORS, TABLE_FACTORS
from settings import LINK, FACTOR_NUM, FACTOR_PERIOD, CURRENT, FUTURE
from settings import FACTOR_STATUS, ACTIVE_STATUS, PURPOSE_PROMO, EAN
from settings import DATE_START, DATE_CREATION, DATE_EXPIRATION, WHS
from settings import INACTIVE_PURPOSE, NAME_HOLDING, PLAN_NFE, PRODUCT
from settings import DATE_REGISTRY, QUANT_REGISTRY, ARCHIVE_DIR, HOLDING
from settings import TABLE_HOLDINGS, TABLE_FIXING_FACTORS, FIRST_PLAN
from settings import MAX_PLAN
from hidden_settings import OPT_CLIENTS
from service import get_data, save_to_excel, print_complete

LINK_FACTOR_NUM = 'Сцепка Номер фактора-Склад-ШК'
PLAN_IN_NFE = 'План в NFE, шт'
VARIANCE = 'Различие, шт'
LINK_DATE = 'Сцепка с датой регистриации'
COL_REPORT = [
    LINK_FACTOR_NUM, LINK, FACTOR_PERIOD, FACTOR_NUM, DATE_CREATION,
    DATE_START, DATE_EXPIRATION, WHS, NAME_HOLDING, HOLDING, EAN, PRODUCT,
    PLAN_IN_NFE, DATE_REGISTRY, QUANT_REGISTRY
]
SUBSTRACT = 'Отнять, шт'
RESULT_COL = 'Расчётная колонка, шт'
MAX_PLAN_ITEM = 'Максимальный план, временный'


def get_archive():
    list_dir = os.listdir(ARCHIVE_DIR)
    list_dir = [dt.datetime.strptime(i, '%d.%m.%Y').date() for i in list_dir]
    list_dir.sort()
    list_dir = [dt.date.strftime(i, '%d.%m.%Y') for i in list_dir]
    folder_path = ARCHIVE_DIR + list_dir.pop() + '\\'
    list_dir = os.listdir(folder_path)
    archive_file = None
    for file in list_dir:
        if re.search(TABLE_REGISTRY_FACTORS, file):
            archive_file = folder_path + file
    excel = pd.ExcelFile(archive_file)
    return excel.parse()


def fix_changes(df_rgs):
    # Удаляем код холдинга
    if HOLDING in df_rgs.columns:
        df_rgs = df_rgs.drop(labels=[HOLDING], axis=1)

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

    replace_col = [
        FACTOR_PERIOD, DATE_CREATION, DATE_START,
        DATE_EXPIRATION, PLAN_IN_NFE, NAME_HOLDING
    ]
    df_rgs = df_rgs.drop(labels=replace_col, axis=1)
    df_fct = df_fct.drop(labels=[PLAN_IN_NFE], axis=1)
    df_fct = df_fct.rename(columns={PLAN_NFE: PLAN_IN_NFE})
    replace_col.append(LINK_FACTOR_NUM)
    df_rgs = df_rgs.merge(df_fct[replace_col], on=LINK_FACTOR_NUM, how='left')
    df_rgs.dropna(subset=[FACTOR_PERIOD], inplace=True)
    return df_rgs


def subtracting_factors(df):
    df[LINK_DATE] = df[LINK_FACTOR_NUM] + df[DATE_REGISTRY].map(str)
    df = df.sort_values(
        by=[FACTOR_NUM, DATE_REGISTRY],
        ascending=[False, False]
    )
    df_subtr = df[df[QUANT_REGISTRY] < 0].copy()
    df_subtr[SUBSTRACT] = df_subtr[QUANT_REGISTRY]
    df = df[df[QUANT_REGISTRY] > 0]

    list_link = list(set(df_subtr[LINK_FACTOR_NUM].to_list()))
    start_df = df[df[LINK_FACTOR_NUM].isin(list_link)][[
        LINK_DATE, LINK_FACTOR_NUM, QUANT_REGISTRY
    ]]
    start_df = start_df.merge(
        df_subtr[[LINK_FACTOR_NUM, SUBSTRACT]],
        on=LINK_FACTOR_NUM, how='left'
    )
    result_df = pd.DataFrame(columns=[
        LINK_DATE, LINK_FACTOR_NUM, QUANT_REGISTRY, SUBSTRACT, RESULT_COL
    ])

    flag = True
    while flag:
        mid_df = start_df.copy().drop_duplicates(subset=LINK_FACTOR_NUM)
        mid_df[RESULT_COL] = mid_df[QUANT_REGISTRY] + mid_df[SUBSTRACT]
        result_df = pd.concat([result_df, mid_df], ignore_index=True)
        result_df.loc[result_df[RESULT_COL] < 0, RESULT_COL] = 0
        if not len(mid_df.loc[mid_df[SUBSTRACT] < 0]) > 0:
            flag = False
        else:
            row_drop = mid_df[LINK_DATE].to_list()
            start_df = start_df[~start_df.isin(row_drop).any(axis=1)]
            start_df = start_df.drop(labels=[SUBSTRACT], axis=1)
            mid_df = mid_df.drop(labels=[SUBSTRACT], axis=1)
            mid_df = mid_df.rename(columns={RESULT_COL: SUBSTRACT})
            start_df = start_df.merge(
                mid_df[[LINK_FACTOR_NUM, SUBSTRACT]],
                on=LINK_FACTOR_NUM, how='left'
            )
            start_df = start_df[start_df[SUBSTRACT] < 0]

    df = df.merge(result_df[[LINK_DATE, RESULT_COL]], on=LINK_DATE, how='left')
    idx = df.loc[df[RESULT_COL].notna()].index
    df.loc[idx, QUANT_REGISTRY] = df.loc[idx, RESULT_COL]
    df = df[df[QUANT_REGISTRY] > 0]
    return df


def merge_holdings(df):
    holdings = get_data(TABLE_HOLDINGS)[[NAME_HOLDING, HOLDING]]
    df = df.merge(holdings.drop_duplicates(), on=NAME_HOLDING, how='left')
    clients = list(set(df[NAME_HOLDING].tolist()))

    for client in clients:
        if OPT_CLIENTS in client:
            df.loc[df[NAME_HOLDING] == client, HOLDING] = 'ОПТ'

    df = utils.void_to(df, HOLDING, '---')
    return df[COL_REPORT]


def sort_and_test(df):
    df = df.sort_values(
        by=[DATE_REGISTRY, FACTOR_NUM],
        ascending=[False, False]
    )
    df_test = df[[LINK_FACTOR_NUM, PLAN_IN_NFE, QUANT_REGISTRY]].groupby([
        LINK_FACTOR_NUM
    ]).agg({
        PLAN_IN_NFE: 'max',
        QUANT_REGISTRY: 'sum'
    }).reset_index()
    if df_test[PLAN_IN_NFE].eq(df_test[QUANT_REGISTRY]).all():
        return df
    save_to_excel(RESULT_DIR + 'ОТКРОЙ МЕНЯ.xlsx', df)
    raise Exception('РЕЕСТР ФОРМИРУЕТСЯ ОШИБОЧНО!!!')


def gen_fixing_factors(df_reg):
    df_reg = df_reg[[
        LINK_FACTOR_NUM, FACTOR_NUM, WHS, EAN, PLAN_IN_NFE
    ]].drop_duplicates()
    df = get_data(TABLE_FIXING_FACTORS)
    df = df.merge(
        df_reg, on=[LINK_FACTOR_NUM, FACTOR_NUM, WHS, EAN], how='outer'
    )
    idx = df[df[FIRST_PLAN].isnull()].index
    df.loc[idx, FIRST_PLAN] = df.loc[idx, PLAN_IN_NFE]
    idx = df[df[MAX_PLAN].isnull()].index
    df.loc[idx, MAX_PLAN] = 0
    df[MAX_PLAN_ITEM] = df[[MAX_PLAN, PLAN_IN_NFE]].max(axis=1)
    idx = df[~df[MAX_PLAN].isnull()].index
    df.loc[idx, MAX_PLAN] = df.loc[idx, MAX_PLAN_ITEM]
    df = df.drop(labels=[PLAN_IN_NFE, MAX_PLAN_ITEM], axis=1)
    return df


def main():
    df = merge_holdings(subtracting_factors(fix_changes(get_archive())))
    df = sort_and_test(df)
    save_to_excel(RESULT_DIR + TABLE_REGISTRY_FACTORS, df)
    df = gen_fixing_factors(df)
    save_to_excel(RESULT_DIR + TABLE_FIXING_FACTORS, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
