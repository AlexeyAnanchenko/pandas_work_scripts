"""
Отчёт по выявлению недогруженных объёмов по факторам

"""
import utils
utils.path_append()

import numpy as np

from service import save_to_excel, get_data, print_complete
from settings import REPORT_UNDERLOADING, REPORT_DIR, TABLE_FACTORS, PAST
from settings import CURRENT, FACTOR_PERIOD, PURPOSE_PROMO, INACTIVE_PURPOSE
from settings import SALES_FACTOR_PERIOD, LINK, LINK_HOLDING, WHS, NAME_HOLDING
from settings import EAN, PRODUCT, LEVEL_3, PLAN_NFE, ADJUSTMENT_PBI, FACT_NFE
from settings import DESCRIPTION, USER
from settings import SALES_CURRENT_FOR_PAST


LINK_PERIOD = 'Сцепка Склад-Холдинг-Штрихкод-Период'
PLAN = 'План, шт'
SALES_TOTAL = 'Продажи Тотал, шт'
SALES = 'Продажи в пределах фактора, шт'


def get_factors():
    """Получаем начальные данные для отчёта"""
    df = get_data(TABLE_FACTORS)
    df = df[
        (df[FACTOR_PERIOD].isin([PAST, CURRENT]))
        & (df[PURPOSE_PROMO] != INACTIVE_PURPOSE)
    ]
    df[LINK_PERIOD] = df[LINK_HOLDING] + df[FACTOR_PERIOD]
    group_df = df.groupby([
        LINK_PERIOD, LINK, LINK_HOLDING, FACTOR_PERIOD, WHS,
        NAME_HOLDING, EAN, PRODUCT, LEVEL_3
    ]).agg({
        PLAN_NFE: 'max',
        ADJUSTMENT_PBI: 'max',
        FACT_NFE: 'max',
        SALES_FACTOR_PERIOD: 'max',
        SALES_CURRENT_FOR_PAST: 'max',
    })
    df = df[[
        LINK_PERIOD, DESCRIPTION, USER
    ]].drop_duplicates(subset=LINK_PERIOD)
    group_df = group_df.merge(df, on=LINK_PERIOD, how='left')
    return group_df


def gen_plan(df):
    """Формируем итоговый столбец плана"""
    df[PLAN] = np.maximum(
        df[PLAN_NFE],
        (df[ADJUSTMENT_PBI] * 0.7)
    ).round(0)
    return df


def gen_sales(df):
    """Формируем итоговые столбцы продаж"""
    df[SALES_TOTAL] = df[SALES_FACTOR_PERIOD]
    idx = df[df[FACTOR_PERIOD == PAST]]
    df.loc[idx, SALES_TOTAL] = (df.loc[idx, SALES_FACTOR_PERIOD]
                                + df.loc[idx, SALES_CURRENT_FOR_PAST])
    df[SALES] = df[[PLAN, SALES_TOTAL]].min(axis=1)
    return df


def main():
    df = gen_plan(get_factors())
    save_to_excel(REPORT_DIR + REPORT_UNDERLOADING, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
