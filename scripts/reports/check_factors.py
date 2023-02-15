"""
Отчёт по проверке факторов согласованных и новых

"""
import utils
utils.path_append()

from service import save_to_excel, get_data, print_complete
from settings import REPORT_CHECK_FACTORS, REPORT_DIR_FINAL, TABLE_FACTORS
from settings import FACTOR_PERIOD, CURRENT, PURPOSE_PROMO, INACTIVE_PURPOSE
from settings import LINK, FUTURE, NAME_HOLDING, RSV_FACTOR_PERIOD_CURRENT
from settings import SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD, FACTOR_NUM
from settings import LINK_HOLDING, RSV_FACTOR_PERIOD_FUTURE, CUTS_PBI
from settings import ADJUSTMENT_PBI, FACT_NFE, NAME_TRAD, TABLE_EXCLUDE
from settings import FACTOR_STATUS, ACTIVE_STATUS, SALES_PBI, RESERVES_PBI


LINK_HOLDING_PERIOD = 'Сцепка Период-Склад-Холдинг-ШК'
LINK_FACTOR = 'Сцепка Номер фактора-Склад-Штрихкод'
CHECK_FACT = 'ПРОВЕРКА ФАКТА ТЕКУЩЕГО ПЕРИОДА'
CHECK_DUPL = 'ПРОВЕРКА НАЛИЧИЯ ДУБЛИКАТОВ'
CANCEL_STATUS = 'Отменен(а)'

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
        & (df[NAME_HOLDING] != NAME_TRAD)
        & (~df[LINK_FACTOR].isin(df_exclude[CHECK_FACT].to_list())),
        CHECK_FACT
    ] = 'Факт ниже продаж и резервов!'
    df.loc[
        (df[FACT_NFE] - df[CUTS_PBI] < (df[SALES_FACTOR_PERIOD]
                                        + df[RSV_FACTOR_PERIOD_CURRENT]))
        & (~df[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df[FACTOR_PERIOD] == CURRENT)
        & (df[NAME_HOLDING] != NAME_TRAD)
        & (~df[LINK_FACTOR].isin(df_exclude[CHECK_FACT].to_list())),
        CHECK_FACT
    ] = 'Факт ниже продаж и резервов!'
    return df


def check_duplicates(df):
    duplicate_df = df[
        ~df[LINK_FACTOR].isin(df_exclude[CHECK_DUPL].to_list())
    ].copy()
    duplicate_df = duplicate_df[duplicate_df.duplicated(
        [LINK_HOLDING_PERIOD], keep='first'
    )][LINK_HOLDING_PERIOD].to_list()
    df.loc[df[LINK_HOLDING_PERIOD].isin(duplicate_df), CHECK_DUPL] = 'Дубликат'
    df = df.drop(labels=[LINK_HOLDING_PERIOD], axis=1)
    return df


def main():
    df = check_duplicates(check_fact(get_factors()))
    save_to_excel(REPORT_DIR_FINAL + REPORT_CHECK_FACTORS, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
