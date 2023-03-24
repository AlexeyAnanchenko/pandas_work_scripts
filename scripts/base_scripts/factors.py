"""
Формирует отчёт по факторам в удобном формате

"""
import utils
utils.path_append()

import os
import pandas as pd
import datetime as dt
from dateutil import relativedelta

from hidden_settings import WAREHOUSE_FACTORS, prefix
from service import save_to_excel, print_complete, get_data
from service import get_mult_clients_dict
from settings import SOURCE_DIR, RESULT_DIR, TABLE_FACTORS, WHS, FACTOR
from settings import FACTOR_NUM, REF_FACTOR, DATE_EXPIRATION, FACTOR_PERIOD
from settings import FACTOR_STATUS, DATE_CREATION, DATE_START, NAME_HOLDING
from settings import EAN, PRODUCT, LEVEL_3, DESCRIPTION, USER, PLAN_NFE
from settings import FACT_NFE, ADJUSTMENT_PBI, SALES_PBI, RESERVES_PBI, QUOTA
from settings import CUTS_PBI, LINK, LINK_HOLDING, PURPOSE_PROMO, TABLE_SALES
from settings import ALL_CLIENTS, SOFT_HARD_RSV, SALES_FACTOR_PERIOD, NAME_TRAD
from settings import AVG_FACTOR_PERIOD, RSV_FACTOR_PERIOD, PAST, CURRENT
from settings import FUTURE, TABLE_SALES_HOLDINGS, TABLE_RESERVE, TOTAL_RSV
from settings import AVG_FACTOR_PERIOD_WHS, SOFT_HARD_RSV_CURRENT, SOFT_RSV
from settings import RSV_FACTOR_PERIOD_CURRENT, SALES_CURRENT_FOR_PAST
from settings import RSV_FACTOR_PERIOD_TOTAL, TABLE_SALES_BY_DATE, HARD_RSV
from settings import SALES_BY_DATE, CUTS_BY_DATE, DATE_SALES, PG_PROGRAMM
from settings import EXPECTED_DATE, TABLE_RSV_BY_DATE, EXCLUDE_STRING, WORD_YES
from settings import SOFT_RSV_BY_DATE, HARD_RSV_BY_DATE, QUOTA_BY_DATE
from settings import TABLE_FIXING_FACTORS, FIRST_PLAN, MAX_PLAN, DEL_ROW
from settings import TABLE_DIRECTORY
from registry_factors import DATE_CREATION_LAST, DATE_START_LAST
from registry_factors import DATE_EXPIRATION_LAST, NAME_HOLDING_LAST
from registry_factors import LINK_FACTOR_NUM


SOURCE_FILE = 'NovoForecastServer_РезультатыПоиска.xlsx'
SOURCE_FILE_PB = 'Статистика факторов PG.xlsx'
SOURCE_FILE_PROMO = 'Акции.xlsx'
WORKING_FACTORS = ['Акция', 'Предзаказ', 'Тендер']
TOWN = 'Город'
TYPE_FACTOR = 'Тип'
ORDER_LOC = 'Сумма Заказ'
SALES_LOC = 'Сумма Продажи'
CUTS_LOC = 'Сумма Урезания'
RESERVES_LOC = 'Сумма Резервы'
HOLDING_LOC = 'Холдинг'
EAN_LOC = 'EAN'
PRODUCT_LOC = 'Наименование продукта'
LEVEL_3_LOC = 'Группа продукта'
LINK_FACTOR = 'Сцепка номер фактора-холдинг-EAN'
FACTOR_NUM_PB = 'Номер акции'
PLAN_LOC = 'План'
FACT_LOC = 'Факт'
USER_LOC = 'Ответственный пользователь'
FACTOR_LOC = 'Фактор'
REF_FACTOR_LOC = 'Ссылка на фактор'
FACTOR_STATUS_LOC = 'Статус'
DATE_CREATION_LOC = 'Дата создания фактора'
DATE_START_LOC = 'Дата начала действия фактора'
DATE_EXPIRATION_LOC = 'Дата окончания действия фактора'
DESCRIPTION_LOC = 'Описание'
NUMBER_PROMO = 'Номер'
PURPOSE = 'Цель'
TRAD_HOLDINGS = [
    'Васильева Татьяна Викторовна ИП (8108456)',
    'Воронина Елена Сергеевна ИП (8179862)',
    'Петровская Татьяна Анатольевна ИП (8143465)'
]
INDEXES = 'Текущие индексы строк'
LINK_UNIQUE = 'Уникальная сцепка'
CHANNEL_TRAD = 'Канал продаж'
NOT_DATA = '<НЕТ ДАННЫХ>'
DEL_FACTOR = 'Удалённый фактор'


def filtered_factors(file=SOURCE_FILE, column=FACTOR, skip=0):
    """Получить отфильтрованный dataframe по нужным факторам"""
    df = pd.ExcelFile(SOURCE_DIR + file).parse(skiprows=skip)
    df = df[df[column].isin(WORKING_FACTORS)]
    df = df[df[TOWN].isin(WAREHOUSE_FACTORS)]
    df = df.rename(columns={TOWN: WHS})
    return df


def add_num_factors(df):
    """Функция добавляет номер фактора и преобразует ссылки"""
    num_factors = []
    ref = {}
    ch_1 = '>'
    ch_2 = '<'
    ch_3 = 'href="'
    ch_4 = '">'
    for row in df[REF_FACTOR_LOC]:
        num_factors.append(row[(row.find(ch_1) + 1):row.find(ch_2, 1)])
        ref[row] = prefix + row[(row.find(ch_3) + len(ch_3)):row.find(ch_4)]
    df[FACTOR_NUM] = num_factors
    df = df.replace({REF_FACTOR_LOC: ref})
    return df


def add_deleted_factors(df):
    """Добавляет удалённые факторы из реестра"""
    df_reg = get_data(TABLE_FIXING_FACTORS)
    df_reg = df_reg.drop(labels=[LINK_FACTOR_NUM], axis=1)
    df_reg = df_reg[
        df_reg[DATE_START_LAST] >= pd.to_datetime(utils.get_factor_start())
    ]
    df_reg = df_reg.rename(columns={
        DATE_CREATION_LAST: DATE_CREATION_LOC,
        DATE_START_LAST: DATE_START_LOC,
        DATE_EXPIRATION_LAST: DATE_EXPIRATION_LOC,
        NAME_HOLDING_LAST: HOLDING_LOC,
        EAN: EAN_LOC
    })
    df[FACTOR_NUM] = df[FACTOR_NUM].astype(int)
    df = df.merge(
        df_reg,
        on=[
            FACTOR_NUM, DATE_CREATION_LOC, DATE_START_LOC, DATE_EXPIRATION_LOC,
            HOLDING_LOC, WHS, EAN_LOC
        ],
        how='outer'
    )

    df_empty = df[df[FACTOR_LOC].isnull()].copy()
    df = df[~df[FACTOR_LOC].isnull()]
    labels = [FACTOR_LOC, USER_LOC, REF_FACTOR_LOC]
    df_empty = df_empty.drop(labels=labels, axis=1)
    df_empty = df_empty.merge(
        df[labels + [FACTOR_NUM]].drop_duplicates(),
        on=FACTOR_NUM, how='left')
    df = pd.concat([df, df_empty], ignore_index=True)

    df = utils.void_to(df, FACTOR_LOC, DEL_FACTOR)
    df = utils.void_to(df, REF_FACTOR_LOC, NOT_DATA)
    df = utils.void_to(df, USER_LOC, NOT_DATA)
    df = utils.void_to(df, FACTOR_STATUS_LOC, DEL_ROW)
    for col in [PLAN_LOC, FACT_LOC, FIRST_PLAN, MAX_PLAN]:
        df = utils.void_to(df, col, 0)

    df_dir = get_data(TABLE_DIRECTORY)[[EAN, PRODUCT, LEVEL_3]]
    df_dir = df_dir.rename(columns={EAN: EAN_LOC})
    df = df.merge(df_dir, on=EAN_LOC, how='left')
    for col in [PRODUCT, LEVEL_3]:
        df = utils.void_to(df, col, NOT_DATA)
    idx = df[df[PRODUCT_LOC].isnull()].index
    df.loc[idx, PRODUCT_LOC] = df.loc[idx, PRODUCT]
    df.loc[idx, LEVEL_3_LOC] = df.loc[idx, LEVEL_3]
    df = df.drop(labels=[PRODUCT, LEVEL_3], axis=1)
    return df


def add_pbi_and_purpose(df):
    """Добавляет столбцы из отчёта в PBI по факторам и цель акции"""
    df_pb = filtered_factors(SOURCE_FILE_PB, TYPE_FACTOR, 2)
    df_pb = utils.void_to(df_pb, HOLDING_LOC, ALL_CLIENTS)
    df = utils.void_to(df, HOLDING_LOC, ALL_CLIENTS)
    df_pb.insert(
        0, LINK_FACTOR,
        df_pb[FACTOR_NUM_PB].map(str) + df_pb[EAN_LOC].map(str)
    )
    df.insert(
        0, LINK_FACTOR,
        df[FACTOR_NUM].map(str) + df[EAN_LOC].map(str)
    )
    df = df.merge(
        df_pb[[LINK_FACTOR, ORDER_LOC, SALES_LOC, RESERVES_LOC, CUTS_LOC]],
        on=LINK_FACTOR, how='left'
    )
    df[LINK] = df[WHS] + df[EAN_LOC].map(str)
    df[LINK_HOLDING] = df[WHS] + df[HOLDING_LOC] + df[EAN_LOC].map(str)

    df_promo = pd.ExcelFile(SOURCE_DIR + SOURCE_FILE_PROMO).parse()
    df_promo = df_promo.rename(columns={NUMBER_PROMO: FACTOR_NUM})
    df[FACTOR_NUM] = df[FACTOR_NUM].astype(int)
    df = df.merge(
        df_promo[[FACTOR_NUM, PURPOSE]],
        on=FACTOR_NUM, how='left'
    )
    return df


def split_by_month(df):
    """Разделяет факторы на текущие, будущие и прошедшие в отдельном столбце"""
    today = dt.date.today() - dt.timedelta(days=1)
    current_month = today.replace(day=1)
    current_month_dt = dt.datetime.combine(current_month, dt.time(0, 0))
    next_month = today + relativedelta.relativedelta(months=1, day=1)
    next_month_dt = dt.datetime.combine(next_month, dt.time(0, 0))
    df[FACTOR_PERIOD] = FUTURE
    idx = df[df[DATE_EXPIRATION_LOC] < next_month_dt].index
    df.loc[idx, FACTOR_PERIOD] = CURRENT
    idx = df[df[DATE_EXPIRATION_LOC] < current_month_dt].index
    df.loc[idx, FACTOR_PERIOD] = PAST
    return df


def reindex_rename(df):
    """Упорядочивает и переименовывает заголовок таблицы"""
    result_df = df[[
        LINK, LINK_HOLDING, FACTOR_LOC, REF_FACTOR_LOC, FACTOR_NUM,
        FACTOR_STATUS_LOC, PURPOSE, FACTOR_PERIOD, DATE_CREATION_LOC,
        DATE_START_LOC, DATE_EXPIRATION_LOC, WHS, HOLDING_LOC, EAN_LOC,
        PRODUCT_LOC, LEVEL_3_LOC, DESCRIPTION_LOC, USER_LOC, FIRST_PLAN,
        MAX_PLAN, PLAN_LOC, FACT_LOC, ORDER_LOC, SALES_LOC, RESERVES_LOC,
        CUTS_LOC
    ]].rename(columns={
        FACTOR_LOC: FACTOR, REF_FACTOR_LOC: REF_FACTOR,
        FACTOR_STATUS_LOC: FACTOR_STATUS, DATE_CREATION_LOC: DATE_CREATION,
        DATE_START_LOC: DATE_START, DATE_EXPIRATION_LOC: DATE_EXPIRATION,
        HOLDING_LOC: NAME_HOLDING, EAN_LOC: EAN, PRODUCT_LOC: PRODUCT,
        LEVEL_3_LOC: LEVEL_3, DESCRIPTION_LOC: DESCRIPTION, USER_LOC: USER,
        PLAN_LOC: PLAN_NFE, FACT_LOC: FACT_NFE, ORDER_LOC: ADJUSTMENT_PBI,
        SALES_LOC: SALES_PBI, RESERVES_LOC: RESERVES_PBI, CUTS_LOC: CUTS_PBI,
        PURPOSE: PURPOSE_PROMO
    })
    return result_df


def merge_by_mult_clients(df, df_merge, mult_clients, static_col, numeric_col):
    """Подтягивает продажи и резервы, если df содержит мульти-клиентов"""
    df_pure = df[~df[NAME_HOLDING].isin(mult_clients)].copy()
    df_pure = df_pure.merge(
        df_merge[static_col + numeric_col], on=static_col, how='left'
    )
    numeric_col_dict = {}

    for col in numeric_col:
        numeric_col_dict[col] = 'sum'

    for client in mult_clients:
        df_mult = df[df[NAME_HOLDING] == client].copy()
        df_merge_copy = df_merge[df_merge[NAME_HOLDING].isin(
            mult_clients[client]
        )].copy()
        df_merge_copy = df_merge_copy.groupby([LINK]).agg(
            numeric_col_dict
        ).reset_index()
        df_mult = df_mult.merge(
            df_merge_copy[[LINK] + numeric_col], on=LINK, how='left'
        )
        df_pure = pd.concat([df_pure, df_mult], ignore_index=True)

    return df_pure


def add_sales_and_rsv(df):
    """Подтягивает продажи и резервы в рамках определённого периода фактора"""
    sales, col_sales = get_data(TABLE_SALES_HOLDINGS)
    rsv = get_data(TABLE_RESERVE)
    static_col = [LINK, LINK_HOLDING, WHS, EAN, NAME_HOLDING]
    num_col_sales = [
        col_sales['pntm_sale'],
        col_sales['last_sale'],
        col_sales['avg_cut_sale']
    ]
    num_col_rsv = [SOFT_HARD_RSV, SOFT_HARD_RSV_CURRENT, TOTAL_RSV]
    mult_clients = get_mult_clients_dict(df, NAME_HOLDING)

    if mult_clients:
        df = merge_by_mult_clients(
            df, sales, mult_clients, static_col, num_col_sales)
        df = merge_by_mult_clients(
            df, rsv, mult_clients, static_col, num_col_rsv)
    else:
        df = df.merge(
            sales[static_col + num_col_sales], on=static_col, how='left'
        )
        df = df.merge(rsv[static_col + num_col_rsv], on=static_col, how='left')

    idx = df[df[FACTOR_PERIOD] == CURRENT].index
    df.loc[idx, SALES_FACTOR_PERIOD] = df.loc[idx, col_sales['last_sale']]
    idx = df[df[FACTOR_PERIOD] == PAST].index
    df.loc[idx, SALES_FACTOR_PERIOD] = df.loc[idx, col_sales['pntm_sale']]
    df.loc[idx, SALES_CURRENT_FOR_PAST] = df.loc[idx, col_sales['last_sale']]
    idx = df[df[FACTOR_PERIOD] == FUTURE].index
    df.loc[idx, SALES_FACTOR_PERIOD] = 0
    df[AVG_FACTOR_PERIOD] = df[col_sales['avg_cut_sale']]
    df[RSV_FACTOR_PERIOD] = df[SOFT_HARD_RSV]
    df[RSV_FACTOR_PERIOD_CURRENT] = df[SOFT_HARD_RSV_CURRENT]
    df[RSV_FACTOR_PERIOD_TOTAL] = df[TOTAL_RSV]

    for col in num_col_sales + num_col_rsv:
        df.drop(col, axis=1, inplace=True)

    return df


def add_total_sales_rsv(df):
    """
    Заменяет данные полученные в 'add_sales_and_rsv', где холдинги не
    указаны, либо где холдинги традиционной торговли + добавляем средние
    продажи по складу
    """
    sales, col_sales = get_data(TABLE_SALES)
    rsv = get_data(TABLE_RESERVE)
    rsv = rsv.groupby([LINK]).agg({
        SOFT_HARD_RSV: 'sum',
        SOFT_HARD_RSV_CURRENT: 'sum',
        TOTAL_RSV: 'sum'
    }).reset_index()
    TRAD_HOLDINGS.append(ALL_CLIENTS)
    df = df.merge(
        sales[[
            LINK, col_sales['pntm_sale'], col_sales['last_sale'],
            col_sales['avg_cut_sale']
        ]],
        on=LINK,
        how='left'
    )
    df = df.merge(rsv, on=LINK, how='left')

    idx = df[
        (df[FACTOR_PERIOD] == CURRENT) & (df[NAME_HOLDING].isin(TRAD_HOLDINGS))
    ].index
    df.loc[idx, SALES_FACTOR_PERIOD] = df.loc[idx, col_sales['last_sale']]
    idx = df[
        (df[FACTOR_PERIOD] == PAST) & (df[NAME_HOLDING].isin(TRAD_HOLDINGS))
    ].index
    df.loc[idx, SALES_FACTOR_PERIOD] = df.loc[idx, col_sales['pntm_sale']]
    df.loc[idx, SALES_CURRENT_FOR_PAST] = df.loc[idx, col_sales['last_sale']]
    idx = df[
        (df[FACTOR_PERIOD] == FUTURE) & (df[NAME_HOLDING].isin(TRAD_HOLDINGS))
    ].index
    df.loc[idx, SALES_FACTOR_PERIOD] = 0

    idx = df[df[NAME_HOLDING].isin(TRAD_HOLDINGS)].index
    df.loc[idx, RSV_FACTOR_PERIOD] = df.loc[idx, SOFT_HARD_RSV]
    df.loc[idx, RSV_FACTOR_PERIOD_CURRENT] = df.loc[idx, SOFT_HARD_RSV_CURRENT]
    df.loc[idx, RSV_FACTOR_PERIOD_TOTAL] = df.loc[idx, TOTAL_RSV]
    df.loc[idx, AVG_FACTOR_PERIOD] = df.loc[idx, col_sales['avg_cut_sale']]
    df[AVG_FACTOR_PERIOD_WHS] = df[col_sales['avg_cut_sale']]
    df.drop(col_sales['last_sale'], axis=1, inplace=True)
    df.drop(col_sales['pntm_sale'], axis=1, inplace=True)
    df.drop(col_sales['avg_cut_sale'], axis=1, inplace=True)
    df.drop(SOFT_HARD_RSV, axis=1, inplace=True)
    df.drop(SOFT_HARD_RSV_CURRENT, axis=1, inplace=True)
    df.drop(TOTAL_RSV, axis=1, inplace=True)

    TRAD_HOLDINGS.remove(ALL_CLIENTS)
    idx = df[df[NAME_HOLDING].isin(TRAD_HOLDINGS)].index
    df.loc[idx, NAME_HOLDING] = NAME_TRAD
    return df


def fill_empty_cells(df):
    """Заполняет пустые ячейки нулями"""
    df_col = df.columns.tolist()
    numeric_col = df_col[df_col.index(PLAN_NFE):]
    for col in numeric_col:
        df = utils.void_to(df, col, 0)
    return df


def link_replace(df):
    """Заменяет сцепку по холдингу"""
    df.drop(LINK_HOLDING, axis=1, inplace=True)
    df.insert(
        1, LINK_HOLDING, df[WHS] + df[NAME_HOLDING] + df[EAN].map(str)
    )
    return df


def proccess_by_date(df, df_merge, cols_add, date_col):
    """Процесс добавления данных по датам"""
    df[INDEXES] = df.index
    df[LINK_UNIQUE] = (df[DATE_START].map(str)
                       + df[DATE_EXPIRATION].map(str)
                       + df[NAME_HOLDING]
                       + df[WHS])
    list_link = list(set(df[LINK_UNIQUE].tolist()))
    for col in cols_add:
        df[col] = 0
    df_pure = df.copy().drop(df.index, axis=0)

    for link in list_link:
        cols_add_copy = cols_add.copy()
        df_mid = df[df[LINK_UNIQUE] == link].copy()
        df_mid = df_mid.drop(cols_add_copy, axis=1)
        begin = df_mid[DATE_START].iloc[0]
        end = df_mid[DATE_EXPIRATION].iloc[0]
        whs = df_mid[WHS].iloc[0]
        holding = df_mid[NAME_HOLDING].iloc[0]
        df_merge_mid = df_merge.copy()

        if '), ' in str(holding):
            holding = get_mult_clients_dict(df_mid, NAME_HOLDING)[holding]
        elif holding == ALL_CLIENTS:
            df_merge_mid.loc[df_merge_mid.index, NAME_HOLDING] = holding
            holding = [holding]
        elif holding == NAME_TRAD:
            idx = df_merge_mid[df_merge_mid[PG_PROGRAMM] == CHANNEL_TRAD].index
            df_merge_mid.loc[idx, NAME_HOLDING] = holding
            holding = [holding]
        else:
            holding = [holding]

        if QUOTA in cols_add_copy:
            df_merge_mid_quota = df_merge_mid[
                (df_merge_mid[NAME_HOLDING].isin(holding))
                & (df_merge_mid[WHS] == whs)
            ].copy().groupby([EAN])[QUOTA].sum().reset_index()
            df_mid = df_mid.merge(df_merge_mid_quota, on=EAN, how='left')
            cols_add_copy.remove(QUOTA)

        df_merge_mid = df_merge_mid[
            (df_merge_mid[date_col] >= begin)
            & (df_merge_mid[date_col] <= end)
            & (df_merge_mid[NAME_HOLDING].isin(holding))
            & (df_merge_mid[WHS] == whs)
        ].copy().groupby([EAN])[cols_add_copy].sum().reset_index()
        df_mid = df_mid.merge(df_merge_mid, on=EAN, how='left')
        df_pure = pd.concat([df_pure, df_mid], ignore_index=True)
        df = df[df[LINK_UNIQUE] != link]

    df_pure = df_pure.sort_values(by=[INDEXES], ascending=[True])
    for col in cols_add:
        df_pure = utils.void_to(df_pure, col, 0)
    df_pure = df_pure.drop(columns=[LINK_UNIQUE, INDEXES], axis=1)
    df_pure.loc[df_pure[df_pure[FACTOR_PERIOD] == PAST].index, QUOTA] = 0
    return df_pure


def add_sales_rsv_by_date(df):
    """Добавляет данные по продажам и урезаниям учитывая даты заявок"""
    df_sales = get_data(TABLE_SALES_BY_DATE)
    df = proccess_by_date(
        df, df_sales, [SALES_BY_DATE, CUTS_BY_DATE], DATE_SALES
    )
    df_rsv = get_data(TABLE_RSV_BY_DATE)
    df_rsv = df_rsv[df_rsv[EXCLUDE_STRING] != WORD_YES]
    df = proccess_by_date(
        df, df_rsv, [SOFT_RSV, HARD_RSV, QUOTA], EXPECTED_DATE
    )
    df = df.rename(columns={
        SOFT_RSV: SOFT_RSV_BY_DATE,
        HARD_RSV: HARD_RSV_BY_DATE,
        QUOTA: QUOTA_BY_DATE
    })
    return df


def main():
    if os.environ.get('SRS_DOWNLOAD') is None:
        from update_data import update_factors_nfe, update_factors_pbi
        from update_data import update_factors_nfe_promo
        update_factors_nfe(SOURCE_FILE)
        update_factors_nfe_promo(SOURCE_FILE_PROMO)
        update_factors_pbi(SOURCE_FILE_PB)

    factors = add_deleted_factors(add_num_factors(filtered_factors()))
    factors = reindex_rename(split_by_month(add_pbi_and_purpose(factors)))
    factors = fill_empty_cells(add_total_sales_rsv(add_sales_and_rsv(factors)))
    factors = add_sales_rsv_by_date(link_replace(factors))
    save_to_excel(RESULT_DIR + TABLE_FACTORS, factors)
    print_complete(__file__)


if __name__ == "__main__":
    main()
