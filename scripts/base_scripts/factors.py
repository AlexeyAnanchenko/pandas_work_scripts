"""
Формирует отчёт по факторам в удобном формате и с трекингом продаж

"""
import utils
utils.path_append()

import pandas as pd
import datetime as dt
from dateutil import relativedelta

from hidden_settings import WAREHOUSE_FACTORS
from service import save_to_excel, print_complete
from settings import SOURCE_DIR, RESULT_DIR, TABLE_FACTORS, WHS, FACTOR
from settings import FACTOR_NUM, REF_FACTOR, DATE_EXPIRATION, FACTOR_PERIOD
from settings import FACTOR_STATUS, DATE_CREATION, DATE_START, NAME_HOLDING
from settings import EAN, PRODUCT, LEVEL_3, DESCRIPTION, USER, PLAN_NFE
from settings import FACT_NFE, ADJUSTMENT_PBI, SALES_PBI, RESERVES_PBI
from settings import CUTS_PBI, LINK, LINK_HOLDING
# from update_data import update_factors_nfe, update_factors_pbi


SOURCE_FILE = 'NovoForecastServer_РезультатыПоиска.xlsx'
SOURCE_FILE_PB = 'Статистика факторов PG.xlsx'
WORKING_FACTORS = ['Акция', 'Предзаказ']
TOWN = 'Город'
TYPE_FACTOR = 'Тип'
ORDER_LOC = 'Сумма Заказ'
SALES_LOC = 'Сумма Продажи'
CUTS_LOC = 'Сумма Урезания'
RESERVES_LOC = 'Сумма Резервы'
HOLDING_LOC = 'Холдинг'
EAN_LOC = 'EAN'
PRODUCT_LOC = 'Продукт'
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


def filtered_factors(file=SOURCE_FILE, column=FACTOR, skip=0):
    """Получить отфильтрованный dataframe по нужным факторам"""
    df = pd.ExcelFile(SOURCE_DIR + file).parse(skiprows=skip)
    df = df[df[column].isin(WORKING_FACTORS)]
    df = df[df[TOWN].isin(WAREHOUSE_FACTORS)]
    df = df.rename(columns={TOWN: WHS})
    return df


def add_num_factors(df):
    """Функция добавляет номер фактора и преобразует ссылки"""
    prefix = 'https://forecast.alidi.ru/'
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


def add_pbi_col(df):
    """Добавляет столбцы из отчёта в PBI по факторам"""
    df_pb = filtered_factors(SOURCE_FILE_PB, TYPE_FACTOR, 2)
    df_pb.insert(
        0, LINK_FACTOR,
        (df_pb[FACTOR_NUM_PB].map(str) + df_pb[HOLDING_LOC]
         + df_pb[EAN_LOC].map(str))
    )
    df.insert(
        0, LINK_FACTOR,
        (df[FACTOR_NUM] + df[HOLDING_LOC] + df[EAN_LOC].map(str))
    )
    df = df.merge(
        df_pb[[LINK_FACTOR, ORDER_LOC, SALES_LOC, RESERVES_LOC, CUTS_LOC]],
        on=LINK_FACTOR, how='left'
    )
    df[LINK] = df[WHS] + df[EAN_LOC].map(str)
    df[LINK_HOLDING] = df[WHS] + df[HOLDING_LOC] + df[EAN_LOC].map(str)
    return df


def split_by_month(df):
    """Разделяет факторы на текущие, будущие и прошедшие в отдельном столбце"""
    today = dt.date.today()
    current_month = today.replace(day=1)
    current_month_dt = dt.datetime.combine(current_month, dt.time(0, 0))
    next_month = today + relativedelta.relativedelta(months=1, day=1)
    next_month_dt = dt.datetime.combine(next_month, dt.time(0, 0))
    df[FACTOR_PERIOD] = 'Будущий'
    idx = df[df[DATE_EXPIRATION_LOC] < next_month_dt].index
    df.loc[idx, FACTOR_PERIOD] = 'Текущий'
    idx = df[df[DATE_EXPIRATION_LOC] < current_month_dt].index
    df.loc[idx, FACTOR_PERIOD] = 'Прошедший'
    return df


def reindex_rename(df):
    result_df = df[[
        LINK, LINK_HOLDING, FACTOR_LOC, REF_FACTOR_LOC, FACTOR_NUM,
        FACTOR_STATUS_LOC, FACTOR_PERIOD, DATE_CREATION_LOC, DATE_START_LOC,
        DATE_EXPIRATION_LOC, WHS, HOLDING_LOC, EAN_LOC, PRODUCT_LOC,
        LEVEL_3_LOC, DESCRIPTION_LOC, USER_LOC, PLAN_LOC, FACT_LOC, ORDER_LOC,
        SALES_LOC, RESERVES_LOC, CUTS_LOC
    ]].rename(columns={
        FACTOR_LOC: FACTOR, REF_FACTOR_LOC: REF_FACTOR,
        FACTOR_STATUS_LOC: FACTOR_STATUS, DATE_CREATION_LOC: DATE_CREATION,
        DATE_START_LOC: DATE_START, DATE_EXPIRATION_LOC: DATE_EXPIRATION,
        HOLDING_LOC: NAME_HOLDING, EAN_LOC: EAN, PRODUCT_LOC: PRODUCT,
        LEVEL_3_LOC: LEVEL_3, DESCRIPTION_LOC: DESCRIPTION, USER_LOC: USER,
        PLAN_LOC: PLAN_NFE, FACT_LOC: FACT_NFE, ORDER_LOC: ADJUSTMENT_PBI,
        SALES_LOC: SALES_PBI, RESERVES_LOC: RESERVES_PBI, CUTS_LOC: CUTS_PBI
    })
    return result_df


def main():
    # update_factors_nfe(SOURCE_FILE)
    # update_factors_pbi(SOURCE_FILE_PB)
    factors = split_by_month(add_pbi_col(add_num_factors(filtered_factors())))
    save_to_excel(RESULT_DIR + TABLE_FACTORS, reindex_rename(factors))
    print_complete(__file__)


if __name__ == "__main__":
    main()
