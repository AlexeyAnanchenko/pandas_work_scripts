"""
Формирует отчёт по максимальным потенциальным продажам текущего месяца

"""
import utils
utils.path_append()

from service import save_to_excel, get_data, print_complete
from settings import REPORT_POTENTIAL_SALES, REPORT_DIR, TABLE_FACTORS
from settings import FACTOR_PERIOD, FACTOR_STATUS, CURRENT, PURPOSE_PROMO
from settings import LINK, LINK_HOLDING, FACTOR, DATE_START, DATE_EXPIRATION
from settings import WHS, NAME_HOLDING, EAN, PRODUCT, LEVEL_3, DESCRIPTION
from settings import USER, PLAN_NFE, FACT_NFE, SALES_FACTOR_PERIOD
from settings import RSV_FACTOR_PERIOD, TABLE_SALES_HOLDINGS, TABLE_RESERVE
from settings import SOFT_HARD_RSV, NAME_TRAD, ALL_CLIENTS


ACTIVE_STATUS = [
    'Полностью согласован(а)',
    'Завершен(а)',
    'Частично согласован(а)'
]
INACTIVE_PURPOSE = 'Минимизация потерь'
FACTOR_SALES = 'Продажи'
ALIDI = 'Alidi Межфилиальные продажи (80000000)'


def get_factors():
    df = get_data(TABLE_FACTORS)
    df = df[
        (df[FACTOR_PERIOD] == CURRENT)
        & (df[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df[PURPOSE_PROMO] != INACTIVE_PURPOSE)
    ]
    df = df[[
        LINK, LINK_HOLDING, FACTOR, DATE_START, DATE_EXPIRATION,
        WHS, NAME_HOLDING, EAN, DESCRIPTION,
        USER, PLAN_NFE, FACT_NFE, SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD
    ]]
    return df


def merge_sales_rsv(df):
    sales, col = get_data(TABLE_SALES_HOLDINGS)
    rsv = get_data(TABLE_RESERVE)
    merge_col = [LINK, LINK_HOLDING, WHS, NAME_HOLDING, EAN]
    data_list = [
        [sales, col['last_sale'], SALES_FACTOR_PERIOD],
        [rsv, SOFT_HARD_RSV, RSV_FACTOR_PERIOD]
    ]

    for data in data_list:
        merge_df = data[0]
        replacement_col = data[1]
        replaceable_col = data[2]
        merge_col.append(replacement_col)
        df = df.rename(columns={replaceable_col: replacement_col})
        merge_df = merge_df[merge_df[replacement_col] != 0]
        df = df.merge(merge_df[merge_col], on=merge_col, how='outer')
        merge_col.pop()

    unnecessary_clients = [NAME_TRAD, ALL_CLIENTS, ALIDI]
    clients = list(set(df[NAME_HOLDING].to_list()))
    mult_clients = [i for i in clients if '), ' in str(i)]
    unnecessary_clients.extend(mult_clients)

    for client in unnecessary_clients:
        df = df[df[NAME_HOLDING] != client]

    return df


def main():
    df = merge_sales_rsv(get_factors())
    save_to_excel(REPORT_DIR + REPORT_POTENTIAL_SALES, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
