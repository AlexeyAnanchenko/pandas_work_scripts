"""
Формирует отчёт по максимальным потенциальным продажам текущего месяца

"""
import utils
utils.path_append()

from service import save_to_excel, get_data, print_complete
from settings import REPORT_POTENTIAL_SALES, REPORT_DIR, TABLE_FACTORS
from settings import FACTOR_PERIOD, FACTOR_STATUS, TABLE_SALES_HOLDINGS
from settings import LINK_HOLDING, NAME_HOLDING, TABLE_RESERVE, WHS, EAN
from settings import TOTAL_RSV, LINK


def get_factors():
    df = get_data(TABLE_FACTORS)
    df = df.loc[(df[FACTOR_PERIOD] == 'Текущий')]
    df = df.loc[(df[FACTOR_STATUS] != 'Не согласован(а)')]
    return df


def add_col(df):
    sales, col_sales = get_data(TABLE_SALES_HOLDINGS)
    reserves = get_data(TABLE_RESERVE)
    static_col = [LINK, LINK_HOLDING, WHS, EAN, NAME_HOLDING]
    clients = list(set(df[NAME_HOLDING].to_list()))
    mult_clients = [i for i in clients if '), ' in str(i)]
    if mult_clients:
        sales.insert(
            0, LINK,
            sales[WHS] + sales[EAN].map(int).map(str)
        )
        sales = utils.replace_mult_clients(
            sales, mult_clients, static_col, [col_sales['last_sale']]
        )
        reserves = utils.replace_mult_clients(
            reserves, mult_clients, static_col, [TOTAL_RSV]
        )
    df = df.merge(
        sales[static_col + [col_sales['last_sale']]],
        on=static_col,
        how='outer'
    )
    # df = df.merge(
    #     reserves,
    #     on=LINK_HOLDING,
    #     how='outer'
    # )
    return df


def main():
    df = add_col(get_factors())
    save_to_excel(REPORT_DIR + REPORT_POTENTIAL_SALES, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
