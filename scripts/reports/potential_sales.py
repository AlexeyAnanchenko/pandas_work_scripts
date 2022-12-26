"""
Формирует отчёт по максимальным потенциальным продажам текущего месяца

"""
import utils
utils.path_append()

from service import save_to_excel, get_data, print_complete, replace_symbol
from settings import REPORT_POTENTIAL_SALES, REPORT_DIR, TABLE_FACTORS
from settings import FACTOR_PERIOD, FACTOR_STATUS, TABLE_SALES_HOLDINGS
from settings import LINK_HOLDING, NAME_HOLDING, TABLE_RESERVE, WHS, EAN


def get_factors():
    df = get_data(TABLE_FACTORS)
    df = df.loc[(df[FACTOR_PERIOD] == 'Текущий')]
    df = df.loc[(df[FACTOR_STATUS] != 'Не согласован(а)')]
    return df


def add_col(df):
    sales, col_sales = get_data(TABLE_SALES_HOLDINGS)
    reserves = get_data(TABLE_RESERVE)

    for table in [df, sales, reserves]:
        replace_symbol(table, NAME_HOLDING, ', ИП', ' ИП')
    clients = list(set(df[NAME_HOLDING].to_list()))
    mult_clients = [i for i in clients if ', ' in str(i)]
    replace_holdings = {}
    for mult_client in mult_clients:
        for client in mult_client.split(', '):
            replace_holdings[client] = mult_client
    for table in [sales, reserves]:
        table = table.replace({NAME_HOLDING: replace_holdings})
        table[LINK_HOLDING] = (table[WHS]
                               + table[NAME_HOLDING]
                               + table[EAN].map(str))
    df[LINK_HOLDING] = df[WHS] + df[NAME_HOLDING] + table[EAN].map(str)
    df = df.merge(
        sales[[LINK_HOLDING, col_sales['last_sale']]],
        on=LINK_HOLDING,
        how='left'
    )
    return df


def main():
    df = add_col(get_factors())
    save_to_excel(REPORT_DIR + REPORT_POTENTIAL_SALES, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
