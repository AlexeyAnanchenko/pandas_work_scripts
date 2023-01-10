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
from settings import RSV_FACTOR_PERIOD


ACTIVE_STATUS = [
    'Полностью согласован(а)',
    'Завершен(а)',
    'Частично согласован(а)'
]
INACTIVE_PURPOSE = 'Минимизация потерь'


def get_factors():
    df = get_data(TABLE_FACTORS)
    df = df[
        (df[FACTOR_PERIOD] == CURRENT)
        & (df[FACTOR_STATUS].isin(ACTIVE_STATUS))
        & (df[PURPOSE_PROMO] != INACTIVE_PURPOSE)
    ]
    df = df[[
        LINK, LINK_HOLDING, FACTOR, DATE_START, DATE_EXPIRATION,
        WHS, NAME_HOLDING, EAN, PRODUCT, LEVEL_3, DESCRIPTION,
        USER, PLAN_NFE, FACT_NFE, SALES_FACTOR_PERIOD, RSV_FACTOR_PERIOD
    ]]
    return df


def main():
    df = get_factors()
    save_to_excel(REPORT_DIR + REPORT_POTENTIAL_SALES, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
