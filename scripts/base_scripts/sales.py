"""
Скрипт подготавливает файл с продажами в удобном формате

"""
import utils
utils.path_append()

from datetime import date, timedelta, datetime
from dateutil.relativedelta import relativedelta

from service import save_to_excel, print_complete, get_data
from settings import WHS, EAN, NAME_HOLDING, PRODUCT, NUM_MONTHS, DATE_SALES
from settings import LINK, LINK_HOLDING, CUTS, SALES, CUTS_SALES, AVARAGE
from settings import RESULT_DIR, TABLE_SALES_HOLDINGS, TABLE_SALES
from settings import TABLE_SALES_BY_DATE, SALES_BY_DATE, CUTS_BY_DATE


SOURCE_FILE = 'Продажи общие.xlsx'
EMPTY_ROWS = 13
REASON_FOR_CUTS = 3
WHS_LOC = 'Склад'
EAN_LOC = 'EAN'
M_HOLDING = 'Основной холдинг'
PRODUCT_NAME = 'Наименование товара'


def get_sales_by_month():
    """Получаем данные по месяцам"""
    df = get_data(TABLE_SALES_BY_DATE)
    last_sale = (date.today() - timedelta(days=1))
    merge_cols = [CUTS_BY_DATE, SALES_BY_DATE]
    date_for_loop = last_sale
    df.insert(
        0, LINK_HOLDING,
        df[WHS] + df[NAME_HOLDING] + df[EAN].map(int).map(str)
    )
    df_main = df[[
        LINK, LINK_HOLDING, WHS, NAME_HOLDING, EAN, PRODUCT
    ]].drop_duplicates()
    for col in merge_cols:
        year = date_for_loop.year
        month = date_for_loop.month
        first_day = datetime(year, month - 2, 1)
        last_day = datetime(year, month - 1, 1) - timedelta(days=1)
        month_name = first_day.strftime('%B')
        for _ in range(NUM_MONTHS):
            df_iter = df[
                (df[DATE_SALES] >= first_day) & (df[DATE_SALES] <= last_day)
            ].copy().groupby([LINK_HOLDING])[col].sum().reset_index()
            df_main = df_main.merge(df_iter, on=LINK_HOLDING, how='left')
            df_main = utils.void_to(df_main, col, 0)
            df_main = df_main.rename(
                columns={col: col + f' {month_name}' + f' {year}'}
            )
            year = first_day.year
            month = first_day.month
            first_day = datetime(year, month + 1, 1)
            last_day = datetime(year, month + 2, 1) - timedelta(days=1)
            month_name = first_day.strftime('%B')
    return df_main


def sales_by_client(df):
    """Формирование фрейма данных продаж в разрезе клиент-склад-шк"""
    col_month = list(df.columns)[-(NUM_MONTHS * (REASON_FOR_CUTS - 1)):]
    correct_month = col_month[:NUM_MONTHS]
    correct_month = [
        col.replace(CUTS_BY_DATE + ' ', '') for col in correct_month
    ]

    # группируем данные по EAN
    col_month_new = {}

    for i in list(df.columns)[-(NUM_MONTHS * (REASON_FOR_CUTS - 1)):]:
        col_month_new[i] = 'sum'

    # добавляем столбцы продажи + урезания
    cut_col = list(col_month_new.keys())[:NUM_MONTHS]
    sales_col = list(col_month_new.keys())[NUM_MONTHS:]
    cuts_sales_col = [CUTS_SALES + i for i in correct_month]

    for i in range(NUM_MONTHS):
        df[cuts_sales_col[i]] = df[[
            cut_col[i], sales_col[i]
        ]].sum(axis=1)

    # формируем столбцы с суммой за все месяца
    def get_sum_all_columns(reason_col, prefix):
        name_col = prefix + 'за {} месяца (-ев)'.format(NUM_MONTHS)
        df[name_col] = df[[col for col in reason_col]].sum(axis=1)
        return name_col

    sum_all_columns = [
        get_sum_all_columns(cut_col, CUTS),
        get_sum_all_columns(sales_col, SALES),
        get_sum_all_columns(cuts_sales_col, CUTS_SALES)
    ]

    # определяем делитель (кол-во дней) для расчёта среднего
    today = date.today() - timedelta(days=1)
    first_day_current_month = date(today.year, today.month, 1)
    delta = relativedelta(months=(3 - 1))
    divide_for_avarage = (
        today - (first_day_current_month - delta)
    ).days
    avarage_days_in_month = 365 / 12

    # создаём финальные столбцы и удаляем суммы за все месяца
    avarage_col = []

    for name_col in sum_all_columns:
        new_name = AVARAGE + name_col
        df[new_name] = (df[name_col]
                        / divide_for_avarage
                        * avarage_days_in_month).round(1)
        avarage_col.append(new_name)

    df.drop(sum_all_columns, axis=1, inplace=True)
    numeric_col = [
        cut_col + sales_col + cuts_sales_col + avarage_col
    ]
    return df, numeric_col


def sales_by_warehouses(dataframe, numeric_columns):
    """Формирование фрейма данных продаж в разрезе склад-шк"""

    sequence_col = [LINK, WHS, EAN, PRODUCT]
    numeric_col_dict = {}

    for col_list in numeric_columns:
        for col in col_list:
            numeric_col_dict[col] = 'sum'

    df = dataframe.groupby(sequence_col).agg(numeric_col_dict).reset_index()
    return df


def main():
    df, numeric_columns = sales_by_client(get_sales_by_month())
    save_to_excel(RESULT_DIR + TABLE_SALES_HOLDINGS, df)
    df = sales_by_warehouses(df, numeric_columns)
    save_to_excel(RESULT_DIR + TABLE_SALES, df)
    print_complete(__file__)


if __name__ == "__main__":
    main()
