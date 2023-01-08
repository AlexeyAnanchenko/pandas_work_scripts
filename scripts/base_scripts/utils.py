"""
Скрипт содержит вспомогательные функции для base_scripts
"""

from sys import path
from os.path import dirname
path.append(dirname(dirname(__file__)))

from settings import NAME_HOLDING, LINK_HOLDING, WHS, EAN


def path_append():
    """Функция добавляет папку с настройками и сервисом в область видимости"""

    path.append(dirname(dirname(__file__)))


def void_to(df, col, value):
    """Заменяет пустые значения в столбце на указанное значение"""
    idx = df[df[col].isnull()].index
    df.loc[idx, col] = value
    return df


def group_mult_clients(df, mult_clients, static_col, numeric_col):
    """
    Функция обновляет холдинги в dataframe по заданным мульти-холдингам
    и возвращает его сгруппированным
    """
    replace_holdings = {}
    for mult_client in mult_clients:
        mult_client_list = mult_client.split('), ')

        for client in mult_client_list:
            if client != mult_client_list[len(mult_client_list) - 1]:
                replace_holdings[client + ')'] = mult_client
            else:
                replace_holdings[client] = mult_client

    df = df.replace({NAME_HOLDING: replace_holdings})
    df = df.drop(columns=[LINK_HOLDING], axis=1)
    df.insert(0, LINK_HOLDING, df[WHS] + df[NAME_HOLDING] + df[EAN].map(str))
    numeric_col_dict = {}

    for col in numeric_col:
        numeric_col_dict[col] = 'sum'

    df = df.groupby(static_col).agg(numeric_col_dict).reset_index()
    return df
