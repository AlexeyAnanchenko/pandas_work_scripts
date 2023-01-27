"""
Скрипт содержит вспомогательные функции для base_scripts
"""

from sys import path
from os.path import dirname
import calendar


def path_append():
    """Функция добавляет папку с настройками и сервисом в область видимости"""
    path.append(dirname(dirname(__file__)))


def void_to(df, col, value):
    """Заменяет пустые значения в столбце на указанное значение"""
    idx = df[df[col].isnull()].index
    df.loc[idx, col] = value
    return df


def get_next_month(current_mon):
    current_month = current_mon
    current_month_num = list(calendar.month_name).index(current_month)
    next_month_num = current_month_num + 1 if current_month_num < 12 else 1
    next_month = calendar.month_name[next_month_num]
    return next_month
