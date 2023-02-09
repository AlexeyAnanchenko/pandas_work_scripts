"""
Скрипт содержит вспомогательные функции для base_scripts
"""

from sys import path
from os.path import dirname


def path_append():
    """Функция добавляет папку с настройками и сервисом в область видимости"""
    path.append(dirname(dirname(__file__)))


def void_to(df, col, value):
    """Заменяет пустые значения в столбце на указанное значение"""
    idx = df[df[col].isnull()].index
    df.loc[idx, col] = value
    return df