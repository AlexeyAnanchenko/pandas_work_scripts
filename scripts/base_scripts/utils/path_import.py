"""
Скрипт содержит функцию для добавления папки с настройками
и сервисом в область видимости
"""

from sys import path
from os.path import dirname


def path_append():
    path.append(dirname(dirname(dirname(__file__))))
