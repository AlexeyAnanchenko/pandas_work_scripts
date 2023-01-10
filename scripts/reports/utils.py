"""
Скрипт содержит вспомогательные функции для base_scripts
"""

from sys import path
from os.path import dirname


def path_append():
    """Функция добавляет папку с настройками и сервисом в область видимости"""
    path.append(dirname(dirname(__file__)))
