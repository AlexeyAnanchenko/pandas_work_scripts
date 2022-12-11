"""
Файл с настройками проекта
"""

from os.path import abspath, dirname


BASE_DIR = dirname(dirname(abspath(__file__)))
SOURCE_DIR = BASE_DIR + '/Исходники/'
RESULT_DIR = BASE_DIR + '/Результаты/'

NUM_MONTHS = 3
TARGET_STOCK = 1

EAN = 'EAN штуки'
PRODUCT = 'Наименование товара'
SU = 'SU штуки'
MSU = 'MSU штуки'
LEVEL_1 = '1-й уровень'
LEVEL_2 = '2-й уровень'
LEVEL_3 = '3-й уровень'
HOLDING = 'Код основного холдинга'
NAME_HOLDING = 'Наименование основного холдинга'
CODES = 'Код Точки доставки / Плательщика / Холдинга / Основного холдинга'
BASE_PRICE = 'Базовая цена без НДС, шт'
ELB_PRICE = 'Цена NIV Эльбрус с НДС, шт'
WHS = 'Склад'
LINK = 'Сцепка Склад-ШК'
LINK_HOLDING = 'Сцепка Склад-Наименование холдинга-ШК'
FULL_REST = 'Общий остаток, шт'
AVAILABLE_REST = 'Доступный остаток, шт'
QUOTA = 'Квота, шт'
FREE_REST = 'Свободный остаток, шт'
SOFT_RSV = 'Мягкий резерв, шт'
HARD_RSV = 'Жесткий резерв, шт'
SOFT_HARD_RSV = 'Мягкие + жёсткие резервы, шт'
QUOTA_BY_AVAILABLE = 'Резерв квота с учётом доступного стока'
TOTAL_RSV = 'В резерве всего'
DATE_RSV = 'Максимальная дата резерва'
CUTS = 'Урезания '
SALES = 'Продажи '
CUTS_SALES = 'Урезания + продажи '
AVARAGE = 'Средние '
OVERSTOCK = f'Сток свыше {TARGET_STOCK} месяца (-ев) продаж, шт'
MATRIX = 'Матрица Эльбрус'
MATRIX_LY = 'Матрица Эльбрус (Лоджистик-Юг)'

TABLE_ASSORTMENT = 'Ассортимент.xlsx'
TABLE_DIRECTORY = 'Справочник_ШК.xlsx'
TABLE_HOLDINGS = 'Холдинги.xlsx'
TABLE_PRICE = 'Прайс.xlsx'
TABLE_PURCHASES = 'Закупки.xlsx'
TABLE_REMAINS = 'Остатки.xlsx'
TABLE_SALES_HOLDINGS = 'Продажи по клиентам и складам.xlsx'
TABLE_SALES = 'Продажи по складам.xlsx'
TABLE_RESERVE = 'Резервы.xlsx'