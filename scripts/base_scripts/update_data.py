import utils
utils.path_append()

import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select

from settings import SOURCE_DIR
from hidden_settings import url_remains, url_reserve, url_price


options = Options()
options.add_argument('–disable-blink-features=BlockCredentialedSubresources')
options.add_argument('headless')
options.add_experimental_option("prefs", {
    "download.default_directory": SOURCE_DIR
})
driver = webdriver.Chrome(options=options)


def view_report_click(driver, waiting_word):
    """Функция нажимает на кнопку 'Посмотреть отчёт' в SRS и ждёт отчёт"""
    view_report_name = 'ReportViewerControl$ctl04$ctl00'
    view_report = driver.find_element(By.NAME, view_report_name)
    view_report.click()
    while waiting_word not in driver.page_source:
        time.sleep(3)


def escort_download_srs(driver, file, folder=None):
    """Функция сопровождает выгрузку файла из SRS в исходники"""
    export_id = 'ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink'
    export = driver.find_element(By.ID, export_id)
    export.click()
    driver.implicitly_wait(2)
    excel = driver.find_element(By.XPATH, "//a[@title='Excel']")
    if file in os.listdir(SOURCE_DIR):
        os.remove(SOURCE_DIR + file)
    excel.click()
    flag = True
    while flag:
        if file not in os.listdir(SOURCE_DIR):
            time.sleep(2)
        else:
            flag = False
    if folder is not None:
        if file in os.listdir(folder):
            os.remove(folder + file)
        shutil.move(SOURCE_DIR + file, folder + file)


def update_remains(source_file):
    """Формирует обновлённый отчёт по остаткам"""
    try:
        driver.get(url_remains)
        driver.implicitly_wait(20)
        view_report_click(driver, 'GCAS')
        escort_download_srs(driver, source_file)
        print('Выгрузка по остаткам обновлена')
    finally:
        driver.quit()


def update_reserve(source_file):
    """Формирует обновлённый отчёт по резервам"""
    try:
        driver.get(url_reserve)
        driver.implicitly_wait(20)
        input_text_name = 'ReportViewerControl$ctl04$ctl05$txtValue'
        input_text = driver.find_element(By.NAME, input_text_name)
        input_text.send_keys('PG')
        button_date_id = 'ReportViewerControl_ctl04_ctl09_ctl01'
        button_date = driver.find_element(By.ID, button_date_id)
        button_date.click()
        driver.implicitly_wait(2)
        check_no_date_name = ('ReportViewerControl$ctl04$'
                              'ctl09$divDropDown$ctl02')
        check_no_date = driver.find_element(By.NAME, check_no_date_name)
        check_no_date.click()
        driver.implicitly_wait(2)
        check_with_date_name = ('ReportViewerControl$ctl04$'
                                'ctl09$divDropDown$ctl03')
        check_with_date = driver.find_element(By.NAME, check_with_date_name)
        check_with_date.click()
        driver.implicitly_wait(2)
        view_report_click(driver, 'Мягкий резерв')
        escort_download_srs(driver, source_file)
        print('Выгрузка по резервам обновлена')
    finally:
        driver.quit()


def update_price(source_file, folder):
    """Формирует обновлённый отчёт по базовым ценам"""
    try:
        driver.get(url_price)
        driver.implicitly_wait(20)
        select_pg = 'ReportViewerControl$ctl04$ctl03$ddValue'
        Select(driver.find_element(By.NAME, select_pg)).select_by_value('64')
        driver.implicitly_wait(2)
        select_whs = 'ReportViewerControl$ctl04$ctl05$ddValue'
        Select(driver.find_element(By.NAME, select_whs)).select_by_value('1')
        time.sleep(3)
        select_null = 'ReportViewerControl$ctl04$ctl17$ddValue'
        Select(driver.find_element(By.NAME, select_null)).select_by_value('1')
        driver.implicitly_wait(2)
        view_report_click(driver, 'GCAS')
        escort_download_srs(driver, source_file, folder)
        print('Выгрузка по базовым ценам обновлена')
    finally:
        driver.quit()
