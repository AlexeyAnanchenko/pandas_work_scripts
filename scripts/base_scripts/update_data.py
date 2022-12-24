import utils
utils.path_append()

import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select

from settings import SOURCE_DIR, FACTOR_START
from hidden_settings import url_remains, url_reserve, url_price
from hidden_settings import url_factors_nfe, url_factors_pbi
from hidden_settings import USER_NFE, PASSWORD_NFE


options = Options()
options.add_argument('–disable-blink-features=BlockCredentialedSubresources')
options.add_argument('headless')
options.add_experimental_option("prefs", {
    "download.default_directory": SOURCE_DIR
})
driver = webdriver.Chrome(options=options)

PROVIDER = 'PROCTER&GAMBLE'
EXPORT_PBI = 'data.xlsx'


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


def update_factors_nfe():
    """Формирует обновлённый отчёт по факторам из NFE"""
    try:
        driver.get(url_factors_nfe)
        driver.implicitly_wait(10)
        time.sleep(1)
        driver.find_element(
            By.ID,
            'ctl00_ctl00_PageContent_Content_Login_FormLayout_Login_TextBox_I'
        ).send_keys(USER_NFE)
        driver.implicitly_wait(2)
        driver.find_element(
            By.ID,
            'ctl00_ctl00_PageContent_Content_Login_FormLayout_Password_TextB'
            'ox_I'
        ).send_keys(PASSWORD_NFE)
        driver.implicitly_wait(2)
        driver.find_element(
            By.ID,
            'ctl00_ctl00_PageContent_Content_Login_FormLayout_Login_Button_CD'
        ).click()
        driver.implicitly_wait(10)
        time.sleep(1)
        driver.find_element(
            By.ID,
            'ctl00_ctl00_PageContent_Content_Scorecards_PageControl_T1T'
        ).click()
        driver.implicitly_wait(2)
        driver.find_element(
            By.ID,
            'ctl00_ctl00_PageContent_Content_Scorecards_PageControl_SearchFilt'
            'er_FormLayout_DateFrom_DateEdit_I'
        ).send_keys(FACTOR_START)
        driver.implicitly_wait(2)
        driver.find_element(
            By.ID,
            'ctl00_ctl00_PageContent_Content_Scorecards_PageControl_SearchFil'
            'ter_FormLayout_Supplier_ComboBox_DDD_L_LBI4T0'
        ).click()
        driver.implicitly_wait(2)
        driver.find_element(
            By.ID,
            'ctl00_ctl00_PageContent_Content_Scorecards_PageControl_SearchFilt'
            'er_FormLayout_Search_Button_I'
        ).click()
        driver.implicitly_wait(2)
        time.sleep(3)
    finally:
        driver.quit()


def update_factors_pbi(source_file):
    """Формирует обновлённый отчёт по факторам из PBI отчёт для PG"""
    try:
        driver.get(url_factors_pbi)
        driver.implicitly_wait(20)
        driver.find_element(By.XPATH, "//div[@title='PG']").click()
        cell_name = 'Город'

        while cell_name not in driver.page_source:
            time.sleep(2)

        cell = driver.find_element(By.XPATH, f"//div[@title='{cell_name}']")
        cell.click()
        driver.implicitly_wait(3)
        driver.find_element(
            By.XPATH, "//button[@aria-label='Дополнительные параметры']"
        ).click()
        driver.implicitly_wait(3)
        driver.find_element(
            By.XPATH, "//div[@title='Экспортировать данные']"
        ).click()
        driver.implicitly_wait(3)
        export = driver.find_element(
            By.XPATH, "//button[@aria-label='Экспортировать']"
        )
        if source_file in os.listdir(SOURCE_DIR):
            os.remove(SOURCE_DIR + source_file)
        if EXPORT_PBI in os.listdir(SOURCE_DIR):
            os.remove(SOURCE_DIR + EXPORT_PBI)
        export.click()
        flag = True

        while flag:
            if EXPORT_PBI not in os.listdir(SOURCE_DIR):
                time.sleep(2)
            else:
                flag = False

        shutil.move(SOURCE_DIR + EXPORT_PBI, SOURCE_DIR + source_file)
        print('Выгрузка по факторам из PBI обновлена!')
    finally:
        driver.quit()
