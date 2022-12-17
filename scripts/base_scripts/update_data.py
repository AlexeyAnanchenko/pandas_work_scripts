import utils
utils.path_append()

import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

from hidden_settings import USER, PASSWORD
from settings import SOURCE_DIR

DOWNLOADS = "C:\\Users\\ananchenko.as\\Downloads\\"

options = Options()
options.add_argument('–disable-blink-features=BlockCredentialedSubresources')
options.add_argument('headless')
options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOADS
})
driver = webdriver.Chrome(options=options)

url_remains = (f'https://{USER}:{PASSWORD}@r.alidi.ru/ReportServer/Pages/'
               'ReportViewer.aspx?%2F%D0%9B%D0%BE%D0%B3%D0%B8%D1%81%D1%82'
               '%D0%B8%D0%BA%D0%B0%2F%D0%97%D0%B0%D0%BA%D1%83%D0%BF%D0%BA'
               '%D0%B8%2F1082%20-%20%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%'
               'D0%BD%D0%BE%D1%81%D1%82%D1%8C%20%D1%82%D0%BE%D0%B2%D0%B0%'
               'D1%80%D0%B0%20%D0%BF%D0%BE%20%D1%81%D0%BA%D0%BB%D0%B0%D0%'
               'B4%D0%B0%D0%BC%20(PG)&rc:showbackbutton=true')
url_reserve = (f'https://{USER}:{PASSWORD}@r.alidi.ru/ReportServer/Pages/'
               'ReportViewer.aspx?%2F%D0%9B%D0%BE%D0%B3%D0%B8%D1%81%D1%8'
               '2%D0%B8%D0%BA%D0%B0%2F%D0%97%D0%B0%D0%BA%D1%83%D0%BF%D0%'
               'BA%D0%B8%2F1275%20-%20%D0%A0%D0%B5%D0%B7%D0%B5%D1%80%D0%'
               'B2%D1%8B%20%D0%B8%20%D1%80%D0%B5%D0%B7%D0%B5%D1%80%D0%B2'
               '%D1%8B-%D0%BA%D0%B2%D0%BE%D1%82%D1%8B%20%D0%BF%D0%BE%20%'
               'D1%85%D0%BE%D0%BB%D0%B4%D0%B8%D0%BD%D0%B3%D0%B0%D0%BC&rc'
               ':showbackbutton=true')


def escort_download(driver, file):
    """
    Функция сопровождает выгрузку файла с сервера отчётов
    до папки с исходниками
    """
    export_id = 'ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink'
    export = driver.find_element(By.ID, export_id)
    export.click()
    driver.implicitly_wait(2)
    excel = driver.find_element(By.XPATH, "//a[@title='Excel']")
    excel.click()
    if file in os.listdir(DOWNLOADS):
        os.remove(SOURCE_DIR + file)
    flag = True

    while flag:
        if file not in os.listdir(DOWNLOADS):
            time.sleep(2)
        else:
            flag = False

    shutil.move(DOWNLOADS + file, SOURCE_DIR + file)


def update_remains(source_file):
    """Скачивает обновлённый отчёт по остаткам"""
    try:
        driver.get(url_remains)
        driver.implicitly_wait(20)
        submit_name = "ReportViewerControl$ctl04$ctl00"
        submit = driver.find_element(By.NAME, submit_name)
        submit.click()

        while 'GCAS' not in driver.page_source:
            time.sleep(3)

        escort_download(driver, source_file)
        print('Выгрузка по остаткам обновлена')
    finally:
        driver.quit()


def update_reserve(source_file):
    """Скачивает обновлённый отчёт по резервам"""
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
        view_report_name = 'ReportViewerControl$ctl04$ctl00'
        view_report = driver.find_element(By.NAME, view_report_name)
        view_report.click()

        while 'Мягкий резерв' not in driver.page_source:
            time.sleep(3)

        escort_download(driver, source_file)
        print('Выгрузка по резервам обновлена')
    finally:
        driver.quit()
