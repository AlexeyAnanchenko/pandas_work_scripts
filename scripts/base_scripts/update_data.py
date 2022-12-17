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
from remains import SOURCE_FILE as SOURCE_FILE_REMAINS

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


def escort_download(dir, file):
    """Функция сопровождает выгрузку с сервера отчётов"""
    if file in os.listdir(DOWNLOADS):
        os.remove(dir + file)
    flag = True

    while flag:
        if file not in os.listdir(DOWNLOADS):
            time.sleep(2)
        else:
            flag = False

    shutil.move(DOWNLOADS + file, dir + file)


def update_remains():
    """Скачивает обновлённый отчёт по остаткам"""
    try:
        driver.get(url_remains)
        driver.implicitly_wait(20)
        submit_name = "ReportViewerControl$ctl04$ctl00"
        submit = driver.find_element(By.NAME, submit_name)
        submit.click()

        while 'GCAS' not in driver.page_source:
            time.sleep(3)

        export_id = 'ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink'
        export = driver.find_element(By.ID, export_id)
        export.click()
        driver.implicitly_wait(2)
        excel = driver.find_element(By.XPATH, "//a[@title='Excel']")
        excel.click()
        escort_download(SOURCE_DIR, SOURCE_FILE_REMAINS)
        print('Выгрузка по остаткам обновлена')
    finally:
        driver.quit()
