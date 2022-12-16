import time
import keyboard
from selenium import webdriver
from selenium.webdriver.common.by import By

from hidden_settings import USER, PASSWORD


driver = webdriver.Chrome()
url = ('https://r.alidi.ru/ReportServer/Pages/ReportViewer.aspx?%2F%D0%'
       '9B%D0%BE%D0%B3%D0%B8%D1%81%D1%82%D0%B8%D0%BA%D0%B0%2F%D0%97%D0%'
       'B0%D0%BA%D1%83%D0%BF%D0%BA%D0%B8%2F1082%20-%20%D0%94%D0%BE%D1%8'
       '1%D1%82%D1%83%D0%BF%D0%BD%D0%BE%D1%81%D1%82%D1%8C%20%D1%82%D0%B'
       'E%D0%B2%D0%B0%D1%80%D0%B0%20%D0%BF%D0%BE%20%D1%81%D0%BA%D0%BB%D'
       '0%B0%D0%B4%D0%B0%D0%BC%20(PG)&rc:showbackbutton=true')

try:
    driver.get(url)
    driver.implicitly_wait(10)
    keyboard.write(USER)
    keyboard.send('Tab')
    keyboard.write(PASSWORD)
    keyboard.send('enter')
    driver.implicitly_wait(20)
    submit = driver.find_element(By.NAME, "ReportViewerControl$ctl04$ctl00")
    submit.click()
    time.sleep(50)
    export = driver.find_element(
        By.ID,
        'ReportViewerControl_ctl05_ctl04_ctl00_ButtonLink'
    )
    export.click()
    driver.implicitly_wait(2)
    excel = driver.find_element(By.XPATH, "//a[@title='Excel']")
    excel.click()
    time.sleep(50)
finally:
    driver.quit()
