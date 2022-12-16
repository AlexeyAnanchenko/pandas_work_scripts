import time
import keyboard

from selenium import webdriver
# from selenium.webdriver.common.by import By

from hidden_settings import USER, PASSWORD


driver = webdriver.Chrome()
url = ('https://r.alidi.ru/Reports/report/%D0%9B%D0%BE%D0%B3%D0%B8%D1%81'
       '%D1%82%D0%B8%D0%BA%D0%B0/%D0%97%D0%B0%D0%BA%D1%83%D0%BF%D0%BA%D0'
       '%B8/1082%20-%20%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%D0%BD%D0%BE%'
       'D1%81%D1%82%D1%8C%20%D1%82%D0%BE%D0%B2%D0%B0%D1%80%D0%B0%20%D0%B'
       'F%D0%BE%20%D1%81%D0%BA%D0%BB%D0%B0%D0%B4%D0%B0%D0%BC%20(PG)')

try:
    driver.get(url)
    time.sleep(2)
    keyboard.write(USER)
    keyboard.send('Tab')
    keyboard.write(PASSWORD)
    keyboard.send('enter')
    time.sleep(10)

    # # Метод find_element позволяет найти нужный элемент на сайте, указав путь к нему. Способы поиска элементов мы обсудим позже
    # # Метод принимает в качестве аргументов способ поиска и значение, по которому мы будем искать
    # # Ищем поле для ввода текста
    # textarea = driver.find_element(By.CSS_SELECTOR, ".textarea")

    # # Напишем текст ответа в найденное поле
    # textarea.send_keys("get()")
    # time.sleep(5)

    # # Найдем кнопку, которая отправляет введенное решение
    # submit_button = driver.find_element(By.CSS_SELECTOR, ".submit-submission")

    # # Скажем драйверу, что нужно нажать на кнопку. После этой команды мы должны увидеть сообщение о правильном ответе
    # submit_button.click()
    # time.sleep(5)
finally:
    driver.quit()
