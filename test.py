import sys
import time

import pandas as pandas
import requests as requests
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl


# Загрузка веб-страницы
url = "https://www.rpachallenge.com/"  # URL страницы, с которой мы будем работать.
driver = webdriver.Chrome()  # Используем Selenium WebDriver (например, Chrome, Firefox, Safari)
driver.get(url)


try:
	print("Скачиваем таблицу с исходными данными...")
	response = requests.get("https://www.rpachallenge.com/assets/downloadFiles/challenge.xlsx")
	print("\033[42mСкачивание успешно завершено\033[0m")
	if len(sys.argv) >= 2:
		file = open(str(sys.argv[1]) + "challenge.xlsx", "wb")
		file.write(response.content)
except requests.exceptions.RequestException:
	print("Ошибка при получении файла с сервера: ", Exception)
except Exception:
	print(Exception)


sheet = pandas.read_excel('challenge.xlsx')

# Нажимаем на кнопку "Старт".
xpath_expression = f"//button[text() = 'Start']"
start_element = driver.find_element(By.XPATH, xpath_expression)
start_element.click()

for row in sheet.iloc():
	for field_name, field_value in row.items():
		# print("key, value: ", key, value)
		# Определяем label, удаляя начальные и конечные пробелы.
		label_text = field_name.strip()
		# Используем XPath для нахождения input по тексту label
		xpath_expression = f"//label[text()='{label_text}']/following-sibling::input"
		input_element = driver.find_element(By.XPATH, xpath_expression)

		input_element.send_keys(str(field_value))

	# time.sleep(0.2)
	# Отправить форму
	submit_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit']")
	submit_button.click()


# Условие для контроля закрытия браузера
close_browser = input("Нажмите Enter, чтобы закрыть браузер...")
if close_browser:
    driver.quit()