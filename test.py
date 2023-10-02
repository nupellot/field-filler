import pandas as pandas
import requests as requests
from selenium import webdriver
from selenium.webdriver.common.by import By

TABLE_NAME = "challenge.xlsx"  # Имя, под которым будет сохраняться таблица с данными.
TABLE_DIRECTORY = ""
SCREENSHOT_DIRECTORY = ""
DEFAULT_DIRECTORY = ""

# Подготовительная работа. #

# Загрузка веб-страницы.
url = "https://www.rpachallenge.com/"  # URL страницы, с которой мы будем работать.
driver = webdriver.Chrome()  # Используем Selenium WebDriver (например, Chrome, Firefox, Safari).
driver.get(url)

try:
    # Забираем таблицу.
    print("Скачиваем таблицу с исходными данными...")
    response = requests.get("https://www.rpachallenge.com/assets/downloadFiles/challenge.xlsx")
    print("\033[42mСкачивание успешно завершено\033[0m")

    # Обработка необязательных параметров.
    file = open(TABLE_DIRECTORY + TABLE_NAME, "wb")
    file.write(response.content)  # Записываем данные в файл.
    file.close()
    sheet = pandas.read_excel(TABLE_DIRECTORY + TABLE_NAME)  # Читаем все данные из таблицы.

except requests.exceptions.RequestException:
    # В случае неудачи при скачивании таблицы используем нашу оффлайн-версию.
    print("Ошибка при получении файла с сервера: ", Exception)
    print("Используем заготовленную заранее копию.")
    sheet = pandas.read_excel(DEFAULT_DIRECTORY + TABLE_NAME)

# Нажимаем на кнопку "Старт".
xpath_expression = f"//button[text() = 'Start']"
start_element = driver.find_element(By.XPATH, xpath_expression)
start_element.click()

# Заполняем формы.
for row in sheet.iloc():  # Итерируемся по строкам.
    for field_name, field_value in row.items():  # Итерируемся по всем объектам в строке.
        # Определяем label, удаляя начальные и конечные пробелы.
        label_text = field_name.strip()
        # Используем XPath для нахождения input по тексту соответствующего label.
        xpath_expression = f"//label[text()='{label_text}']/following-sibling::input"
        # Находим поле, куда будем вводить данные.
        input_element = driver.find_element(By.XPATH, xpath_expression)
        # Заполняем определенное поле формы.
        input_element.send_keys(str(field_value))

    # Отправляем форму
    submit_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit']")
    submit_button.click()

# Сохраняем скриншот в случае успеха.
if driver.find_element(By.XPATH, "//div[text()[contains(.,'Congratulations!')]]"):
    driver.save_screenshot("Congratulations.png")

# Условие для контроля закрытия браузера
close_browser = input("Нажмите Enter, чтобы закрыть браузер...")
if close_browser:
    driver.quit()
