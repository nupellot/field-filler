from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

# Загрузка веб-страницы
url = "https://www.rpachallenge.com/"  # Замените на URL нужной вам страницы
driver = webdriver.Chrome()  # Используйте свой WebDriver (например, Chrome, Firefox, Safari)
driver.get(url)

# Открываем файл Excel
workbook = openpyxl.load_workbook('challenge.xlsx')
sheet = workbook.active


for row in sheet.iter_rows(values_only=True):
	# Помимо данных текущей строки всегда забираем с собой заголовки таблицы.
	form_data = dict(zip([i.value for i in sheet[1]], row))
	# print("form_data: ", form_data)
	# print("form_data.items(): ", form_data.items())

	# Вводим все данные в нужные поля формы.
	for field_name, field_value in form_data.items():
		# Определяем label, к которому будем клеить значение.
		label_text = field_name
		# Используем XPath для нахождения input по тексту label
		xpath_expression = f"//label[text()='{label_text}']/following-sibling::input"
		input_element = driver.find_element(By.XPATH, xpath_expression)

		input_element.send_keys(str(field_value))

	# Отправить форму
	submit_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit']")
	submit_button.click()


# submit_button.click()

# Условие для контроля закрытия браузера
close_browser = input("Нажмите Enter, чтобы закрыть браузер...")
if close_browser:
    driver.quit()