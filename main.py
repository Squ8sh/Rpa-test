import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Загрузка данных из Excel файла
data = pd.read_excel('challenge.xlsx')

# Запуск браузера
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

try:
    # Переход на сайт
    driver.get("https://www.rpachallenge.com/")

    # Находим и нажимаем кнопку Start
    start_button = driver.find_element(By.XPATH, "//button[text()='Start']")
    start_button.click()

    # Ожидание загрузки форм
    WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//input")))

    # Заполнение форм
    for index in range(len(data)):  # Используем длину данных
        # Получаем данные для поля ввода
        row = data.iloc[index]  # Получаем строку из файла Excel
        inputs = driver.find_elements(By.XPATH, "//input")

        # Заполняем каждое поле
        for i, input_field in enumerate(inputs):
            input_field.clear()  # Стираем, если там что-то есть
            input_field.send_keys(str(row.iloc[i]))  # Заполнение данным из файла

        # Ожидание кнопки Submit и нажатие на нее через CSS-селектор
        submit_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit']"))
        )

        submit_button = driver.find_element(By.CLASS_NAME, 'btn.uiColorButton')
        submit_button.click()

        # Ожидание, чтобы форма была обработана и появилась следующая форма
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//input"))  # Ждем следующую форму
        )

    # Снимок экрана
    driver.save_screenshot('result.png')

finally:
    close_browser = input("Закрыть браузер? (y/n): ")
    if close_browser.lower() == 'y':
        driver.quit()
    else:
        print("Браузер остался открыт. Закройте его вручную, когда закончите.")
