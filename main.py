import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import logging

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Точка входа в программу
def main():
    # Загрузка данных из Excel файла
    try:
        data = pd.read_excel('challenge.xlsx')
        logging.info("Данные успешно загружены из Excel файла.")
    except FileNotFoundError as e:
        logging.error(f"Ошибка: Файл не найден - {e}")
        return
    except Exception as e:
        logging.error(f"Ошибка при загрузке данных из Excel файла: {e}")
        return

    # Получаем количество строк в DataFrame
    num_rows = len(data)
    logging.info(f"Количество строк в Excel: {num_rows}")

    # Запуск браузера
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        logging.info("Браузер успешно запущен.")
    except Exception as e:
        logging.error(f"Ошибка при запуске браузера: {e}")
        return

    try:
        # Переход на сайт
        driver.get("https://www.rpachallenge.com/")
        logging.info("Переход на сайт выполнен.")

        # Находим и нажимаем кнопку Start
        try:
            start_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[text()='Start']"))
            )
            start_button.click()
            logging.info("Кнопка 'Start' нажата.")
        except Exception as e:
            logging.error(f"Ошибка при нажатии кнопки 'Start': {e}")
            return

        # Переменная для отслеживания заполненных страниц
        filled_pages = 0

        while filled_pages < num_rows:  # Заполняем, пока есть строки в Excel
            logging.info(f"Заполнение страницы {filled_pages + 1} из {num_rows}.")
            try:
                # Ожидание загрузки полей ввода
                WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//input")))

                # Сопоставляем названия полей сайта с колонками Excel
                field_mapping = {
                    'labelFirstName': 'First Name',
                    'labelLastName': 'Last Name',
                    'labelCompanyName': 'Company Name',
                    'labelRole': 'Role in Company',
                    'labelAddress': 'Address',
                    'labelEmail': 'Email',
                    'labelPhone': 'Phone Number'
                }

                # Проходим по каждому полю формы
                for key, value in field_mapping.items():
                    input_field = driver.find_element(By.XPATH, f"//input[@ng-reflect-name='{key}']")
                    input_field.clear()  # Стираем, если там что-то есть
                    input_field.send_keys(str(data[value][filled_pages]))  # Заполняем данными из Excel
                logging.info(f"Страница {filled_pages + 1} заполнена.")

                # Нажимаем кнопку Submit
                submit_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' or contains(@class, 'btn') and contains(@class, 'uiColorButton')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)  # Прокрутка к кнопке
                driver.execute_script("arguments[0].click();", submit_button)  # Используем JavaScript для клика
                logging.info(f"Кнопка 'Submit' на странице {filled_pages + 1} нажата.")

                filled_pages += 1  # Увеличиваем счетчик заполненных страниц

            except Exception as e:
                logging.error(f"Ошибка при заполнении страницы {filled_pages + 1}: {e}")
                return

            
    except Exception as e:
        logging.error(f"Общая ошибка в процессе выполнения: {e}")
    finally:
        # Ожидание перед закрытием браузера
        close_browser = input("Закрыть браузер? (y/n): ")
        if close_browser.lower() == 'y':
            driver.quit()
            logging.info("Браузер закрыт.")
        else:
            logging.info("Браузер остался открыт. Закройте его вручную, когда закончите.")

if __name__ == "__main__":
    main()
