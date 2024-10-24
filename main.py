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

        # Заполнение форм на 10 страницах
        for page in range(10):
            logging.info(f"Заполнение страницы {page + 1} из 10.")
            try:
                # Ожидание загрузки полей ввода
                WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//input")))
                inputs = driver.find_elements(By.XPATH, "//input[@ng-reflect-name]")

                # Заполняем каждое поле формы
                for i, input_field in enumerate(inputs):
                    input_field.clear()  # Стираем, если там что-то есть
                    input_field.send_keys(str(data.iloc[page, i]))  # Заполняем данными из Excel
                logging.info(f"Страница {page + 1} заполнена.")

                # Нажимаем кнопку Submit
                submit_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' or contains(@class, 'btn') and contains(@class, 'uiColorButton')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)  # Прокрутка к кнопке
                driver.execute_script("arguments[0].click();", submit_button)  # Используем JavaScript для клика
                logging.info(f"Кнопка 'Submit' на странице {page + 1} нажата.")

            except Exception as e:
                logging.error(f"Ошибка при заполнении страницы {page + 1}: {e}")
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
