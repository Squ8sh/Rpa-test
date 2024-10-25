import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def main():
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
        driver.get("https://www.rpachallenge.com/")
        logging.info("Переход на сайт выполнен.")

        try:
            start_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[text()='Start']"))
            )
            start_button.click()
            logging.info("Кнопка 'Start' нажата.")
        except Exception as e:
            logging.error(f"Ошибка при нажатии кнопки 'Start': {e}")
            return

        for page in range(10):
            logging.info(f"Заполнение страницы {page + 1} из 10.")
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//input")))

                field_mapping = {
                    'labelFirstName': 'First Name',
                    'labelLastName': 'Last Name',
                    'labelCompanyName': 'Company Name',
                    'labelRole': 'Role in Company',
                    'labelAddress': 'Address',
                    'labelEmail': 'Email',
                    'labelPhone': 'Phone Number'
                }
                for key, value in field_mapping.items():
                    input_field = driver.find_element(By.XPATH, f"//input[@ng-reflect-name='{key}']")
                    input_field.clear()  
                    input_field.send_keys(str(data[value][page]))  
                logging.info(f"Страница {page + 1} заполнена.")

                submit_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' or contains(@class, 'btn') and contains(@class, 'uiColorButton')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)  
                driver.execute_script("arguments[0].click();", submit_button) 
                logging.info(f"Кнопка 'Submit' на странице {page + 1} нажата.")

            except Exception as e:
                logging.error(f"Ошибка при заполнении страницы {page + 1}: {e}")
                return

    except Exception as e:
        logging.error(f"Общая ошибка в процессе выполнения: {e}")
    finally:
        close_browser = input("Закрыть браузер? (y/n): ")
        if close_browser.lower() == 'y':
            driver.quit()
            logging.info("Браузер закрыт.")
        else:
            logging.info("Браузер остался открыт. Закройте его вручную, когда закончите.")

if __name__ == "__main__":
    main()
