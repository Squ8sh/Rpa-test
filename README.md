# Автоматизация заполнения форм с помощью Selenium и Pandas

Этот проект предназначен для автоматизации процесса заполнения форм на сайте [RPA Challenge](https://www.rpachallenge.com/) с использованием библиотеки Selenium и данных из Excel файла.

## Описание

Скрипт загружает данные из файла `challenge.xlsx`, открывает браузер, переходит на указанный сайт и автоматически заполняет формы на 10 страницах. Для каждого поля ввода данные берутся из Excel файла, а затем отправляются с помощью кнопки "Submit".

## Установка

1. Убедитесь, что у вас установлен Python (рекомендуется версия 3.6 и выше).
2. Установите необходимые библиотеки:

```bash
pip install pandas selenium webdriver-manager openpyxl
