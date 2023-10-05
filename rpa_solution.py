"""
A script for synchronously filling out forms on https://rpachallenge.com/
 using a playwright and a panda.
 """
import logging
import os

import pandas as pd
from playwright.sync_api import sync_playwright

# Logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
fh = logging.FileHandler('./log.txt')
fh.setLevel(logging.INFO)
fh.setFormatter(formatter)
logger.addHandler(fh)

# Constants
URL = 'https://rpachallenge.com/'
FIRST_NAME = '//input[@ng-reflect-name="labelFirstName"]'
LAST_NAME = '//input[@ng-reflect-name="labelLastName"]'
COMPANY_NAME = '//input[@ng-reflect-name="labelCompanyName"]'
ROLE_IN_COMPANY = '//input[@ng-reflect-name="labelRole"]'
ADDRESS = '//input[@ng-reflect-name="labelAddress"]'
EMAIL = '//input[@ng-reflect-name="labelEmail"]'
PHONE_NUMBER = '//input[@ng-reflect-name="labelPhone"]'
DOWNLOAD_EXCEL = "Download excel"
START = '//button[text()="Start"]'
SUBMIT = '//input[@value="Submit"]'


class RPASolver:
    """Основной класс взаимодействия со страницей."""

    def __init__(self, url, excel_path, screenshot_path):
        """ Инициализация объекта."""
        self.url = url
        self.excel_path_from_user = excel_path
        self.absolute_path_to_excel = ""
        self.screenshot_path = screenshot_path

    def download_excel_file(self, page) -> str:
        """Скачать Excel файл в определенное место
        через нажатие кнопки на странице."""
        logger.info("Downloading excel file")
        # Кликаем на кнопку скачать файл.
        with page.expect_download() as download_info:
            page.get_by_text(DOWNLOAD_EXCEL).click()
        download = download_info.value
        # К пути, указанному пользователем добавляем название файла с сайта по умолчанию.
        absolute_path_to_excel = self.excel_path_from_user + download.suggested_filename
        download.save_as(absolute_path_to_excel)
        logger.info(f"Excel file {download.suggested_filename} "
                    f"downloaded in {self.excel_path_from_user}")
        return absolute_path_to_excel

    def fill_input_field(self, page, selector, value):
        """Заполнить определенное поле по переданному индикатору."""
        # Ждем доступности поля и заполняем.
        page.wait_for_selector(selector)
        page.fill(selector, value)
        logger.info(f"Filling input field with selector: {selector} and value: {value}")

    def click_the_element(self, page, selector):
        """Кликнуть по элементу на странице, по переданному селектору."""
        page.wait_for_selector(selector)
        page.click(selector)

    def fill_form(self, page, row):
        """Заполнить одну форму, кликнуть 'Submit'."""
        self.fill_input_field(page, FIRST_NAME, row['First Name'])
        self.fill_input_field(page, LAST_NAME, row['Last Name '])
        self.fill_input_field(page, COMPANY_NAME, row['Company Name'])
        self.fill_input_field(page, ROLE_IN_COMPANY, row['Role in Company'])
        self.fill_input_field(page, ADDRESS, row['Address'])
        self.fill_input_field(page, EMAIL, row['Email'])
        self.fill_input_field(page, PHONE_NUMBER, str(row['Phone Number']))
        self.click_the_element(page, SUBMIT)
        logger.info(f"Filled form for row: {row}")

    def open_page_and_fill_the_forms(self):
        """Основной метод, запустить браузер, заполнить все формы."""
        with sync_playwright() as playwright:
            # Чтобы наглядно просмотреть работу браузера
            # в launch можно передать аргумент headless=False
            browser = playwright.chromium.launch()
            page = browser.new_page(accept_downloads=True)
            page.goto(self.url)
            logger.info(f"Opening page {self.url}")
            # Скачать Excel файл.
            self.absolute_path_to_excel = self.download_excel_file(page)
            # Начать заполнение форм
            self.click_the_element(page, START)
            df = pd.read_excel(self.absolute_path_to_excel)

            for index, row in df.iterrows():
                self.fill_form(page, row)
            page.screenshot(path=self.screenshot_path)
            logger.info("Screenshot saved")
            browser.close()


def validate_input(excel_path, screenshot_folder):
    """Проверка введенных пользователем значений. """
    if not os.path.isdir(excel_path):
        raise ValueError("Invalid Excel file path.")
    if not os.path.isdir(screenshot_folder):
        raise ValueError("Invalid screenshot folder path.")


def main():
    """Main function."""
    excel_path_from_user = input("Введите полный путь до папки, в которую"
                                 " следует сохранить файл Excel: ")
    screenshot_path_from_user = input("Введите полный путь до папки, "
                                      "в которую следует сохранить скриншот: ")
    screenshot_path = screenshot_path_from_user + "screenshot.png"

    try:
        validate_input(excel_path_from_user, screenshot_path_from_user)
        logger.info("Starting the RPA solver")
        solver = RPASolver(URL, excel_path_from_user, screenshot_path)
        solver.open_page_and_fill_the_forms()
        logger.info("Code execution completed successfully.")

    except Exception as e:
        logger.error(f"An error occurred: {str(e)}")


if __name__ == "__main__":
    main()
