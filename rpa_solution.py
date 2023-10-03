import pandas as pd
from playwright.sync_api import sync_playwright


# Прочитать и сохранить в переменную df содержимое excel-файла.
df = pd.read_excel('challenge.xlsx')

# Запускаем браузер и переходим на страницу 'https://rpachallenge.com/'
with sync_playwright() as playwright:
    # Чтобы наглядно просмотреть работу браузера в launch можно передать аргумент headless=False
    browser = playwright.chromium.launch()
    page = browser.new_page()
    page.goto('https://rpachallenge.com/')

    # Ждем, пока страница загрузится, индикатор - появление кнопки "Start".
    page.wait_for_selector('//button[text()="Start"]')

    # Нажимаем кнопку "Start".
    page.click('//button[text()="Start"]')
    page.wait_for_selector('.row')

    # Циклом заполняем форму данными из excel-файла и нажимаем "Submit".
    for index, row in df.iterrows():
        page.fill('//input[@ng-reflect-name="labelFirstName"]', row['First Name'])
        page.fill('//input[@ng-reflect-name="labelFirstName"]', row['First Name'])
        page.fill('//input[@ng-reflect-name="labelLastName"]', row['Last Name '])
        page.fill('//input[@ng-reflect-name="labelCompanyName"]', row['Company Name'])
        page.fill('//input[@ng-reflect-name="labelRole"]', row['Role in Company'])
        page.fill('//input[@ng-reflect-name="labelAddress"]', row['Address'])
        page.fill('//input[@ng-reflect-name="labelEmail"]', row['Email'])
        page.fill('//input[@ng-reflect-name="labelPhone"]', str(row['Phone Number']))
        page.click('//input[@value="Submit"]')

    # Делаем скриншот результата работы.
    page.screenshot(path="screenshot.png")
    # Закрываем браузер.
    browser.close()
