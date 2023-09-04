import logging
from os import listdir
from os.path import join
from time import sleep

import requests
from pywinauto import Application, ElementNotFoundError
from selenium import webdriver
from selenium.common import ElementClickInterceptedException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

from logger import setup_logger
from config import BASE_PATH, BRANCH_MAPPINGS


def get_app(title: str, backend: str = 'win32'):
    app = None
    while not app:
        try:
            app = Application(backend=backend).connect(title=title)
        except ElementNotFoundError:
            sleep(.5)
            continue
    return app


def sign_auth(branch: str) -> None:
    password_folder = join(BASE_PATH, branch)
    logging.info(f'password_folder: {password_folder}')
    password = listdir(password_folder)[0]
    logging.info(f'password: {password}')
    auth_key_path = join(password_folder, password, listdir(join(password_folder, password))[0])

    choose_auth_app = get_app(title='Открыть файл')
    choose_auth_app.top_window().type_keys(f'{auth_key_path}~', with_spaces=True)
    password_app = get_app(title='Формирование ЭЦП в формате CMS')
    password_app.top_window().type_keys(f'{password}~~~~~~')


def get_forms(driver: webdriver.Chrome):
    while len(driver.get_cookies()) == 0:
        sleep(.5)

    bin_number = driver.get_cookie('username')['value']
    session_key = driver.get_cookie('SESSION')['value']

    url = f'https://cabinet.stat.gov.kz/reports/getNewActiveReport/{bin_number}'

    querystring = {'lang': 'ru', 'state': 'all', 'page': '1', 'start': '0', 'limit': '20'}

    headers = {
        'accept': 'application/json',
        'cookie': f'SESSION={session_key}; username={bin_number}'
    }

    response = requests.get(url, headers=headers, params=querystring)

    return response.json()


def driver_init(driver_path: str | None):
    service = Service(executable_path=ChromeDriverManager().install() if not driver_path else driver_path)
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    driver = webdriver.Chrome(
        service=service,
        options=options
    )
    return driver


def main():
    setup_logger()

    # {'17', '02', '14', '05', '08', '06', '13'}

    driver_path = None

    all_forms = {}
    for branch, branch_name in BRANCH_MAPPINGS.items():
        driver = driver_init(driver_path)
        if not isinstance(driver.service, str):
            driver_path = driver.service.path
        wait = WebDriverWait(driver, 10)
        with driver:
            driver.get('https://cabinet.stat.gov.kz/')
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a#idLogin'))).click()
            agree_container = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#AgreeId')))
            agree_container.find_element(By.CLASS_NAME, 'x-btn-button').click()
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#lawAlertCheck'))).click()
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#loginButton'))).click()

            sign_auth(branch=branch)

            try:
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#tab-1168-btnInnerEl'))).click()
                forms = get_forms(driver=driver)
                all_forms[branch] = forms
            except ElementClickInterceptedException:
                print(branch)

    print(all_forms)

    for branch, data in all_forms.items():
        print(branch, [form['form']['name'] for form in data['list']])
    pass


if __name__ == '__main__':
    main()
