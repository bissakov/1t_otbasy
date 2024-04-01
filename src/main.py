import logging
from os import listdir
from os.path import join
from time import sleep
from typing import Any, Dict

import requests
from selenium import webdriver
from selenium.common import ElementClickInterceptedException, TimeoutException
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.expected_conditions import element_to_be_clickable, presence_of_element_located
from selenium.webdriver.support.ui import WebDriverWait
from tqdm import tqdm
from webdriver_manager.chrome import ChromeDriverManager

from src.config import BASE_PATH, BRANCH_MAPPINGS
from src.logger import logger
from src.utils import get_app


WAIT_TIME = 10


def sign_auth(branch: str) -> None:
    password_folder = join(BASE_PATH, branch)
    logger.info(f'password_folder: {password_folder}')

    password = listdir(password_folder)[0]
    logger.info(f'password: {password}')
    auth_key_path = join(password_folder, password, listdir(join(password_folder, password))[0])

    choose_auth_app = get_app(title='Открыть файл')
    choose_auth_app.top_window().type_keys(f'{auth_key_path}~', with_spaces=True)

    password_app = get_app(title='Формирование ЭЦП в формате CMS')
    password_app.top_window().type_keys(f'{password}~')
    sleep(.5)
    password_app.top_window().type_keys('~')


def get_forms(driver: webdriver.Chrome) -> Dict[str, Any]:
    while len(driver.get_cookies()) == 0:
        sleep(.5)

    bin_number = driver.get_cookie('username')['value']
    session_key = driver.get_cookie('SESSION')['value']

    url = f'https://cabinet.stat.gov.kz/reports/getNewActiveReport/{bin_number}'

    querystring = {'lang': 'ru', 'state': 'all',
                   'page': '1', 'start': '0', 'limit': '20'}

    headers = {
        'accept': 'application/json',
        'cookie': f'SESSION={session_key}; username={bin_number}'
    }

    response = requests.get(url, headers=headers, params=querystring)

    return response.json()


def driver_init(driver_path: str | None) -> webdriver.Chrome:
    service = ChromeService(executable_path=driver_path)
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    return webdriver.Chrome(service=service,options=options)


def login(driver: webdriver.Chrome, branch: str) -> None:
    wait = WebDriverWait(driver, WAIT_TIME)
    driver.get('https://cabinet.stat.gov.kz/')

    wait.until(element_to_be_clickable((By.CSS_SELECTOR, 'a#idLogin'))).click()
    agree_container = wait.until(presence_of_element_located((By.CSS_SELECTOR, '#AgreeId')))
    agree_container.find_element(By.CLASS_NAME, 'x-btn-button').click()
    wait.until(element_to_be_clickable((By.CSS_SELECTOR, '#lawAlertCheck'))).click()
    wait.until(element_to_be_clickable((By.CSS_SELECTOR, '#loginButton'))).click()

    sign_auth(branch=branch)


def main() -> None:
    # {'17', '02', '14', '05', '08', '06', '13'}

    driver_path = ChromeDriverManager().install()

    branches = (branch_mapping['branch'] for branch_mapping in BRANCH_MAPPINGS)

    for branch in tqdm(iterable=branches, total=len(BRANCH_MAPPINGS), smoothing=0, desc='Проверка 1-Т'):
        driver = driver_init(driver_path)
        wait = WebDriverWait(driver, WAIT_TIME)
        with driver:
            login(driver=driver, branch=branch)

            try:
                wait.until(element_to_be_clickable((By.CSS_SELECTOR, '#tab-1168-btnInnerEl'))).click()
            except ElementClickInterceptedException:
                pass
            except TimeoutException:
                pass

            # wait.until(element_to_be_clickable((By.CSS_SELECTOR, 'a#tab-1170'))).click()
            #
            # table = wait.until(presence_of_element_located((By.CSS_SELECTOR, '#reportGridId-body')))

            # rows = driver.find_elements(By.CLASS_NAME, 'x-grid-cell-gridcolumn-1145')
            # texts = [row.text for row in rows if row.text]
            # index = texts.index("1-Т (квартальная)") if "1-Т (квартальная)" in texts else -1
            #
            # if index == -1:
            #     raise Exception('1-Т не найдена')
            #
            # rows[index].click()

            # wait.until(element_to_be_clickable((By.CSS_SELECTOR, '#ext-gen1978'))).click()
            #
            # wait.until(element_to_be_clickable((By.CSS_SELECTOR, '#createReportId-btnIconEl'))).click()
            #
            # window_handles = driver.window_handles
            #
            # driver.switch_to.window(window_handles[1])
            # wait.until(element_to_be_clickable((By.CSS_SELECTOR, '#btn-opendata'))).click()
            #
            # wait.until(element_to_be_clickable((By.CSS_SELECTOR, 'body > div:nth-child(16) > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button:nth-child(1)'))).click()
            # wait.until(element_to_be_clickable((By.CSS_SELECTOR, 'body > div:nth-child(18) > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button:nth-child(1)'))).click()

            form_list = [form['form']['name'] for form in get_forms(driver=driver)['list']]
            if '1-Т (квартальная)' in form_list:
                pass

    # Aa1234

if __name__ == '__main__':
    main()
