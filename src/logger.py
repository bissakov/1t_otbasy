import datetime
import logging
import warnings

import urllib3
from os import makedirs
from os.path import join, dirname
from selenium.webdriver.remote.remote_connection import LOGGER


root_folder = join(dirname(dirname(__file__)), 'logs')
makedirs(root_folder, exist_ok=True)
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime).19s %(levelname)s %(name)s %(filename)s %(funcName)s : %(message)s')

today = datetime.date.today()
year_month_folder = join(root_folder, today.strftime('%Y/%B'))
makedirs(year_month_folder, exist_ok=True)

file_handler = logging.FileHandler(
    join(year_month_folder, f'{today.strftime("%d.%m.%y")}.log'),
    encoding='utf-8'
)
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

httpcore_logger = logging.getLogger('httpcore')
httpcore_logger.setLevel(logging.INFO)

urllib3.disable_warnings()

urllib3_logger = logging.getLogger('urllib3')
urllib3_logger.setLevel(logging.INFO)

LOGGER.setLevel(logging.WARNING)

warnings.simplefilter(action='ignore', category=UserWarning)
