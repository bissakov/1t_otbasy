import json
import os
from datetime import datetime
from os.path import abspath, dirname, join
from dataclasses import dataclass
import dotenv

from src import date_utils
from src.telegram_bot import TelegramBot


@dataclass
class Credentials:
    user: str
    password: str


dotenv.load_dotenv()

ROOT_FOLDER = dirname(dirname(__file__))

BASE_PATH: str = r'\\dbu157\c$\ЭЦП ключи'
with open(file=join(ROOT_FOLDER, 'branch_mappings.json'), mode='r', encoding='utf-8') as branch_mappings_file:
    BRANCH_MAPPINGS: list[dict[str, str]] = json.load(branch_mappings_file)

PROCESS_PATH = r'C:\CBS_R\COLVIR.EXE'
CREDENTIALS = Credentials(user=os.getenv('COLVIR_USR'), password=os.getenv('COLVIR_PSW'))
BOT = TelegramBot(token=os.getenv('TOKEN'), chat_id=os.getenv('CHAT_ID'))

PROJECT_FOLDER = dirname(abspath(''))

date_helper = date_utils.DateHelper(today=datetime.now())
REPORTS_FOLDER = join(PROJECT_FOLDER, 'reports', date_helper.get_prev_quarter_str_name())
JSON_FOLDER = join(PROJECT_FOLDER, 'json', date_helper.get_prev_quarter_str_name())
PREV_QUARTER_DATE_RANGES = date_helper.get_prev_quarter_date_ranges()
