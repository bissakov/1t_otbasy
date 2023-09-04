import json
from os.path import dirname, join

ROOT_FOLDER = dirname(dirname(__file__))

BASE_PATH: str = r'\\dbu157\c$\ЭЦП ключи'
with open(file=join(ROOT_FOLDER, 'branch_mappings.json'), mode='r', encoding='utf-8') as branch_mappings_file:
    BRANCH_MAPPINGS: dict[str, str] = json.load(branch_mappings_file)