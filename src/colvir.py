import os
import shutil
from dataclasses import dataclass
from datetime import datetime, timedelta
from os.path import basename, dirname, exists, getmtime, join
from time import sleep
from typing import Optional

import openpyxl
import win32com.client as win32
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.timings import TimeoutError as TimingsTimeoutError
from tqdm import tqdm

from src import colvir_utils
from src import utils
from src.config import PREV_QUARTER_DATE_RANGES, REPORTS_FOLDER
from src.logger import logger


@dataclass
class Report:
    mode: str
    code: str
    branch: str
    file_path: str
    date_ranges: tuple[str, str]
    app: Optional[Application] = None


def is_correct_file(excel_full_file_path: str, excel: win32.Dispatch) -> bool:
    extension = excel_full_file_path.split('.')[-1]
    excel_full_file_path_no_ext = '.'.join(excel_full_file_path.split('.')[0:-1])
    excel_copy_path = f'{excel_full_file_path_no_ext}_copy.{extension}'
    shutil.copyfile(src=excel_full_file_path, dst=excel_copy_path)
    xlsx_file_path = f'{excel_full_file_path_no_ext}.xlsx'

    if not exists(path=xlsx_file_path):
        wb = excel.Workbooks.Open(excel_copy_path)
        wb.SaveAs(xlsx_file_path, FileFormat=51)
        wb.Close()

    workbook: Workbook = openpyxl.load_workbook(xlsx_file_path, data_only=True)
    sheet: Worksheet = workbook.active
    os.remove(excel_copy_path)
    os.remove(excel_full_file_path)

    return next((True for row in sheet.iter_rows(max_row=50) for cell in row if cell.alignment.horizontal), False)


def is_file_exported(full_file_name: str, excel: win32.CDispatch) -> bool:
    if not os.path.exists(path=full_file_name):
        return False
    if os.path.getsize(filename=full_file_name) == 0:
        return False
    try:
        os.rename(src=full_file_name, dst=full_file_name)
    except OSError:
        return False
    if not is_correct_file(excel_full_file_path=full_file_name, excel=excel):
        return False
    return True


def filter_reports_colvir(app: Application, report_code: str):
    report_win = utils.get_window(app=app, title='Выбор отчета')
    report_win.send_keystrokes('{F9}')

    report_filter_win = utils.get_window(app=app, title='Фильтр')
    report_filter_win['Edit4'].set_text(text=report_code)
    report_filter_win['OK'].send_keystrokes('~')


def fill_file_form(app: Application, report: Report) -> None:
    directory = dirname(report.file_path)
    file_name = basename(report.file_path)

    os.makedirs(directory, exist_ok=True)
    report_win = utils.get_window(app=app, title='Выбор отчета')
    file_win = app.window(title='Файл отчета ')
    while not file_win.exists():
        report_win['Экспорт в файл...'].send_keystrokes('~')
        sleep(.5)

    file_win['Edit4'].set_text(file_name)
    file_win['Edit2'].set_text(directory)
    file_win['ComboBox'].select(11)
    file_win['OK'].send_keystrokes('~')

    params_win = utils.get_window(title='Параметры отчета ')

    start_date, end_date = report.date_ranges

    if 'Z_160_PR_FORMOBWVEDZP' in file_name:
        params_win['Edit2'].set_text(report.branch)
        params_win['Edit4'].set_text(start_date)
        params_win['Edit6'].set_text(end_date)
        params_win['Edit8'].set_text('Ш')
    elif 'Z_160_SHTAT_RASTANOVKA_V' in file_name:
        params_win['Edit2'].set_text(report.branch)
        params_win['Edit4'].set_text(end_date)
    elif 'Z_160_HREMPLOYTAKETOWORK' in file_name:
        params_win['Edit2'].set_text(report.branch)
        params_win['Edit4'].set_text(start_date)
        params_win['Edit6'].set_text(end_date)
    elif 'Z_160_DISEMPLOYEE' in file_name:
        params_win['Edit2'].set_text(start_date)
        params_win['Edit4'].set_text(end_date)
        params_win['Edit6'].set_text(report.branch)
    params_win['OK'].send_keystrokes('~')


def export_report(report: Report) -> None:
    app = colvir_utils.open_colvir()
    utils.choose_mode(app=app, mode=report.mode)

    filter_win_title = 'PRS_GR4' if report.mode == 'PRS' else 'Фильтр'
    mode_filter_win = utils.get_window(app=app, title=filter_win_title)
    mode_filter_win['OK'].send_keystrokes('~')

    main_win_title = 'Персонал (зарплата)' if report.mode == 'CRD' else 'Персонал'
    main_win = utils.get_window(app=app, title=main_win_title)
    main_win.send_keystrokes('{F5}')

    report_win = utils.get_window(app=app, title='Выбор отчета')
    filter_reports_colvir(app=app, report_code=report.code)

    while (preview_checkbox := report_win['Предварительный просмотр']).is_checked():
        preview_checkbox.click()

    fill_file_form(app=app, report=report)
    report.app = app


def get_monitoring_reports(reports: list[Report]) -> None:
    for report in tqdm(reports, leave=False, smoothing=0, desc='Reports'):
        app = None
        while True:
            try:
                export_report(report=report)
            except (ElementNotFoundError, TimingsTimeoutError):
                if isinstance(app, Application):
                    app.kill()
                continue
            finally:
                break


def close_sessions(reports: list[Report]) -> None:
    pbar = tqdm(total=len(reports), leave=False, smoothing=0, desc='Closing sessions')
    with pbar, utils.dispatch(application='Excel.Application') as excel:
        while any(isinstance(r.app, Application) for r in reports):
            for i, report in enumerate(reports):
                if report.app is None:
                    continue

                if is_file_exported(full_file_name=report.file_path, excel=excel):
                    report.app.kill()
                    report.app = None
                    pbar.update(1)


def filter_reports(reports: list[Report]) -> list[Report]:
    filtered_reports = []
    for report in reports:
        file_path = report.file_path \
            if not report.file_path.endswith('.xls') \
            else report.file_path.replace('.xls', '.xlsx')
        if exists(file_path):
            current_time = datetime.now()
            file_creation_time = os.path.getctime(file_path)
            file_creation_datetime = datetime.fromtimestamp(file_creation_time)
            time_difference = current_time - file_creation_datetime
            twelve_hours = timedelta(hours=12)
            if time_difference > twelve_hours:
                filtered_reports.append(report)
            else:
                continue
        else:
            filtered_reports.append(report)
    return filtered_reports


def get_reports() -> list[Report]:
    reports_folder = REPORTS_FOLDER
    # branches = ['18', '19', '15', '08', '02',
    #             '20', '05', '07', '12', '13',
    #             '03', '09', '17', '04', '06',
    #             '14', '11', '01', '26', '21']

    branches = ['18']

    date_ranges = PREV_QUARTER_DATE_RANGES
    reports = []
    for branch in branches:
        code = 'Z_160_PR_FORMOBWVEDZP'
        file_path = join(reports_folder, f'{code}_{branch}.xls')
        report = Report(mode='CRD', code=code, branch=branch,
                        date_ranges=date_ranges, file_path=file_path)
        reports.append(report)

    reports.insert(1, Report(mode='PRS', code='Z_160_HREMPLOYTAKETOWORK',
                             branch='00', date_ranges=date_ranges,
                             file_path=join(reports_folder, 'Z_160_HREMPLOYTAKETOWORK_00.xls')))
    reports.insert(2, Report(mode='PRS', code='Z_160_SHTAT_RASTANOVKA_V',
                             branch='00', date_ranges=date_ranges,
                             file_path=join(reports_folder, 'Z_160_SHTAT_RASTANOVKA_V_00.xls')))
    reports.insert(3, Report(mode='PRS', code='Z_160_DISEMPLOYEE',
                             branch='00', date_ranges=date_ranges,
                             file_path=join(reports_folder, 'Z_160_DISEMPLOYEE_00.xls')))
    reports = filter_reports(reports=reports)
    return reports


def main():
    utils.kill_all_processes(proc_name='COLVIR')

    reports = get_reports()

    if not reports:
        logger.info('No reports to export')
        return

    get_monitoring_reports(reports=reports)
    close_sessions(reports=reports)

    # pass

    # Z_160_HREMPLOYTAKETOWORK
    # Z_160_DISEMPLOYEE
    # ['01', '02', '03', '04', '05', '06', '07', '08', '09', '11', '12', '13', '14', '15', '17', '18', '19', '20', '21', '26']


if __name__ == '__main__':
    import time
    start = time.time()
    main()
    end = time.time()
    print(f'Elapsed time: {end - start:.2f}')
