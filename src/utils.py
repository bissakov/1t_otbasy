from contextlib import contextmanager
from time import sleep
from typing import Optional

import psutil
import win32com.client as win32
from pywinauto import Application, ElementNotFoundError, WindowSpecification

from src.config import PROCESS_PATH


def kill_all_processes(proc_name: str) -> None:
    for proc in psutil.process_iter():
        if proc_name in proc.name():
            try:
                process = psutil.Process(proc.pid)
                process.terminate()
            except psutil.AccessDenied:
                continue


def get_app(title: str, backend: str = 'win32') -> Application:
    app = None
    while not app:
        try:
            app = Application(backend=backend).connect(title=title)
        except ElementNotFoundError:
            sleep(.1)
            continue
    return app


def choose_mode(mode: str, app: Application | None = None) -> None:
    if not app:
        app = Application(backend='win32').connect(path=PROCESS_PATH)
    mode_win = app.window(title='Выбор режима')
    mode_win['Edit2'].set_text(text=mode)
    mode_win['Edit2'].send_keystrokes('~')


def get_window(title: str, app: Optional[Application] = None, wait_for: str = 'exists', timeout: int = 20,
               regex: bool = False, found_index: int = 0) -> WindowSpecification:
    if not app:
        app = Application(backend='win32').connect(path=PROCESS_PATH)
    window = app.window(title=title, found_index=found_index) \
        if not regex else app.window(title_re=title, found_index=found_index)
    window.wait(wait_for=wait_for, timeout=timeout)
    sleep(.5)
    return window


@contextmanager
def dispatch(application: str) -> None:
    app = win32.Dispatch(application)
    app.DisplayAlerts = False
    try:
        yield app
    finally:
        kill_all_processes(proc_name='EXCEL')
