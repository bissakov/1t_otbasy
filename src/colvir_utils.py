from time import sleep

import pywinauto
from pywinauto import Application

from src.config import CREDENTIALS, PROCESS_PATH
from src.utils import choose_mode, get_app


def login(app: Application | None = None) -> None:
    if not app:
        app = get_app(title='Вход в систему')
    login_win = app.window(title='Вход в систему')
    login_win['Edit2'].set_text(text=CREDENTIALS.user)
    login_win['Edit'].set_text(text=CREDENTIALS.password)
    login_win['OK'].send_keystrokes('~')


def confirm(app: Application | None = None) -> None:
    if not app:
        app = Application(backend='win32').connect(path=PROCESS_PATH)
    dialog = app.window(title='Colvir Banking System', found_index=0)
    timeout = 0
    while not dialog.window(best_match='OK').exists():
        if timeout >= 5.0:
            raise pywinauto.findwindows.ElementNotFoundError
        timeout += .1
        sleep(.1)
    dialog.send_keystrokes('~')


def check_interactivity(app: Application | None = None) -> None:
    if not app:
        app = Application(backend='win32').connect(path=PROCESS_PATH)
    choose_mode(app=app, mode='EXTRCT')
    if (filter_win := app.window(title='Фильтр')).exists():
        filter_win.close()
    else:
        raise pywinauto.findwindows.ElementNotFoundError


def open_colvir(max_tries: int = 10) -> Application:
    retry_count: int = 0
    app = None
    while retry_count < max_tries:
        try:
            app = Application().start(cmd_line=PROCESS_PATH)
            login(app=app)
            confirm(app=app)
            check_interactivity(app=app)
            break
        except pywinauto.findwindows.ElementNotFoundError:
            retry_count += 1
            if app:
                app.kill()
            continue
    if retry_count == max_tries:
        raise Exception('max_retries exceeded')
    return app
