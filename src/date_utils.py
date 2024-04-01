import calendar
from datetime import datetime
from typing import Tuple


class DateHelper:
    def __init__(self, today: datetime):
        self.today = today

    def get_prev_quarter_data(self) -> Tuple[int, int]:
        current_quarter = ((self.today.month - 1) // 3) + 1
        prev_quarter = current_quarter - 1 if current_quarter > 1 else 4
        year = self.today.year if prev_quarter != 4 else self.today.year - 1
        return prev_quarter, year

    def get_prev_quarter_str_name(self) -> str:
        quarter, year = self.get_prev_quarter_data()
        return f'{quarter} квартал {year}'

    @staticmethod
    def get_last_day(year: int, month: int) -> int:
        _, last_day = calendar.monthrange(year, month)
        return last_day

    def get_prev_quarter_date_ranges(self) -> tuple[str, str]:
        prev_quarter, _ = self.get_prev_quarter_data()

        prev_quarter_year = self.today.year if prev_quarter != 4 else self.today.year - 1
        end_month = prev_quarter * 3
        start_date = datetime(prev_quarter_year, end_month - 2, 1)
        end_date = datetime(prev_quarter_year, end_month, self.get_last_day(prev_quarter_year, end_month))

        start_date_str = start_date.strftime('%d.%m.%y')
        end_date_str = end_date.strftime('%d.%m.%y')

        return start_date_str, end_date_str
