import sys
import argparse
from typing import Tuple, Union

import openpyxl

from collections import OrderedDict
from datetime import datetime, date, time
from warnings import filterwarnings


class Timesheet:

    def __init__(self, file_path: str):
        self._file_path = file_path
        self._workbook = openpyxl.load_workbook(file_path, read_only=False)
        self._sheet = self._workbook.active

        # TODO make column names configurable
        self._col_date = "A"
        self._col_start = "B"
        self._col_break_start = "C"
        self._col_break_end = "D"
        self._col_end = "E"

        self._dirty = False

    def enter(self, day: date = None,
              start: time = None, end: time = None,
              break_start: time = None, break_end: time = None) -> Tuple[Union[int, None], bool]:
        row_index = self._get_row_index(day)
        if row_index is None:
            return None, False
        if start is not None:
            self._set_start(row_index, start)
        if end is not None:
            self._set_end(row_index, end)
        if break_start is not None:
            self._set_break_start(row_index, break_start)
        if break_end is not None:
            self._set_break_end(row_index, break_end)
        return row_index, self.is_dirty

    def save(self, file_path=None):
        file_path = file_path if file_path is not None else self._file_path
        self._workbook.save(file_path)
        print(f"Timesheet updated: {file_path}", file=sys.stderr)
        self._dirty = False

    def print_row_for_day(self, day: date):
        row_index = self._get_row_index(day)
        if row_index is None:
            return
        self.print_row(row_index)

    def print_row(self, row_index: int, pretty: bool = False):
        row_dict = OrderedDict({
            "date": self._get_day(row_index),
            "start": self._get_start(row_index),
            "break_start": self._get_break_start(row_index),
            "break_end": self._get_break_end(row_index),
            "end": self._get_end(row_index)
        })
        if pretty:
            from tabulate import tabulate
            print(tabulate([row_dict], headers="keys", tablefmt="fancy_grid"))
        else:
            day = row_dict["date"]
            times = ", ".join([f"{k}@{v}" for k, v in row_dict.items() if k != "date" and v is not None])
            print(f"{day} -> {times}")

    @property
    def is_dirty(self):
        return self._dirty

    def _get_row_index(self, day: date):
        row_for_day = (row_index for row_index, cell in enumerate(self._sheet[self._col_date])
                       if isinstance(cell.value, datetime) and cell.value.date() == day)
        row_index = next(row_for_day, None)
        if row_index is None:
            print(f"Could not find row for day {day} in {self._file_path}", file=sys.stderr)
        return row_index

    def _get_day(self, row_index: int) -> date:
        day = self._get_column(self._col_date, row_index).value
        if isinstance(day, datetime):
            return day.date()

    def _get_start(self, row_index: int):
        return self._get_column(self._col_start, row_index).value

    def _set_start(self, row_index: int, start: time):
        self._set_column(self._col_start, row_index, start)

    def _get_end(self, row_index: int):
        return self._get_column(self._col_end, row_index).value

    def _set_end(self, row_index: int, end: time):
        self._set_column(self._col_end, row_index, end)

    def _get_break_start(self, row_index: int):
        return self._get_column(self._col_break_start, row_index).value

    def _set_break_start(self, row_index: int, break_start: time):
        self._set_column(self._col_break_start, row_index, break_start)

    def _get_break_end(self, row_index: int):
        return self._get_column(self._col_break_end, row_index).value

    def _set_break_end(self, row_index: int, break_end: time):
        self._set_column(self._col_break_end, row_index, break_end)

    def _get_column(self, column: str, row_index: int):
        return self._sheet[column + str(row_index)]

    def _set_column(self, column: str, row_index: int, value):
        self._get_column(column, row_index).value = value
        self._dirty = True


def parse_date(x, date_format="%Y-%m-%d") -> date:
    if isinstance(x, date):
        return x
    try:
        return datetime.strptime(x, date_format).date()
    except ValueError:
        print(f"Invalid day: {x}", file=sys.stderr)
        sys.exit(-1)


def parse_time(x, time_format="%H:%M") -> time:
    if isinstance(x, time):
        return x
    try:
        return datetime.strptime(x, time_format).time()
    except ValueError:
        print(f"Invalid time: {x}", file=sys.stderr)
        sys.exit(-1)


if __name__ == "__main__":
    today = datetime.today().date()
    now = datetime.now().replace(second=0, microsecond=0).time()

    parser = argparse.ArgumentParser()
    parser.add_argument("file",
                        help="Path to timesheet file (.xls, .xlsx)")
    parser.add_argument("-d", "--date", required=False, nargs="?", const=today, default=today,
                        help="The date for which entries should be added to the timesheet (yyyy-MM-dd)")
    parser.add_argument("-s", "--start", required=False, nargs="?", const=now,
                        help="The start time (hh:mm)")
    parser.add_argument("-e", "--end", required=False, nargs="?", const=now,
                        help="The end time (hh:mm)")
    parser.add_argument("-bs", "--break-start", required=False, nargs="?", const=now,
                        help="The start time of the break (hh:mm)")
    parser.add_argument("-be", "--break-end", required=False, nargs="?", const=now,
                        help="The end time of the break (hh:mm)")
    parser.add_argument("-p", "--pretty", required=False, action="store_true", default=False,
                        help="Pretty-print the updated row")
    args = parser.parse_args()

    for attr in ["date", "start", "break_start", "break_end", "end"]:
        value = getattr(args, attr)
        if value:
            value = parse_date(value) if attr == "date" else parse_time(value)
            setattr(args, attr, value)

    filterwarnings('ignore', category=UserWarning, module='openpyxl')

    timesheet = Timesheet(args.file)
    row, updated = timesheet.enter(args.date, args.start, args.end, args.break_start, args.break_end)
    if row:
        timesheet.print_row(row, args.pretty)
    if updated:
        timesheet.save()
