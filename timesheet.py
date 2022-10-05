import sys
import argparse
import openpyxl

from collections import OrderedDict
from datetime import datetime, date, time
from warnings import filterwarnings


class Timesheet:

    def __init__(self, file_path: str, pretty_print=True):
        self._file_path = file_path
        self._workbook = openpyxl.load_workbook(file_path, read_only=False)
        self._sheet = self._workbook.active

        self._col_date = self._sheet["A"]
        self._col_start = self._sheet["B"]
        self._col_break_start = self._sheet["C"]
        self._col_break_end = self._sheet["D"]
        self._col_end = self._sheet["E"]

        self._pretty_print = pretty_print

    def enter(self, day: date = None,
              start: str = None, end: str = None,
              break_start: str = None, break_end: str = None) -> bool:
        day = day if day is not None else datetime.today().date()
        row_for_day = (row_index for row_index, cell in enumerate(self._col_date)
                       if isinstance(cell.value, datetime) and cell.value.date() == day)
        row_index = next(row_for_day, None)
        if row_index is None:
            print(f"Could not find row for day {day} in {self._file_path}", file=sys.stderr)
            return False
        if start is not None:
            self._set_start(row_index, start)
        if end is not None:
            self._set_end(row_index, end)
        if break_start is not None:
            self._set_break_start(row_index, break_start)
        if break_end is not None:
            self._set_break_end(row_index, break_end)
        self._print_row(row_index)
        return True

    def save(self, file_path=None):
        file_path = file_path if file_path is not None else self._file_path
        self._workbook.save(file_path)
        print(f"Timesheet updated: {file_path}", file=sys.stderr)

    def _print_row(self, row_index: int):
        row = OrderedDict({
            "date": self._get_day(row_index),
            "start": self._get_start(row_index),
            "break_start": self._get_break_start(row_index),
            "break_end": self._get_break_end(row_index),
            "end": self._get_end(row_index)
        })
        if self._pretty_print:
            from tabulate import tabulate
            print(tabulate([row], headers="keys", tablefmt="fancy_grid"))
        else:
            day = row["date"]
            times = ", ".join([f"{key}@{value}" for key, value in row.items() if key != "date"])
            print(f"{day} -> {times}")

    def _get_day(self, row_index: int) -> date:
        day = self._col_date[row_index].value
        if isinstance(day, datetime):
            return day.date()

    def _get_start(self, row_index: int):
        return self._col_start[row_index].value

    def _set_start(self, row_index: int, start: str):
        self._col_start[row_index].value = start

    def _get_end(self, row_index: int):
        return self._col_end[row_index].value

    def _set_end(self, row_index: int, end: str):
        self._col_end[row_index].value = end

    def _get_break_start(self, row_index: int):
        return self._col_break_start[row_index].value

    def _set_break_start(self, row_index: int, break_start: str):
        self._col_break_start[row_index].value = break_start

    def _get_break_end(self, row_index: int):
        return self._col_break_end[row_index].value

    def _set_break_end(self, row_index: int, break_end: str):
        self._col_break_end[row_index].value = break_end


def parse_date(day_str, date_format="%Y-%m-%d") -> date:
    try:
        return datetime.strptime(day_str, date_format).date()
    except ValueError:
        print(f"Invalid day: {args.day}", file=sys.stderr)
        sys.exit(-1)


def parse_time(time_str, time_format="%H:%M") -> time:
    try:
        return datetime.strptime(time_str, time_format).time()
    except ValueError:
        print(f"Invalid time: {time_str}", file=sys.stderr)
        sys.exit(-1)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("file",
                        help="Path to timesheet file (.xls, .xlsx)")
    parser.add_argument("-d", "--date", required=False,
                        help="The date for which entries should be added to the timesheet (yyyy-MM-dd)")
    parser.add_argument("-s", "--start", required=False,
                        help="The start time (hh:mm)")
    parser.add_argument("-e", "--end", required=False,
                        help="The end time (hh:mm)")
    parser.add_argument("-bs", "--break-start", required=False,
                        help="The start time of the break (hh:mm)")
    parser.add_argument("-be", "--break-end", required=False,
                        help="The end time of the break (hh:mm)")
    parser.add_argument("-p", "--pretty", required=False, action="store_true", default=False,
                        help="Pretty-print the updated row")
    args = parser.parse_args()

    if args.date:
        fmt = "%Y-%m-%d"
        args.date = parse_date(args.date, fmt)
    for time in [args.start, args.break_start, args.break_end, args.end]:
        if time:
            parse_time(time)

    filterwarnings('ignore', category=UserWarning, module='openpyxl')

    timesheet = Timesheet(args.file, args.pretty)
    if timesheet.enter(args.date, args.start, args.end, args.break_start, args.break_end):
        timesheet.save()
