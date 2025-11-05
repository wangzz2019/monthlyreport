import argparse
import xlwt
from typing import List, Any
import calendar
import datetime as dt
import random
import holidays


FIXED_HEADERS = [
    "Start Time",
    "End Time",
    "Remark",
    "日期",
    "OT Start 残業開始",
    "OT End 残業終了",
]


def generate_xls(rows: List[List[Any]], file_path: str, sheet_name: str = "Sheet1") -> None:
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheet_name)

    # Write fixed header row
    for col_idx, h in enumerate(FIXED_HEADERS):
        ws.write(0, col_idx, h)

    # Write data rows, pad/truncate to match header length
    for row_idx, r in enumerate(rows, start=1):
        values = (r + [""] * len(FIXED_HEADERS))[: len(FIXED_HEADERS)]
        for col_idx, val in enumerate(values):
            ws.write(row_idx, col_idx, val)

    wb.save(file_path)


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate monthly XLS with fixed headers.")
    parser.add_argument("year", type=int, help="Year in YYYY format, e.g. 2025")
    parser.add_argument("month", type=int, help="Month in 1-12")
    return parser.parse_args()


def _business_days_jp(year: int, month: int):
    jp_holidays = holidays.country_holidays("JP", years=[year])
    cal = calendar.Calendar()
    for d in cal.itermonthdates(year, month):
        if d.month != month:
            continue
        if d.weekday() >= 5:
            continue
        if d in jp_holidays:
            continue
        yield d


def _rand_time_hour(hour: int) -> str:
    return f"{hour:02d}:{random.randint(0,59):02d}"


def generate_month_rows(year: int, month: int) -> List[List[Any]]:
    rows: List[List[Any]] = []
    for d in _business_days_jp(year, month):
        start = _rand_time_hour(9)
        end = _rand_time_hour(18)
        rows.append([start, end, "", d.strftime("%Y-%m-%d"), "", ""])
    return rows


if __name__ == "__main__":
    args = _parse_args()
    year = args.year
    month = args.month
    if year < 1:
        raise SystemExit("Year must be positive")
    if not (1 <= month <= 12):
        raise SystemExit("Month must be between 1 and 12")

    filename = f"{year}{month:02d}.xls"

    rows: List[List[Any]] = generate_month_rows(year, month)

    generate_xls(rows, filename)
    print(f"Wrote {filename}")
