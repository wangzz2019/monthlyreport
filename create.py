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


def _rand_datetime_str(d: dt.date, hour: int) -> str:
    minute = random.randint(0, 59)
    return f"{d.year}-{d.month}-{d.day} {hour:02d}:{minute:02d}"


def generate_month_rows(year: int, month: int) -> List[List[Any]]:
    rows: List[List[Any]] = []
    for d in _business_days_jp(year, month):
        start = _rand_datetime_str(d, 9)
        end = _rand_datetime_str(d, 18)
        rows.append([start, end, "", d.strftime("%Y-%m-%d"), "", ""])
    return rows


def _validate_date(year: int, month: int) -> None:
    """Validate year and month are within reasonable range (last 3 months to next 3 months)."""
    if year < 1:
        raise SystemExit("Year must be positive")
    if not (1 <= month <= 12):
        raise SystemExit("Month must be between 1 and 12")
    
    today = dt.date.today()
    current_year = today.year
    current_month = today.month
    
    try:
        input_date = dt.date(year, month, 1)
    except ValueError:
        raise SystemExit(f"Invalid date: {year}-{month:02d}")
    
    three_months_ago = today - dt.timedelta(days=90)
    three_months_later = today + dt.timedelta(days=90)
    
    if input_date < three_months_ago or input_date > three_months_later:
        suggestion = ""
        if year > current_year + 100:
            suggested_year = year - 1000
            if 1900 <= suggested_year <= current_year + 10:
                suggestion = f"\nDid you mean {suggested_year} instead of {year}?"
        
        error_msg = (
            f"Date {year}-{month:02d} is outside the valid range.\n"
            f"Please use dates between {three_months_ago.strftime('%Y-%m')} "
            f"and {three_months_later.strftime('%Y-%m')} "
            f"(last 3 months to next 3 months).{suggestion}"
        )
        raise SystemExit(error_msg)


if __name__ == "__main__":
    args = _parse_args()
    year = args.year
    month = args.month
    
    _validate_date(year, month)

    filename = f"{year}{month:02d}.xls"

    rows: List[List[Any]] = generate_month_rows(year, month)

    generate_xls(rows, filename)
    print(f"Wrote {filename}")
