import argparse
import xlwt
from typing import List, Any, Optional, Tuple
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


def _get_allowed_date_range(months_offset: int = 4) -> Tuple[dt.date, dt.date]:
    """
    Calculate the allowed date range based on current date.
    Returns (min_date, max_date) tuple.
    """
    today = dt.date.today()
    
    min_year = today.year
    min_month = today.month - months_offset
    while min_month < 1:
        min_month += 12
        min_year -= 1
    min_date = dt.date(min_year, min_month, 1)
    
    max_year = today.year
    max_month = today.month + months_offset
    while max_month > 12:
        max_month -= 12
        max_year += 1
    last_day = calendar.monthrange(max_year, max_month)[1]
    max_date = dt.date(max_year, max_month, last_day)
    
    return min_date, max_date


def _suggest_year_correction(year: int) -> Optional[int]:
    """
    Suggest a corrected year if the input appears to be a typo.
    Common typo: first digit wrong (e.g., 3025 -> 2025)
    """
    current_year = dt.date.today().year
    year_str = str(year)
    
    if len(year_str) == 4:
        for first_digit in ['1', '2', '3']:
            suggested = int(first_digit + year_str[1:])
            if 1900 <= suggested <= 2100 and abs(suggested - current_year) <= 10:
                return suggested
    
    return None


def _validate_date_range(year: int, month: int) -> None:
    """
    Validate that the requested year/month is within allowed range.
    Raises SystemExit with helpful message if validation fails.
    """
    min_date, max_date = _get_allowed_date_range()
    
    try:
        requested_date = dt.date(year, month, 1)
    except ValueError:
        raise SystemExit(f"Invalid date: year={year}, month={month}")
    
    if requested_date < min_date or requested_date > max_date:
        error_msg = (
            f"Date {year}/{month:02d} is outside the allowed range.\n"
            f"Allowed range: {min_date.strftime('%Y/%m')} to {max_date.strftime('%Y/%m')} "
            f"(±4 months from today)"
        )
        
        suggested_year = _suggest_year_correction(year)
        if suggested_year and suggested_year != year:
            error_msg += f"\n\nDid you mean {suggested_year} instead of {year}?"
        
        raise SystemExit(error_msg)


if __name__ == "__main__":
    args = _parse_args()
    year = args.year
    month = args.month
    if year < 1:
        raise SystemExit("Year must be positive")
    if not (1 <= month <= 12):
        raise SystemExit("Month must be between 1 and 12")
    
    _validate_date_range(year, month)

    filename = f"{year}{month:02d}.xls"

    rows: List[List[Any]] = generate_month_rows(year, month)

    generate_xls(rows, filename)
    print(f"Wrote {filename}")
