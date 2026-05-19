"""주차 계산 유틸리티 — 매달 첫 수요일이 속한 주 = 해당 월 1주차 (주 시작: 월요일)"""

from datetime import datetime, timedelta


def get_week_info(date: datetime) -> tuple[str, str]:
    """
    Returns (월, 주차) strings, e.g. ('01월', '1주')
    Rule: the week containing the first Wednesday of each month is Week 1.
    """
    y, m = date.year, date.month
    first_day = datetime(y, m, 1)
    days_to_wed = (2 - first_day.weekday()) % 7
    first_wednesday = first_day + timedelta(days=days_to_wed)
    week1_start = first_wednesday - timedelta(days=first_wednesday.weekday())

    delta = (date.date() - week1_start.date()).days
    week_num = max(delta // 7 + 1, 1)
    return f'{m:02d}월', f'{week_num}주'


def week_date_range(year: int, month: int, week_num: int) -> tuple[datetime, datetime]:
    """해당 주차의 월요일~일요일 반환"""
    first_day = datetime(year, month, 1)
    days_to_wed = (2 - first_day.weekday()) % 7
    first_wednesday = first_day + timedelta(days=days_to_wed)
    week1_start = first_wednesday - timedelta(days=first_wednesday.weekday())
    monday = week1_start + timedelta(weeks=week_num - 1)
    sunday = monday + timedelta(days=6)
    return monday, sunday


def prev_week_range() -> tuple[datetime, datetime]:
    """오늘 기준 이전 주 (월~일) 반환"""
    today = datetime.today()
    this_monday = today - timedelta(days=today.weekday())
    last_monday = this_monday - timedelta(days=7)
    last_sunday = last_monday + timedelta(days=6)
    return last_monday, last_sunday
