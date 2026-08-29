"""
date_helper.py

Date and Quarter Helper Functions
"""

import calendar
from datetime import datetime


# ==========================================================
# Quarter -> Month Numbers
# ==========================================================

QUARTER_MONTHS = {
    "Q1": [1, 2, 3],
    "Q2": [4, 5, 6],
    "Q3": [7, 8, 9],
    "Q4": [10, 11, 12],
}


# ==========================================================
# Month Titles
# Example:
# July'2026
# ==========================================================

def get_months(year: int, quarter: str) -> list[str]:

    return [

        datetime(year, month, 1).strftime("%B'%Y")

        for month in QUARTER_MONTHS[quarter]

    ]


# ==========================================================
# Weekday Dates
#
# weekday:
# Monday = 0
# Tuesday = 1
# ...
# Sunday = 6
#
# Returns exactly five values.
# Example:
# ["03", "10", "17", "24", "31"]
# ==========================================================

def get_weekday_dates(
    year: int,
    month: int,
    weekday: int
) -> list[str]:

    dates = [

        f"{week[weekday]:02d}"

        for week in calendar.monthcalendar(year, month)

        if week[weekday] != 0

    ]

    while len(dates) < 5:

        dates.append("N/A")

    return dates[:5]


# ==========================================================
# Count Weekdays
#
# Example:
# count_weekdays(2026, 4, 4)
# -> 4 Fridays
# ==========================================================

def count_weekdays(
    year: int,
    month: int,
    weekday: int
) -> int:

    return len(

        [

            week[weekday]

            for week in calendar.monthcalendar(year, month)

            if week[weekday] != 0

        ]

    )