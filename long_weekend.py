import datetime
import calendar
from typing import TypeVar

from holidays import country_holidays, list_supported_countries

HolidaySpec = TypeVar("HolidaySpec", tuple[int, int], str)
HolidayDictSpec = dict[str, HolidaySpec]

# Define U.S. federal holidays with fixed and dynamic dates
FEDERAL_HOLIDAYS: HolidayDictSpec = {
    "New Year's Day": (1, 1),
    "Martin Luther King Jr. Day": "third_monday_january",
    "Presidents' Day": "third_monday_february",
    "Memorial Day": "last_monday_may",
    "Independence Day": (7, 4),
    "Labor Day": "first_monday_september",
    "Columbus Day": "second_monday_october",
    "Veterans Day": (11, 11),
    "Thanksgiving Day": "fourth_thursday_november",
    "Christmas Day": (12, 25),
}

def calculate_holiday_date(year, holiday_rule):
    """Calculate the date of a holiday given its rule."""
    month_map = {
        "january": 1,
        "february": 2,
        "march": 3,
        "april": 4,
        "may": 5,
        "june": 6,
        "july": 7,
        "august": 8,
        "september": 9,
        "october": 10,
        "november": 11,
        "december": 12,
    }
    if isinstance(holiday_rule, tuple):
        # Fixed-date holiday
        month, day = holiday_rule
        return datetime.date(year, month, day)
    elif holiday_rule.startswith("first"):
        # First weekday in the month
        parts = holiday_rule.split("_")
        month = month_map[parts[2]]
        weekday = getattr(calendar, parts[1].upper())
        return next_weekday_in_month(year, month, weekday)
    elif holiday_rule.startswith("last"):
        # Last weekday in the month
        parts = holiday_rule.split("_")
        month = month_map[parts[2]]
        weekday = getattr(calendar, parts[1].upper())
        return last_weekday_in_month(year, month, weekday)
    elif holiday_rule.startswith("second") or holiday_rule.startswith("third") or holiday_rule.startswith("fourth"):
        # nth weekday in the month
        parts = holiday_rule.split("_")
        n = {"second": 2, "third": 3, "fourth": 4}[parts[0]]
        weekday = getattr(calendar, parts[1].upper())
        month = month_map[parts[2]]
        return nth_weekday_in_month(year, month, weekday, n)

def next_weekday_in_month(year, month, weekday):
    """Get the first occurrence of a weekday in a month."""
    for day in range(1, 8):
        if datetime.date(year, month, day).weekday() == weekday:
            return datetime.date(year, month, day)

def last_weekday_in_month(year, month, weekday):
    """Get the last occurrence of a weekday in a month."""
    last_day = calendar.monthrange(year, month)[1]
    for day in range(last_day, last_day - 7, -1):
        if datetime.date(year, month, day).weekday() == weekday:
            return datetime.date(year, month, day)

def nth_weekday_in_month(year, month, weekday, n):
    """Get the nth occurrence of a weekday in a month."""
    count = 0
    for day in range(1, calendar.monthrange(year, month)[1] + 1):
        if datetime.date(year, month, day).weekday() == weekday:
            count += 1
            if count == n:
                return datetime.date(year, month, day)

def suggest_long_weekends(year: int, holidays: HolidayDictSpec):
    """Suggest long weekends for the given year."""
    suggestions = []
    for holiday, rule in holidays.items():
        holiday_date = calculate_holiday_date(year, rule)
        holiday_weekday = holiday_date.weekday()

        if holiday_weekday == 0:  # Monday
            suggestion = f"{holiday} on {holiday_date} - Take the preceding Friday off for a 4-day weekend."
        elif holiday_weekday == 4:  # Friday
            suggestion = f"{holiday} on {holiday_date} - Take the following Monday for a 4-day weekend or the preceding Thursday and the following Monday for a 5-day weekend."
        elif holiday_weekday == 1:  # Tuesday
            suggestion = f"{holiday} on {holiday_date} - Take the preceding Monday off for a 4-day weekend."
        elif holiday_weekday == 3:  # Thursday
            suggestion = f"{holiday} on {holiday_date} - Take the following Friday off for a 4-day weekend."
        elif holiday_weekday == 2:  # Wednesday
            suggestion = f"{holiday} on {holiday_date} - Take the following Thursday and Friday off for a 5-day weekend."
        else:  # Weekend holidays
            suggestion = f"{holiday} on {holiday_date} falls on a weekend. Consider an alternative holiday."

        suggestions.append(suggestion)

    return suggestions

if __name__ == "__main__":
    country = ''
    country_dict = list_supported_countries(include_aliases=False)
    countries = [country for country in country_dict.keys()]
    while country.upper() not in countries:
        country = input(f"Enter they country (l for list): ").upper()
        if country == 'L':
            print(countries)

    subdivs = country_dict[country]
    if len(subdivs) > 0:
        print(f"Subdivisions for {country} are not supported yet: {subdivs}")

    year = int(input("Enter the year to plan long weekends: "))

    us_holidays = {name: (date.month, date.day) for date, name in sorted(country_holidays(country, years=year).items()) }
    s2 = suggest_long_weekends(year, us_holidays)
    print("\nLong Weekend Suggestions:")
    for suggestion in s2:
        print(f"- {suggestion}")

