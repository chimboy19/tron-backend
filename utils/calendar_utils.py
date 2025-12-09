
# backend/utils/calendar_utils.py
from models import Calendar
from datetime import datetime, timedelta, date
import logging

log = logging.getLogger(__name__)

def load_calendar():
    try:
        cal_entries = Calendar.query.all()
    except Exception as e:
        log.warning(f"Failed to load Calendar entries: {e}", exc_info=True)
        return {}

    calendar_dict = {}
    for entry in cal_entries:
        try:
            d = entry.date
            if isinstance(d, datetime):
                d = d.date()
            date_key = d.strftime('%Y%m%d')
            calendar_dict[date_key] = {
                "is_holiday": bool(entry.is_holiday),
                "day_of_week": entry.day_of_week
            }
        except Exception as e:
            log.debug(f"Skipping malformed calendar entry {entry}: {e}")

    return calendar_dict


def is_business_day(date_obj, calendar_dict):
    """Check if a given date object is a business day based on the calendar dict.

    If a date isn't in the calendar dict we assume it's a business day by default.
    """
    if date_obj is None:
        return True

    if isinstance(date_obj, datetime):
        date_obj = date_obj.date()

    try:
        date_str = date_obj.strftime('%Y%m%d')
    except Exception:
        try:
            date_str = str(date_obj)
        except Exception:
            return True

    info = calendar_dict.get(date_str)
    if info:
        is_holiday = info.get("is_holiday")
        day_of_week = info.get("day_of_week")
        is_weekend = day_of_week in ["日", "土", "Sun", "Sat", "土曜", "日曜"]
        return not (is_holiday or is_weekend)
    else:
        log.debug(f"Date {date_str} not found in calendar, assuming business day.")
        return True


def add_business_days(start_date_obj, days, calendar_dict):
    """Add a number of business days to a start date."""
    if start_date_obj is None:
        start_date_obj = datetime.utcnow().date()

    if isinstance(start_date_obj, datetime):
        start_date_obj = start_date_obj.date()

    current_date = start_date_obj
    added = 0
    while added < max(0, int(days or 0)):
        current_date += timedelta(days=1)
        if is_business_day(current_date, calendar_dict):
            added += 1
    return current_date


def get_delivery_date(lead_time_days, calendar_dict):
    """Calculate delivery date from today + lead time (in business days).

    If lead_time_days <= 0, returns the next business day (not today).
    """
    today = datetime.utcnow().date()
    try:
        lt = int(lead_time_days or 0)
    except Exception:
        lt = 0

    if lt <= 0:
       
        current_date = today
        while True:
            current_date += timedelta(days=1)
            if is_business_day(current_date, calendar_dict):
                return current_date
    else:
        return add_business_days(today, lt, calendar_dict)
