from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime, time, timedelta
from typing import Iterable

import pandas as pd


EMPTY_TOKENS = {"", "-", "—", "N/A", "NA", "NULL", "NONE", "NAN", "NAT"}
NON_BILLABLE_FLIGHT_TOKENS = (
    "LANDASAN",
    "RWYINS",
    "RWYINSP",
    "INSPCTN",
    "CAR",
    "VFR",
    "IFR",
    "TEST",
    "TEST1",
    "TEST2",
    "MPS",
)


def clean_text(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    while len(text) >= 2 and (text[0], text[-1]) in {("'", "'"), ('"', '"')}:
        text = text[1:-1].strip()
    return "" if text.upper() in EMPTY_TOKENS else text


def is_present(value: object) -> bool:
    return clean_text(value) != ""


def normalize_code(value: object) -> str:
    text = unicodedata.normalize("NFKC", clean_text(value)).upper()
    return re.sub(r"[^A-Z0-9]+", "", text)


def normalize_display_text(value: object) -> str:
    text = clean_text(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return text


def normalize_date(value: object) -> str:
    """Normalize Excel dates, YYMMDD, and common date strings to YYYY-MM-DD."""
    if not is_present(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        integer_text = str(int(number)) if number.is_integer() else ""
        if len(integer_text) == 6:
            try:
                return datetime.strptime(integer_text, "%y%m%d").strftime("%Y-%m-%d")
            except ValueError:
                pass
        if 2_400 <= number <= 100_000:
            try:
                return (datetime(1899, 12, 30) + timedelta(days=number)).strftime(
                    "%Y-%m-%d"
                )
            except (OverflowError, ValueError):
                pass

    text = clean_text(value)
    if re.fullmatch(r"\d{6}", text):
        try:
            return datetime.strptime(text, "%y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            pass
    if re.fullmatch(r"\d{8}", text):
        formats = ("%Y%m%d", "%d%m%Y") if text.startswith(("19", "20")) else ("%d%m%Y", "%Y%m%d")
        for date_format in formats:
            try:
                return datetime.strptime(text, date_format).strftime("%Y-%m-%d")
            except ValueError:
                continue
    iso_match = re.match(r"^(\d{4})-(\d{2})-(\d{2})", text)
    if iso_match:
        try:
            return datetime(
                int(iso_match.group(1)),
                int(iso_match.group(2)),
                int(iso_match.group(3)),
            ).strftime("%Y-%m-%d")
        except ValueError:
            pass
    try:
        parsed = pd.to_datetime(text, errors="raise", dayfirst=True)
        return parsed.strftime("%Y-%m-%d")
    except (TypeError, ValueError, OverflowError):
        return re.sub(r"\s+", " ", text).upper()


def _time_from_value(value: object) -> time | None:
    if not is_present(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, datetime):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, time):
        return value.replace(second=0, microsecond=0)
    if isinstance(value, timedelta):
        minutes = int(value.total_seconds() // 60) % (24 * 60)
        return time(minutes // 60, minutes % 60)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        if 0 <= number < 1:
            minutes = int(round(number * 24 * 60)) % (24 * 60)
            return time(minutes // 60, minutes % 60)
        if number.is_integer() and 0 <= number <= 2359:
            digits = f"{int(number):04d}"
            hour, minute = int(digits[:-2]), int(digits[-2:])
            if hour < 24 and minute < 60:
                return time(hour, minute)

    text = clean_text(value).upper().replace(".", ":")
    full_match = re.fullmatch(r"(\d{1,2}):(\d{2})(?::\d{2})?", text)
    if full_match:
        hour, minute = int(full_match.group(1)), int(full_match.group(2))
        if hour < 24 and minute < 60:
            return time(hour, minute)
    compact_match = re.fullmatch(r"\d{1,4}(?:\.0+)?", text)
    if compact_match:
        digits = text.split(".")[0].zfill(4)
        hour, minute = int(digits[:-2]), int(digits[-2:])
        if hour < 24 and minute < 60:
            return time(hour, minute)
    return None


def _has_explicit_date(value: object) -> bool:
    if isinstance(value, (pd.Timestamp, datetime, date)) and not isinstance(value, time):
        return True
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) >= 2_400
    text = clean_text(value)
    return bool(
        re.search(r"\b(?:19|20)\d{2}[-/]\d{1,2}[-/]\d{1,2}\b", text)
        or re.search(r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b", text)
        or re.match(r"^(?:19|20)\d{6}(?:\D|$)", text)
    )


def parse_datetime_value(
    value: object, flight_date: object
) -> tuple[pd.Timestamp | None, bool]:
    """Return a naive timestamp plus whether the source contained an explicit date."""
    if not is_present(value):
        return None, False
    explicit_date = _has_explicit_date(value)

    if isinstance(value, pd.Timestamp):
        return value.tz_localize(None) if value.tzinfo else value, True
    if isinstance(value, datetime):
        timestamp = pd.Timestamp(value)
        return timestamp.tz_localize(None) if timestamp.tzinfo else timestamp, True
    if isinstance(value, date) and not isinstance(value, datetime):
        return pd.Timestamp(value), True
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        if number >= 2_400:
            try:
                return pd.Timestamp(datetime(1899, 12, 30) + timedelta(days=number)), True
            except (OverflowError, ValueError):
                pass

    if explicit_date:
        text = clean_text(value)
        try:
            parsed = pd.to_datetime(text, errors="raise", dayfirst=False)
        except (TypeError, ValueError, OverflowError):
            try:
                parsed = pd.to_datetime(text, errors="raise", dayfirst=True)
            except (TypeError, ValueError, OverflowError):
                return None, True
        timestamp = pd.Timestamp(parsed)
        return timestamp.tz_localize(None) if timestamp.tzinfo else timestamp, True

    parsed_time = _time_from_value(value)
    normalized_flight_date = normalize_date(flight_date)
    if parsed_time is None or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", normalized_flight_date):
        return None, False
    base_date = datetime.strptime(normalized_flight_date, "%Y-%m-%d").date()
    return pd.Timestamp(datetime.combine(base_date, parsed_time)), False


def normalize_record_times(
    flight_date: object, atd_value: object, ata_value: object
) -> dict[str, object]:
    """Normalize ATD/ATA while preserving overnight arrival chronology."""
    atd_datetime, atd_explicit_date = parse_datetime_value(atd_value, flight_date)
    ata_datetime, ata_explicit_date = parse_datetime_value(ata_value, flight_date)

    if (
        atd_datetime is not None
        and ata_datetime is not None
        and not ata_explicit_date
        and ata_datetime < atd_datetime
    ):
        ata_datetime += pd.Timedelta(days=1)

    return {
        "ATD_DATETIME": atd_datetime,
        "ATA_DATETIME": ata_datetime,
        "ATD_TIME_DISPLAY": format_time(atd_datetime) or fallback_time_display(atd_value),
        "ATA_TIME_DISPLAY": format_time(ata_datetime) or fallback_time_display(ata_value),
        "ATD_EXPLICIT_DATE": atd_explicit_date,
        "ATA_EXPLICIT_DATE": ata_explicit_date,
    }


def normalize_eobt_datetime(
    flight_date: object, eobt_value: object
) -> pd.Timestamp | None:
    """Normalize EOBT to a naive datetime anchored to DATE OF FLIGHT."""
    eobt_datetime, _ = parse_datetime_value(eobt_value, flight_date)
    return eobt_datetime


def fallback_time_display(value: object) -> str:
    parsed_time = _time_from_value(value)
    if parsed_time:
        return parsed_time.strftime("%H:%M")
    return clean_text(value)


def format_time(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return pd.Timestamp(value).strftime("%H:%M")


def format_datetime(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return pd.Timestamp(value).strftime("%Y-%m-%d %H:%M")


def movement_datetime(
    movement: object,
    atd_datetime: pd.Timestamp | None,
    ata_datetime: pd.Timestamp | None,
) -> pd.Timestamp | None:
    move = normalize_code(movement)
    return ata_datetime if move == "A" else atd_datetime if move == "D" else None


def parse_priority_timestamp(value: object) -> pd.Timestamp | None:
    if not is_present(value):
        return None
    try:
        parsed = pd.to_datetime(clean_text(value), errors="raise", dayfirst=False)
    except (TypeError, ValueError, OverflowError):
        try:
            parsed = pd.to_datetime(clean_text(value), errors="raise", dayfirst=True)
        except (TypeError, ValueError, OverflowError):
            return None
    timestamp = pd.Timestamp(parsed)
    return timestamp.tz_localize(None) if timestamp.tzinfo else timestamp


def parse_message_number(value: object) -> float:
    text = clean_text(value)
    if not text:
        return float("-inf")
    try:
        return float(text)
    except ValueError:
        match = re.search(r"-?\d+(?:\.\d+)?", text)
        return float(match.group()) if match else float("-inf")


def is_excluded_flight_number(value: object) -> bool:
    """Return True for hard-excluded non-billable/internal movements."""
    flight = normalize_code(value)
    return any(token in flight for token in NON_BILLABLE_FLIGHT_TOKENS)


def is_non_billable_flight(value: object) -> bool:
    """Backward-compatible alias for the hard exclusion rule."""
    return is_excluded_flight_number(value)


def duplicate_group_key(values: Iterable[object]) -> str:
    return " | ".join(clean_text(value) for value in values)
