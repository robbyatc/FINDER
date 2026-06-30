from __future__ import annotations

from pathlib import Path
from typing import Iterable, Mapping

import pandas as pd

from comparison import (
    FIELD_KEYS,
    actual_file_format,
    actual_file_format_label,
    detect_mapping,
    excel_sheet_names,
    read_actual_excel,
    read_source_csv,
)
from utils.normalizer import (
    clean_text,
    duplicate_group_key,
    format_datetime,
    format_time,
    is_excluded_flight_number as _is_excluded_flight_number,
    is_present,
    movement_datetime,
    normalize_code,
    normalize_date,
    normalize_display_text,
    normalize_eobt_datetime,
    normalize_record_times,
    parse_message_number,
    parse_priority_timestamp,
)


RESULT_COLUMNS = [
    "DATE OF FLIGHT",
    "ACTUAL MOVEMENT DATE",
    "FLIGHT NUMBER",
    "AERODROME",
    "TO FROM",
    "AC REGISTER",
    "ATD",
    "ATA",
    "D/A/L/O",
    "ARRIVAL GATE",
    "DEPARTURE GATE",
    "DEPARTURE RUNWAY",
    "ARRIVAL RUNWAY",
]

VALIDATION_FOUND = "ADA DI STREAM"
VALIDATION_NOT_FOUND = "ADA DI DAT TIDAK ADA DI STREAM"
VALIDATION_REVIEW = "PERLU REVIEW STREAM"
VALIDATION_VALUES = (
    VALIDATION_FOUND,
    VALIDATION_NOT_FOUND,
    VALIDATION_REVIEW,
)
DISPLAY_RESULT_COLUMNS = RESULT_COLUMNS + ["VALIDASI"]

AUDIT_COLUMNS = [
    "SOURCE DATA",
    "SOURCE ROW",
    "STATUS",
    "MATCH_REASON",
    "STREAM STATUS FLIGHT",
    "DAT MOVEMENT DATETIME",
    "STREAM MOVEMENT DATETIME",
    "TIME DIFFERENCE MINUTES",
    "DUPLICATE GROUP KEY",
    "SELECTED DAT RECORD",
    "STREAM MATCH KEY",
    "STREAM VALIDATION RESULT",
    "STREAM MATCH DATE USED",
    "STREAM MATCH MODE",
    "BILLING CATEGORY",
    "DAT_RECOVERY_USED",
    "DAT_RECOVERY_REASON",
    "DAT_RECOVERY_SOURCE_DATE",
    "DAT_RECOVERY_SOURCE_ROW",
    "ORIGINAL_DAT_DATE",
    "RECOVERED_DAT_DATE",
    "USED_FOR_RECOVERY",
    "FLIGHT_INSTANCE_KEY",
    "MOVEMENT_TIME_BUCKET",
]

DETAIL_COLUMNS = DISPLAY_RESULT_COLUMNS + AUDIT_COLUMNS

DUPLICATE_AUDIT_COLUMNS = [
    "DUPLICATE_GROUP_KEY",
    "SELECTED_RECORD_FLAG",
    "DUPLICATE_REASON",
    "COMPLETENESS_SCORE",
    "HAS_MOVEMENT_TIME",
    "FLIGHT_INSTANCE_KEY",
    "MOVEMENT_TIME_BUCKET",
    "USED_FOR_RECOVERY",
    "DAT_RECOVERY_REASON",
    "DAT_RECOVERY_SOURCE_DATE",
    "DAT_RECOVERY_SOURCE_ROW",
]

RECOVERY_AUDIT_COLUMNS = [
    "DAT_RECOVERY_USED",
    "DAT_RECOVERY_REASON",
    "DAT_RECOVERY_SOURCE_DATE",
    "DAT_RECOVERY_SOURCE_ROW",
    "ORIGINAL_DAT_DATE",
    "RECOVERED_DAT_DATE",
    "USED_FOR_RECOVERY",
    "STREAM MATCH DATE USED",
    "STREAM MATCH MODE",
]

MATCH_KEY_SPECS = [
    ("DATE OF FLIGHT", "eobd"),
    ("FLIGHT NUMBER", "flight"),
    ("AERODROME", "adep"),
    ("TO FROM", "ades"),
    ("D/A/L/O", "movement"),
]

COMPLETENESS_FIELDS = [
    "AC REGISTER",
    "ATD",
    "ATA",
    "ARRIVAL GATE",
    "DEPARTURE GATE",
    "DEPARTURE RUNWAY",
    "ARRIVAL RUNWAY",
]

REQUIRED_DAT_FIELDS = ["flight", "eobd", "adep", "ades"]
REQUIRED_STREAM_FIELDS = ["flight", "eobd", "adep", "ades", "movement"]
DEFAULT_INVALID_STREAM_STATUSES = ("OTHER",)

STANDARDIZED_COLUMNS = RESULT_COLUMNS + [
    "SOURCE DATA",
    "SOURCE ROW",
    "STREAM STATUS FLIGHT",
    "TIMESTAMP",
    "MESSAGE NUM",
    "EOBT",
    "EOBT_DATETIME",
    "ATD_DATETIME",
    "ATA_DATETIME",
    "MOVEMENT_DATETIME",
    "MOVEMENT_TIME_DISPLAY",
    "MOVEMENT_DATE_DIFFERS_FROM_FLIGHT_DATE",
    "RAW ATD",
    "RAW ATA",
    "DAT_MOVEMENT_DATETIME",
    "DAT_MOVEMENT_TIME_DISPLAY",
    "STREAM_MOVEMENT_DATETIME",
    "STREAM_MOVEMENT_TIME_DISPLAY",
    "DAT_RECOVERY_USED",
    "DAT_RECOVERY_REASON",
    "DAT_RECOVERY_SOURCE_DATE",
    "DAT_RECOVERY_SOURCE_ROW",
    "ORIGINAL_DAT_DATE",
    "RECOVERED_DAT_DATE",
    "USED_FOR_RECOVERY",
]


def is_excluded_flight_number(flight_number: object) -> bool:
    """Hard filter for non-billable/internal callsigns."""
    return _is_excluded_flight_number(flight_number)


def read_uploaded_table(data: bytes, filename: str) -> tuple[pd.DataFrame, str]:
    """Read CSV, XLS, XLSX, or an HTML report carrying an .xls extension."""
    suffix = Path(filename).suffix.lower()
    physical_format = actual_file_format(data)

    if suffix in {".csv", ".tsv", ".txt"} or physical_format == "unknown":
        try:
            frame, encoding = read_source_csv(data)
            return frame, f"Delimited text · {encoding}"
        except Exception:
            if physical_format == "unknown":
                raise

    sheets = excel_sheet_names(data)
    frame = read_actual_excel(data, sheets[0])
    return frame, f"{actual_file_format_label(data)} · {sheets[0]}"


def detected_mapping(frame: pd.DataFrame) -> dict[str, str | None]:
    return detect_mapping(frame.columns)


def validate_required_columns(
    dep_mapping: Mapping[str, str | None],
    arr_mapping: Mapping[str, str | None],
    stream_mapping: Mapping[str, str | None],
) -> list[str]:
    issues: list[str] = []
    labels = {
        "flight": "Flight Number/Callsign",
        "eobd": "Date of Flight/EOBD",
        "adep": "ADEP/Aerodrome",
        "ades": "ADES/TO FROM",
        "movement": "D/A/L/O",
    }
    for dataset, mapping, required in (
        ("DAT DEP", dep_mapping, REQUIRED_DAT_FIELDS),
        ("DAT ARR", arr_mapping, REQUIRED_DAT_FIELDS),
        ("STREAM", stream_mapping, REQUIRED_STREAM_FIELDS),
    ):
        missing = [labels[field] for field in required if not mapping.get(field)]
        if missing:
            issues.append(f"{dataset}: {', '.join(missing)}")
    return issues


def _raw_value(
    row: pd.Series, mapping: Mapping[str, str | None], field: str
) -> object:
    column = mapping.get(field)
    return row[column] if column and column in row.index else ""


def _prefer(primary: object, fallback: object) -> str:
    return normalize_display_text(primary) if is_present(primary) else normalize_display_text(fallback)


def standardize_dataset(
    frame: pd.DataFrame,
    mapping: Mapping[str, str | None],
    source_name: str,
    default_movement: str | None = None,
) -> pd.DataFrame:
    records: list[dict[str, object]] = []
    for source_index, (_, row) in enumerate(frame.iterrows(), start=2):
        raw = {field: _raw_value(row, mapping, field) for field in FIELD_KEYS}
        date_of_flight = normalize_date(raw["eobd"])
        move = normalize_code(default_movement or raw["movement"])
        normalized_times = normalize_record_times(
            date_of_flight, raw["atd"], raw["ata"]
        )

        generic_gate = raw["parking"]
        generic_runway = raw["runway"]
        arrival_gate = raw["arrival_gate"]
        departure_gate = raw["departure_gate"]
        arrival_runway = raw["arrival_runway"]
        departure_runway = raw["departure_runway"]
        if move == "A":
            arrival_gate = _prefer(arrival_gate, generic_gate)
            arrival_runway = _prefer(arrival_runway, generic_runway)
        elif move == "D":
            departure_gate = _prefer(departure_gate, generic_gate)
            departure_runway = _prefer(departure_runway, generic_runway)

        atd_datetime = normalized_times["ATD_DATETIME"]
        ata_datetime = normalized_times["ATA_DATETIME"]
        eobt_datetime = normalize_eobt_datetime(date_of_flight, raw["eobt"])
        move_datetime = movement_datetime(move, atd_datetime, ata_datetime)
        flight_date_value = (
            pd.Timestamp(date_of_flight).date()
            if date_of_flight and pd.notna(pd.to_datetime(date_of_flight, errors="coerce"))
            else None
        )
        movement_date_differs = bool(
            move_datetime is not None
            and flight_date_value is not None
            and pd.Timestamp(move_datetime).date() != flight_date_value
        )

        record = {
            "DATE OF FLIGHT": date_of_flight,
            "ACTUAL MOVEMENT DATE": (
                pd.Timestamp(move_datetime).strftime("%Y-%m-%d")
                if move_datetime is not None
                else ""
            ),
            "FLIGHT NUMBER": normalize_code(raw["flight"]),
            "AERODROME": normalize_code(raw["adep"]),
            "TO FROM": normalize_code(raw["ades"]),
            "AC REGISTER": normalize_code(raw["register"]),
            "ATD": normalized_times["ATD_TIME_DISPLAY"],
            "ATA": normalized_times["ATA_TIME_DISPLAY"],
            "D/A/L/O": move,
            "ARRIVAL GATE": normalize_display_text(arrival_gate),
            "DEPARTURE GATE": normalize_display_text(departure_gate),
            "DEPARTURE RUNWAY": normalize_code(departure_runway),
            "ARRIVAL RUNWAY": normalize_code(arrival_runway),
            "SOURCE DATA": source_name,
            "SOURCE ROW": source_index,
            "STREAM STATUS FLIGHT": normalize_code(raw["status_flight"]),
            "TIMESTAMP": normalize_display_text(raw["timestamp"]),
            "MESSAGE NUM": normalize_display_text(raw["message_num"]),
            "EOBT": format_time(eobt_datetime) or normalize_display_text(raw["eobt"]),
            "EOBT_DATETIME": eobt_datetime,
            "ATD_DATETIME": atd_datetime,
            "ATA_DATETIME": ata_datetime,
            "MOVEMENT_DATETIME": move_datetime,
            "MOVEMENT_TIME_DISPLAY": normalized_times[
                "ATA_TIME_DISPLAY" if move == "A" else "ATD_TIME_DISPLAY"
            ],
            "MOVEMENT_DATE_DIFFERS_FROM_FLIGHT_DATE": movement_date_differs,
            "RAW ATD": clean_text(raw["atd"]),
            "RAW ATA": clean_text(raw["ata"]),
            "DAT_RECOVERY_USED": False,
            "DAT_RECOVERY_REASON": "",
            "DAT_RECOVERY_SOURCE_DATE": "",
            "DAT_RECOVERY_SOURCE_ROW": "",
            "ORIGINAL_DAT_DATE": "",
            "RECOVERED_DAT_DATE": "",
            "USED_FOR_RECOVERY": False,
        }
        if source_name == "STREAM":
            record["STREAM_MOVEMENT_DATETIME"] = move_datetime
            record["STREAM_MOVEMENT_TIME_DISPLAY"] = record["MOVEMENT_TIME_DISPLAY"]
            record["DAT_MOVEMENT_DATETIME"] = None
            record["DAT_MOVEMENT_TIME_DISPLAY"] = ""
        else:
            record["DAT_MOVEMENT_DATETIME"] = move_datetime
            record["DAT_MOVEMENT_TIME_DISPLAY"] = record["MOVEMENT_TIME_DISPLAY"]
            record["STREAM_MOVEMENT_DATETIME"] = None
            record["STREAM_MOVEMENT_TIME_DISPLAY"] = ""
        records.append(record)
    return pd.DataFrame(records, columns=STANDARDIZED_COLUMNS)


def recover_adjacent_date_movements(dat_normalized: pd.DataFrame) -> pd.DataFrame:
    """Recover missing DAT movement details from the next raw DAT date.

    Recovery intentionally runs on the normalized, non-deduplicated DAT pool. A
    source record can be consumed only once and is later retained as an audit
    duplicate instead of being reconciled a second time.
    """
    working = dat_normalized.copy().reset_index(drop=True)
    if working.empty:
        return working

    for column, default in (
        ("DAT_RECOVERY_USED", False),
        ("DAT_RECOVERY_REASON", ""),
        ("DAT_RECOVERY_SOURCE_DATE", ""),
        ("DAT_RECOVERY_SOURCE_ROW", ""),
        ("ORIGINAL_DAT_DATE", ""),
        ("RECOVERED_DAT_DATE", ""),
        ("USED_FOR_RECOVERY", False),
    ):
        if column not in working.columns:
            working[column] = default

    parsed_dates = pd.to_datetime(working["DATE OF FLIGHT"], errors="coerce")
    order = working.assign(__date_order=parsed_dates).sort_values(
        ["__date_order", "SOURCE ROW"], kind="stable"
    ).index
    identity_columns = ["FLIGHT NUMBER", "AERODROME", "TO FROM", "D/A/L/O"]
    recover_display_fields = [
        "AC REGISTER",
        "ATD",
        "ATA",
        "ARRIVAL GATE",
        "DEPARTURE GATE",
        "DEPARTURE RUNWAY",
        "ARRIVAL RUNWAY",
    ]
    recover_internal_fields = [
        "ATD_DATETIME",
        "ATA_DATETIME",
        "RAW ATD",
        "RAW ATA",
    ]

    for index in order:
        row = working.loc[index]
        movement = normalize_code(row["D/A/L/O"])
        if movement not in {"A", "D"}:
            continue
        movement_column = "ATA_DATETIME" if movement == "A" else "ATD_DATETIME"
        if _timestamp(row[movement_column]) is not None:
            continue

        original_date = pd.to_datetime(row["DATE OF FLIGHT"], errors="coerce")
        if pd.isna(original_date):
            continue
        original_date = pd.Timestamp(original_date).normalize()
        recovery_date = original_date + pd.Timedelta(days=1)
        window_start = original_date + pd.Timedelta(hours=23)
        window_end = recovery_date + pd.Timedelta(hours=6)

        candidate_mask = ~working["USED_FOR_RECOVERY"].fillna(False).astype(bool)
        candidate_mask &= parsed_dates.eq(recovery_date)
        for column in identity_columns:
            candidate_mask &= working[column].map(normalize_code).eq(
                normalize_code(row[column])
            )
        register = normalize_code(row["AC REGISTER"])
        if register:
            candidate_mask &= working["AC REGISTER"].map(normalize_code).eq(register)
        candidate_mask.loc[index] = False

        candidate_indices = working.index[candidate_mask]
        evaluations: list[tuple[tuple[object, ...], int, pd.Timestamp]] = []
        eobt_datetime = _timestamp(row.get("EOBT_DATETIME")) or window_start
        for candidate_index in candidate_indices:
            candidate_datetime = _timestamp(
                working.at[candidate_index, movement_column]
            )
            if candidate_datetime is None:
                continue
            if candidate_datetime < window_start or candidate_datetime > window_end:
                continue
            if candidate_datetime < eobt_datetime:
                continue
            time_distance = abs(
                (candidate_datetime - eobt_datetime).total_seconds()
            )
            candidate_timestamp = parse_priority_timestamp(
                working.at[candidate_index, "TIMESTAMP"]
            )
            timestamp_priority = -(
                candidate_timestamp.value if candidate_timestamp is not None else 0
            )
            evaluations.append(
                (
                    (
                        time_distance,
                        timestamp_priority,
                        -parse_message_number(
                            working.at[candidate_index, "MESSAGE NUM"]
                        ),
                        -int(working.at[candidate_index, "SOURCE ROW"]),
                    ),
                    int(candidate_index),
                    candidate_datetime,
                )
            )
        if not evaluations:
            continue

        _, candidate_index, recovered_movement_datetime = min(
            evaluations, key=lambda item: item[0]
        )
        candidate = working.loc[candidate_index]
        for column in recover_display_fields + recover_internal_fields:
            if not is_present(working.at[index, column]) and is_present(candidate[column]):
                working.at[index, column] = candidate[column]

        recovered_atd = _timestamp(working.at[index, "ATD_DATETIME"])
        recovered_ata = _timestamp(working.at[index, "ATA_DATETIME"])
        recovered_movement_datetime = movement_datetime(
            movement, recovered_atd, recovered_ata
        ) or recovered_movement_datetime
        working.at[index, "MOVEMENT_DATETIME"] = recovered_movement_datetime
        working.at[index, "DAT_MOVEMENT_DATETIME"] = recovered_movement_datetime
        working.at[index, "MOVEMENT_TIME_DISPLAY"] = format_time(
            recovered_movement_datetime
        )
        working.at[index, "DAT_MOVEMENT_TIME_DISPLAY"] = format_time(
            recovered_movement_datetime
        )
        working.at[index, "ACTUAL MOVEMENT DATE"] = recovered_movement_datetime.strftime(
            "%Y-%m-%d"
        )
        working.at[index, "MOVEMENT_DATE_DIFFERS_FROM_FLIGHT_DATE"] = True
        working.at[index, "DAT_RECOVERY_USED"] = True
        working.at[index, "DAT_RECOVERY_REASON"] = (
            "ADJACENT DATE MIDNIGHT RECOVERY"
        )
        working.at[index, "DAT_RECOVERY_SOURCE_DATE"] = candidate["DATE OF FLIGHT"]
        working.at[index, "DAT_RECOVERY_SOURCE_ROW"] = int(candidate["SOURCE ROW"])
        working.at[index, "ORIGINAL_DAT_DATE"] = row["DATE OF FLIGHT"]
        working.at[index, "RECOVERED_DAT_DATE"] = candidate["DATE OF FLIGHT"]
        working.at[candidate_index, "USED_FOR_RECOVERY"] = True
        working.at[candidate_index, "DAT_RECOVERY_REASON"] = (
            "USED AS ADJACENT DATE MIDNIGHT RECOVERY SOURCE"
        )

    return working


def _prepare_keys(frame: pd.DataFrame, side: str) -> pd.DataFrame:
    prepared = frame.copy().reset_index(drop=True)
    key_columns: list[str] = []
    for column, field in MATCH_KEY_SPECS:
        key_column = f"__key_{field}"
        if field == "eobd":
            prepared[key_column] = prepared[column].map(normalize_date)
        else:
            prepared[key_column] = prepared[column].map(normalize_code)
        key_columns.append(key_column)

    prepared["__complete_key"] = prepared[key_columns].ne("").all(axis=1)
    if prepared.empty:
        prepared["__match_key"] = pd.Series(dtype=object)
        prepared["__movement_match_key"] = pd.Series(dtype=object)
        prepared["__complete_movement_key"] = pd.Series(dtype=bool)
        prepared["DUPLICATE_GROUP_KEY"] = pd.Series(dtype=object)
        return prepared
    prepared["__match_key"] = prepared[key_columns].astype(str).agg("␟".join, axis=1)
    prepared["DUPLICATE_GROUP_KEY"] = prepared[
        [column for column, _ in MATCH_KEY_SPECS]
    ].apply(lambda row: duplicate_group_key(row.tolist()), axis=1)
    incomplete = ~prepared["__complete_key"]
    prepared.loc[incomplete, "__match_key"] = prepared.loc[incomplete].apply(
        lambda row: f"__INCOMPLETE__{side}__{row['SOURCE DATA']}__{row['SOURCE ROW']}",
        axis=1,
    )
    movement_dates = prepared["MOVEMENT_DATETIME"].map(
        lambda value: (
            pd.Timestamp(value).strftime("%Y-%m-%d")
            if _timestamp(value) is not None
            else ""
        )
    )
    movement_key_columns = [
        movement_dates.rename("__movement_date"),
        prepared["__key_flight"],
        prepared["__key_adep"],
        prepared["__key_ades"],
        prepared["__key_movement"],
    ]
    movement_key_frame = pd.concat(movement_key_columns, axis=1)
    prepared["__complete_movement_key"] = movement_key_frame.ne("").all(axis=1)
    prepared["__movement_match_key"] = movement_key_frame.astype(str).agg(
        "␟".join, axis=1
    )
    incomplete_movement = ~prepared["__complete_movement_key"]
    prepared.loc[incomplete_movement, "__movement_match_key"] = prepared.loc[
        incomplete_movement
    ].apply(
        lambda row: (
            f"__INCOMPLETE_MOVEMENT__{side}__"
            f"{row['SOURCE DATA']}__{row['SOURCE ROW']}"
        ),
        axis=1,
    )
    return prepared


def _duplicate_reason(row: pd.Series, selected: pd.Series) -> str:
    if bool(selected["HAS_MOVEMENT_TIME"]) and not bool(row["HAS_MOVEMENT_TIME"]):
        return "MOVEMENT TIME EMPTY; BETTER RECORD SELECTED"
    if int(row["COMPLETENESS_SCORE"]) < int(selected["COMPLETENESS_SCORE"]):
        return "LOWER COMPLETENESS SCORE"
    row_timestamp = row["__priority_timestamp"]
    selected_timestamp = selected["__priority_timestamp"]
    if row_timestamp < selected_timestamp:
        return "OLDER TIMESTAMP"
    if float(row["__message_number"]) < float(selected["__message_number"]):
        return "LOWER MESSAGE NUMBER"
    return "LOWER PRIORITY DUPLICATE; STABLE TIE-BREAK"


def _movement_time_bucket(value: object, hours: int = 6) -> str:
    timestamp = _timestamp(value)
    if timestamp is None:
        return "NO-TIME"
    bucket_hour = (timestamp.hour // hours) * hours
    return f"{bucket_hour:02d}-{bucket_hour + hours - 1:02d}"


def _assign_flight_instance_keys(working: pd.DataFrame) -> pd.DataFrame:
    """Resolve DAT records into movement instances before selecting duplicates."""
    result = working.copy()
    result["MOVEMENT_TIME_BUCKET"] = result["DAT_MOVEMENT_DATETIME"].map(
        _movement_time_bucket
    )
    result["FLIGHT_INSTANCE_KEY"] = ""
    key_columns = [f"__key_{field}" for _, field in MATCH_KEY_SPECS]

    for _, group in result.groupby(key_columns, dropna=False, sort=False):
        active = group.loc[~group["USED_FOR_RECOVERY"].fillna(False).astype(bool)]
        movement_rows = active.loc[active["DAT_MOVEMENT_DATETIME"].notna()].copy()
        resolved_register: dict[int, str] = {}

        if not movement_rows.empty:
            movement_rows["__actual_date"] = movement_rows[
                "DAT_MOVEMENT_DATETIME"
            ].map(lambda value: pd.Timestamp(value).strftime("%Y-%m-%d"))
            movement_rows["__bucket"] = movement_rows["DAT_MOVEMENT_DATETIME"].map(
                _movement_time_bucket
            )
            for _, same_slot in movement_rows.groupby(
                ["__actual_date", "__bucket"], sort=False
            ):
                registers = {
                    normalize_code(value)
                    for value in same_slot["AC REGISTER"]
                    if normalize_code(value)
                }
                slot_register = next(iter(registers)) if len(registers) == 1 else ""
                for row_index, slot_row in same_slot.iterrows():
                    resolved_register[int(row_index)] = (
                        normalize_code(slot_row["AC REGISTER"]) or slot_register
                    )

        movement_instance_map: dict[tuple[str, str, str], dict[str, object]] = {}
        for row_index, movement_row in movement_rows.iterrows():
            movement_dt = pd.Timestamp(movement_row["DAT_MOVEMENT_DATETIME"])
            actual_date = movement_dt.strftime("%Y-%m-%d")
            bucket = _movement_time_bucket(movement_dt)
            register = resolved_register.get(
                int(row_index), normalize_code(movement_row["AC REGISTER"])
            )
            instance_token = (actual_date, bucket, register)
            existing = movement_instance_map.get(instance_token)
            candidate = {
                "index": int(row_index),
                "datetime": movement_dt,
                "actual_date": actual_date,
                "bucket": bucket,
                "register": register,
            }
            if existing is None or movement_dt < existing["datetime"]:
                movement_instance_map[instance_token] = candidate
        movement_instances = list(movement_instance_map.values())

        for row_index, row in group.iterrows():
            if bool(row.get("USED_FOR_RECOVERY", False)):
                actual_date = clean_text(row.get("ACTUAL MOVEMENT DATE"))
                bucket = _movement_time_bucket(row.get("DAT_MOVEMENT_DATETIME"))
                register = normalize_code(row.get("AC REGISTER"))
                suffix = f"RECOVERY-SOURCE-{row.get('SOURCE DATA', '')}-{row.get('SOURCE ROW', '')}"
            else:
                movement_dt = _timestamp(row.get("DAT_MOVEMENT_DATETIME"))
                if movement_dt is not None:
                    actual_date = movement_dt.strftime("%Y-%m-%d")
                    bucket = _movement_time_bucket(movement_dt)
                    register = resolved_register.get(
                        int(row_index), normalize_code(row.get("AC REGISTER"))
                    )
                    suffix = ""
                else:
                    row_register = normalize_code(row.get("AC REGISTER"))
                    candidates = movement_instances
                    matching_register = [
                        item
                        for item in candidates
                        if row_register and item["register"] == row_register
                    ]
                    if matching_register:
                        candidates = matching_register
                    eobt_datetime = _timestamp(row.get("EOBT_DATETIME"))
                    selected_instance = None
                    if len(candidates) == 1:
                        selected_instance = candidates[0]
                    elif candidates and eobt_datetime is not None:
                        selected_instance = min(
                            candidates,
                            key=lambda item: abs(
                                (item["datetime"] - eobt_datetime).total_seconds()
                            ),
                        )
                    if selected_instance is not None:
                        actual_date = str(selected_instance["actual_date"])
                        bucket = str(selected_instance["bucket"])
                        register = str(selected_instance["register"] or row_register)
                        suffix = ""
                    else:
                        fallback_datetime = _timestamp(row.get("EOBT_DATETIME"))
                        actual_date = clean_text(row.get("DATE OF FLIGHT"))
                        bucket = _movement_time_bucket(fallback_datetime)
                        register = row_register
                        suffix = (
                            f"AMBIGUOUS-{row.get('SOURCE DATA', '')}-{row.get('SOURCE ROW', '')}"
                            if candidates
                            else ""
                        )

            instance_values = [
                row.get("DATE OF FLIGHT", ""),
                row.get("FLIGHT NUMBER", ""),
                row.get("AERODROME", ""),
                row.get("TO FROM", ""),
                row.get("D/A/L/O", ""),
                register,
                actual_date,
                bucket,
            ]
            if suffix:
                instance_values.append(suffix)
            instance_key = duplicate_group_key(instance_values)
            result.at[row_index, "MOVEMENT_TIME_BUCKET"] = bucket
            result.at[row_index, "FLIGHT_INSTANCE_KEY"] = instance_key

    result["DUPLICATE_GROUP_KEY"] = result["FLIGHT_INSTANCE_KEY"]
    return result


def deduplicate_dat(dat_prepared: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Select one best DAT record per flight instance after recovery."""
    working = dat_prepared.copy()
    if "USED_FOR_RECOVERY" not in working.columns:
        working["USED_FOR_RECOVERY"] = False
    working = _assign_flight_instance_keys(working)
    working["COMPLETENESS_SCORE"] = working[COMPLETENESS_FIELDS].apply(
        lambda row: sum(is_present(value) for value in row), axis=1
    )
    working["HAS_MOVEMENT_TIME"] = working["DAT_MOVEMENT_DATETIME"].notna()
    working["__priority_timestamp"] = working["TIMESTAMP"].map(parse_priority_timestamp)
    working["__priority_timestamp"] = working["__priority_timestamp"].map(
        lambda value: value if value is not None else pd.Timestamp.min
    )
    working["__message_number"] = working["MESSAGE NUM"].map(parse_message_number)

    eligible = working.loc[
        working["__complete_key"]
        & ~working["USED_FOR_RECOVERY"].fillna(False).astype(bool)
    ]
    complete = eligible.sort_values(
        ["FLIGHT_INSTANCE_KEY"]
        + [
            "HAS_MOVEMENT_TIME",
            "COMPLETENESS_SCORE",
            "__priority_timestamp",
            "__message_number",
            "SOURCE ROW",
        ],
        ascending=[True, False, False, False, False, False],
        kind="stable",
    )
    selected_complete_indices = complete.groupby(
        "FLIGHT_INSTANCE_KEY", dropna=False, sort=False
    ).head(1).index
    selected_indices = set(selected_complete_indices) | set(
        working.index[
            ~working["__complete_key"]
            & ~working["USED_FOR_RECOVERY"].fillna(False).astype(bool)
        ]
    )
    working["SELECTED_RECORD_FLAG"] = working.index.map(
        lambda index: index in selected_indices
    )

    dat_unique = working.loc[working["SELECTED_RECORD_FLAG"]].copy()
    duplicate_dat = working.loc[~working["SELECTED_RECORD_FLAG"]].copy()
    selected_by_key = {
        str(row["FLIGHT_INSTANCE_KEY"]): row for _, row in dat_unique.iterrows()
    }
    if not duplicate_dat.empty:
        def duplicate_reason(row: pd.Series) -> str:
            if bool(row.get("USED_FOR_RECOVERY", False)):
                return "USED FOR ADJACENT DATE MIDNIGHT RECOVERY"
            selected = selected_by_key.get(str(row["FLIGHT_INSTANCE_KEY"]))
            if selected is None:
                return "NON-SELECTED FLIGHT INSTANCE"
            return _duplicate_reason(row, selected)

        duplicate_dat["DUPLICATE_REASON"] = duplicate_dat.apply(
            duplicate_reason, axis=1
        )
    else:
        duplicate_dat["DUPLICATE_REASON"] = pd.Series(dtype=object)

    dat_unique["SELECTED DAT RECORD"] = True
    duplicate_dat["SELECTED DAT RECORD"] = False
    duplicate_dat["STATUS"] = "DUPLICATE DAT"
    duplicate_dat["MATCH_REASON"] = duplicate_dat["DUPLICATE_REASON"]
    duplicate_dat["DUPLICATE_GROUP_KEY"] = duplicate_dat["DUPLICATE_GROUP_KEY"]
    return dat_unique.reset_index(drop=True), duplicate_dat.reset_index(drop=True)


def _timestamp(value: object) -> pd.Timestamp | None:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return pd.Timestamp(value)


def _match_key_for_date(row: pd.Series, match_date: object) -> str | None:
    values = [
        normalize_date(match_date),
        normalize_code(row.get("FLIGHT NUMBER", "")),
        normalize_code(row.get("AERODROME", "")),
        normalize_code(row.get("TO FROM", "")),
        normalize_code(row.get("D/A/L/O", "")),
    ]
    if not all(values):
        return None
    return "␟".join(values)


def _stream_search_specs(dat_row: pd.Series) -> list[dict[str, str]]:
    """Return actual-movement lookup first, followed by date fallbacks."""
    requested = [
        (
            dat_row.get("ACTUAL MOVEMENT DATE", ""),
            "ACTUAL MOVEMENT DATE",
            "ACTUAL MOVEMENT DATE MATCH",
            "MOVEMENT DATE",
        ),
        (
            dat_row.get("DATE OF FLIGHT", ""),
            "DATE OF FLIGHT",
            "ORIGINAL DATE MATCH",
            "DATE OF FLIGHT",
        )
    ]
    if bool(dat_row.get("DAT_RECOVERY_USED", False)):
        requested.append(
            (
                dat_row.get("RECOVERED_DAT_DATE", ""),
                "RECOVERED_DAT_DATE",
                "RECOVERED DATE MATCH",
                "DATE OF FLIGHT",
            )
        )

    specs: list[dict[str, str]] = []
    seen_lookups: set[tuple[str, str]] = set()
    for match_date, date_used, mode, lookup_index in requested:
        key = _match_key_for_date(dat_row, match_date)
        lookup_token = (lookup_index, key or "")
        if key is None or lookup_token in seen_lookups:
            continue
        seen_lookups.add(lookup_token)
        specs.append(
            {
                "key": key,
                "date_used": date_used,
                "mode": mode,
                "date_value": normalize_date(match_date),
                "lookup_index": lookup_index,
            }
        )
    return specs


def _candidate_evaluation(
    dat_row: pd.Series,
    stream_row: pd.Series,
    invalid_statuses: set[str],
    treat_invalid_status_as_missing: bool,
    time_tolerance_minutes: int,
    match_date_used: str,
    match_mode: str,
) -> dict[str, object]:
    dat_datetime = _timestamp(dat_row["DAT_MOVEMENT_DATETIME"])
    stream_datetime = _timestamp(stream_row["STREAM_MOVEMENT_DATETIME"])
    stream_status = normalize_code(stream_row["STREAM STATUS FLIGHT"])
    dat_register = normalize_code(dat_row.get("AC REGISTER", ""))
    stream_register = normalize_code(stream_row.get("AC REGISTER", ""))
    register_mismatch = bool(
        dat_register and stream_register and dat_register != stream_register
    )
    invalid_status = bool(
        treat_invalid_status_as_missing and stream_status in invalid_statuses
    )
    time_difference: float | None = None
    date_mismatch = False
    allowed_movement_dates: set[str] = set()
    if dat_datetime is not None:
        allowed_movement_dates.add(dat_datetime.strftime("%Y-%m-%d"))
    if bool(dat_row.get("DAT_RECOVERY_USED", False)):
        for allowed_date in (
            dat_row.get("ACTUAL MOVEMENT DATE", ""),
            dat_row.get("RECOVERED_DAT_DATE", ""),
        ):
            normalized_allowed_date = normalize_date(allowed_date)
            if normalized_allowed_date:
                allowed_movement_dates.add(normalized_allowed_date)
    if dat_datetime is not None and stream_datetime is not None:
        time_difference = abs(
            (stream_datetime - dat_datetime).total_seconds() / 60.0
        )
        date_mismatch = (
            stream_datetime.strftime("%Y-%m-%d") not in allowed_movement_dates
        )

    movement_field = "ATA" if dat_row["D/A/L/O"] == "A" else "ATD"
    invalid_reasons: list[str] = []
    if invalid_status:
        invalid_reasons.append(f"STREAM STATUS {stream_status}")
    if dat_datetime is None:
        invalid_reasons.append("DAT INVALID MOVEMENT TIME")
    if stream_datetime is None:
        invalid_reasons.append("STREAM INVALID MOVEMENT TIME")
    if date_mismatch:
        invalid_reasons.append(f"STREAM INVALID {movement_field} DATE")

    priority_timestamp = parse_priority_timestamp(stream_row["TIMESTAMP"])
    priority_value = (
        -priority_timestamp.value if priority_timestamp is not None else float("inf")
    )
    selection_sort = (
        1 if invalid_status else 0,
        1 if stream_datetime is None else 0,
        1 if date_mismatch else 0,
        (
            1
            if time_difference is None
            or time_difference > time_tolerance_minutes
            else 0
        ),
        1 if register_mismatch else 0,
        time_difference if time_difference is not None else float("inf"),
        {
            "ACTUAL MOVEMENT DATE MATCH": 0,
            "RECOVERED DATE MATCH": 1,
            "ORIGINAL DATE MATCH": 2,
        }.get(match_mode, 3),
        priority_value,
        -parse_message_number(stream_row["MESSAGE NUM"]),
        -int(stream_row["SOURCE ROW"]),
    )
    validation_result = {
        "RECOVERED DATE MATCH": "VALID STREAM CANDIDATE BY RECOVERED DATE",
        "ACTUAL MOVEMENT DATE MATCH": (
            "VALID STREAM CANDIDATE BY ACTUAL MOVEMENT DATE"
        ),
    }.get(match_mode, "VALID STREAM CANDIDATE")
    if invalid_reasons:
        validation_result = " / ".join(invalid_reasons)
    if register_mismatch:
        register_note = "STREAM AC REGISTER MISMATCH"
        validation_result = (
            f"{validation_result} / {register_note}"
            if validation_result
            else register_note
        )
    return {
        "stream_row": stream_row,
        "stream_status": stream_status,
        "dat_datetime": dat_datetime,
        "stream_datetime": stream_datetime,
        "time_difference": time_difference,
        "date_mismatch": date_mismatch,
        "allowed_movement_dates": sorted(allowed_movement_dates),
        "register_mismatch": register_mismatch,
        "invalid_reasons": invalid_reasons,
        "validation_result": validation_result,
        "match_date_used": match_date_used,
        "match_mode": match_mode,
        "selection_sort": selection_sort,
    }


def _result_from_dat(
    dat_row: pd.Series,
    status: str,
    reason: str,
    selected_evaluation: Mapping[str, object] | None,
    validasi: str | None = None,
) -> dict[str, object]:
    result = {column: dat_row.get(column, "") for column in RESULT_COLUMNS}
    stream_row = (
        selected_evaluation.get("stream_row") if selected_evaluation else None
    )
    stream_status = (
        selected_evaluation.get("stream_status", "") if selected_evaluation else ""
    )
    stream_datetime = (
        selected_evaluation.get("stream_datetime") if selected_evaluation else None
    )
    time_difference = (
        selected_evaluation.get("time_difference") if selected_evaluation else None
    )
    validation_result = (
        selected_evaluation.get("validation_result", "STREAM NOT FOUND")
        if selected_evaluation
        else "STREAM NOT FOUND"
    )
    match_date_used = (
        selected_evaluation.get("match_date_used", "") if selected_evaluation else ""
    )
    match_mode = (
        selected_evaluation.get("match_mode", "NO STREAM MATCH")
        if selected_evaluation
        else "NO STREAM MATCH"
    )
    if validasi is None:
        if selected_evaluation is None:
            validasi = VALIDATION_NOT_FOUND
        elif status == "MATCHED":
            validasi = VALIDATION_FOUND
        else:
            validasi = VALIDATION_REVIEW
    result.update(
        {
            "VALIDASI": validasi,
            "SOURCE DATA": dat_row["SOURCE DATA"],
            "SOURCE ROW": int(dat_row["SOURCE ROW"]),
            "STATUS": status,
            "MATCH_REASON": reason,
            "STREAM STATUS FLIGHT": stream_status,
            "DAT MOVEMENT DATETIME": format_datetime(
                dat_row["DAT_MOVEMENT_DATETIME"]
            ),
            "STREAM MOVEMENT DATETIME": format_datetime(stream_datetime),
            "TIME DIFFERENCE MINUTES": (
                round(float(time_difference), 1)
                if time_difference is not None
                else ""
            ),
            "DUPLICATE GROUP KEY": dat_row["DUPLICATE_GROUP_KEY"],
            "SELECTED DAT RECORD": True,
            "STREAM MATCH KEY": dat_row["DUPLICATE_GROUP_KEY"],
            "STREAM VALIDATION RESULT": validation_result,
            "STREAM MATCH DATE USED": match_date_used,
            "STREAM MATCH MODE": match_mode,
            "BILLING CATEGORY": "BILLING REVIEW",
            "DAT_RECOVERY_USED": bool(dat_row.get("DAT_RECOVERY_USED", False)),
            "DAT_RECOVERY_REASON": dat_row.get("DAT_RECOVERY_REASON", ""),
            "DAT_RECOVERY_SOURCE_DATE": dat_row.get(
                "DAT_RECOVERY_SOURCE_DATE", ""
            ),
            "DAT_RECOVERY_SOURCE_ROW": dat_row.get(
                "DAT_RECOVERY_SOURCE_ROW", ""
            ),
            "ORIGINAL_DAT_DATE": dat_row.get("ORIGINAL_DAT_DATE", ""),
            "RECOVERED_DAT_DATE": dat_row.get("RECOVERED_DAT_DATE", ""),
            "USED_FOR_RECOVERY": bool(dat_row.get("USED_FOR_RECOVERY", False)),
            "FLIGHT_INSTANCE_KEY": dat_row.get("FLIGHT_INSTANCE_KEY", ""),
            "MOVEMENT_TIME_BUCKET": dat_row.get("MOVEMENT_TIME_BUCKET", ""),
        }
    )
    if isinstance(stream_row, pd.Series):
        result["STREAM SOURCE ROW"] = int(stream_row["SOURCE ROW"])
    return result


def _result_from_stream(stream_row: pd.Series) -> dict[str, object]:
    result = {column: stream_row.get(column, "") for column in RESULT_COLUMNS}
    result.update(
        {
            "VALIDASI": "",
            "SOURCE DATA": "STREAM",
            "SOURCE ROW": int(stream_row["SOURCE ROW"]),
            "STATUS": "EXTRA IN STREAM",
            "MATCH_REASON": "DAT NOT FOUND",
            "STREAM STATUS FLIGHT": stream_row["STREAM STATUS FLIGHT"],
            "DAT MOVEMENT DATETIME": "",
            "STREAM MOVEMENT DATETIME": format_datetime(
                stream_row["STREAM_MOVEMENT_DATETIME"]
            ),
            "TIME DIFFERENCE MINUTES": "",
            "DUPLICATE GROUP KEY": "",
            "SELECTED DAT RECORD": False,
            "STREAM MATCH KEY": stream_row["DUPLICATE_GROUP_KEY"],
            "STREAM VALIDATION RESULT": "DAT NOT FOUND",
            "STREAM MATCH DATE USED": "",
            "STREAM MATCH MODE": "NO STREAM MATCH",
            "BILLING CATEGORY": "",
            "DAT_RECOVERY_USED": False,
            "DAT_RECOVERY_REASON": "",
            "DAT_RECOVERY_SOURCE_DATE": "",
            "DAT_RECOVERY_SOURCE_ROW": "",
            "ORIGINAL_DAT_DATE": "",
            "RECOVERED_DAT_DATE": "",
            "USED_FOR_RECOVERY": False,
            "FLIGHT_INSTANCE_KEY": "",
            "MOVEMENT_TIME_BUCKET": stream_row.get("MOVEMENT_TIME_BUCKET", ""),
        }
    )
    return result


def _records_frame(records: list[dict[str, object]]) -> pd.DataFrame:
    frame = pd.DataFrame(records)
    for column in DETAIL_COLUMNS:
        if column not in frame.columns:
            frame[column] = ""
    if not frame.empty:
        frame = frame.sort_values(
            ["DATE OF FLIGHT", "FLIGHT NUMBER", "D/A/L/O"], kind="stable"
        ).reset_index(drop=True)
    return frame


def _duplicate_output(duplicates: pd.DataFrame) -> pd.DataFrame:
    if duplicates.empty:
        return pd.DataFrame(
            columns=RESULT_COLUMNS
            + ["SOURCE DATA", "SOURCE ROW", "STATUS", "MATCH_REASON"]
            + DUPLICATE_AUDIT_COLUMNS
        )
    output = duplicates.copy()
    output["STATUS"] = "DUPLICATE DAT"
    output["MATCH_REASON"] = output["DUPLICATE_REASON"]
    columns = (
        RESULT_COLUMNS
        + ["SOURCE DATA", "SOURCE ROW", "STATUS", "MATCH_REASON"]
        + DUPLICATE_AUDIT_COLUMNS
    )
    return output[columns].sort_values(
        ["DATE OF FLIGHT", "FLIGHT NUMBER", "D/A/L/O"], kind="stable"
    ).reset_index(drop=True)


def _excluded_output(dat_excluded: pd.DataFrame, stream_excluded: pd.DataFrame) -> pd.DataFrame:
    records = [
        record
        for frame in (dat_excluded, stream_excluded)
        for record in frame.to_dict("records")
    ]
    columns = RESULT_COLUMNS + [
        "SOURCE DATA",
        "SOURCE ROW",
        "STATUS",
        "MATCH_REASON",
    ]
    if not records:
        return pd.DataFrame(columns=columns)
    output = pd.DataFrame.from_records(records)
    output["STATUS"] = "EXCLUDED NON-BILLABLE"
    output["MATCH_REASON"] = "FLIGHT NUMBER MATCHES INTERNAL/NON-BILLABLE RULE"
    return output[columns].sort_values(
        ["SOURCE DATA", "DATE OF FLIGHT", "FLIGHT NUMBER"], kind="stable"
    ).reset_index(drop=True)


def reconcile_dat_vs_stream(
    dat_dep: pd.DataFrame,
    dat_arr: pd.DataFrame,
    stream: pd.DataFrame,
    dep_mapping: Mapping[str, str | None] | None = None,
    arr_mapping: Mapping[str, str | None] | None = None,
    stream_mapping: Mapping[str, str | None] | None = None,
    time_tolerance_minutes: int = 30,
    invalid_stream_statuses: Iterable[str] | None = None,
    treat_invalid_stream_status_as_missing: bool = True,
) -> dict[str, object]:
    dep_mapping = dict(dep_mapping or detected_mapping(dat_dep))
    arr_mapping = dict(arr_mapping or detected_mapping(dat_arr))
    stream_mapping = dict(stream_mapping or detected_mapping(stream))
    issues = validate_required_columns(dep_mapping, arr_mapping, stream_mapping)
    if issues:
        raise ValueError("Kolom wajib belum lengkap: " + " | ".join(issues))
    if time_tolerance_minutes <= 0 or time_tolerance_minutes > 120:
        raise ValueError("Time tolerance harus lebih dari 0 dan maksimal 120 menit.")

    invalid_statuses = {
        normalize_code(value)
        for value in (invalid_stream_statuses or DEFAULT_INVALID_STREAM_STATUSES)
        if normalize_code(value)
    }
    dep_normalized = standardize_dataset(dat_dep, dep_mapping, "DAT DEP", "D")
    arr_normalized = standardize_dataset(dat_arr, arr_mapping, "DAT ARR", "A")
    stream_normalized = standardize_dataset(stream, stream_mapping, "STREAM")
    if dep_normalized.empty:
        dat_combined_normalized = arr_normalized.copy().reset_index(drop=True)
    elif arr_normalized.empty:
        dat_combined_normalized = dep_normalized.copy().reset_index(drop=True)
    else:
        dat_combined_normalized = pd.DataFrame.from_records(
            dep_normalized.to_dict("records") + arr_normalized.to_dict("records"),
            columns=STANDARDIZED_COLUMNS,
        )

    dat_excluded_mask = dat_combined_normalized["FLIGHT NUMBER"].map(
        is_excluded_flight_number
    )
    stream_excluded_mask = stream_normalized["FLIGHT NUMBER"].map(
        is_excluded_flight_number
    )
    dat_excluded = dat_combined_normalized.loc[dat_excluded_mask].copy()
    stream_excluded = stream_normalized.loc[stream_excluded_mask].copy()
    dat_combined = dat_combined_normalized.loc[~dat_excluded_mask].copy().reset_index(
        drop=True
    )
    stream_standard = stream_normalized.loc[~stream_excluded_mask].copy().reset_index(
        drop=True
    )
    dep = dat_combined.loc[dat_combined["SOURCE DATA"].eq("DAT DEP")].copy()
    arr = dat_combined.loc[dat_combined["SOURCE DATA"].eq("DAT ARR")].copy()
    raw_dat_normalized = dat_combined.copy(deep=True)
    dat_recovered = recover_adjacent_date_movements(raw_dat_normalized)
    excluded_non_billable = _excluded_output(dat_excluded, stream_excluded)

    dat_prepared = _prepare_keys(dat_recovered, "DAT")
    stream_prepared = _prepare_keys(stream_standard, "STREAM")
    dat_unique, duplicate_raw = deduplicate_dat(dat_prepared)
    duplicate_dat = _duplicate_output(duplicate_raw)

    stream_groups_by_flight_date = {
        str(key): group.copy()
        for key, group in stream_prepared.loc[
            stream_prepared["__complete_key"]
        ].groupby("__match_key", sort=False)
    }
    stream_groups_by_movement_date = {
        str(key): group.copy()
        for key, group in stream_prepared.loc[
            stream_prepared["__complete_movement_key"]
        ].groupby("__movement_match_key", sort=False)
    }

    matched_records: list[dict[str, object]] = []
    missing_records: list[dict[str, object]] = []
    need_review_records: list[dict[str, object]] = []
    audit_records: list[dict[str, object]] = []
    candidate_stream_source_rows: set[int] = set()

    for _, dat_row in dat_unique.iterrows():
        search_specs = _stream_search_specs(dat_row)
        evaluations_by_source_row: dict[int, dict[str, object]] = {}
        for search_spec in search_specs:
            candidate_groups = (
                stream_groups_by_movement_date
                if search_spec["lookup_index"] == "MOVEMENT DATE"
                else stream_groups_by_flight_date
            )
            candidate_group = candidate_groups.get(search_spec["key"])
            if candidate_group is not None and not candidate_group.empty:
                for _, stream_row in candidate_group.iterrows():
                    evaluation = _candidate_evaluation(
                        dat_row,
                        stream_row,
                        invalid_statuses,
                        treat_invalid_stream_status_as_missing,
                        time_tolerance_minutes,
                        search_spec["date_used"],
                        search_spec["mode"],
                    )
                    source_row = int(stream_row["SOURCE ROW"])
                    candidate_stream_source_rows.add(source_row)
                    existing = evaluations_by_source_row.get(source_row)
                    if (
                        existing is None
                        or evaluation["selection_sort"] < existing["selection_sort"]
                    ):
                        evaluations_by_source_row[source_row] = evaluation
        evaluations = list(evaluations_by_source_row.values())
        if not evaluations:
            not_found_reason = (
                "STREAM NOT FOUND AFTER ACTUAL MOVEMENT, ORIGINAL, AND RECOVERED DATE SEARCH"
                if bool(dat_row.get("DAT_RECOVERY_USED", False))
                else "STREAM NOT FOUND AFTER ACTUAL MOVEMENT AND ORIGINAL DATE SEARCH"
            )
            result = _result_from_dat(
                dat_row, "MISSING IN STREAM", not_found_reason, None
            )
            missing_records.append(result)
            audit_records.append(result.copy())
            continue

        selected = min(evaluations, key=lambda item: item["selection_sort"])
        invalid_reasons = list(selected["invalid_reasons"])
        difference = selected["time_difference"]
        register_mismatch = bool(selected["register_mismatch"])
        if invalid_reasons:
            status = "MISSING IN STREAM"
            reason = " / ".join(invalid_reasons)
        elif difference is None:
            status = "MISSING IN STREAM"
            reason = "STREAM INVALID MOVEMENT TIME"
        elif float(difference) <= time_tolerance_minutes:
            if register_mismatch:
                status = "NEED REVIEW"
                reason = "STREAM AC REGISTER MISMATCH"
            else:
                status = "MATCHED"
                reason = {
                    "RECOVERED DATE MATCH": "VALID STREAM MATCH BY RECOVERED DATE",
                    "ACTUAL MOVEMENT DATE MATCH": (
                        "VALID STREAM MATCH BY ACTUAL MOVEMENT DATE"
                    ),
                }.get(str(selected["match_mode"]), "VALID STREAM MATCH")
        elif float(difference) <= 120:
            status = "NEED REVIEW"
            reason = (
                "STREAM TIME DIFFERENCE / STREAM AC REGISTER MISMATCH"
                if register_mismatch
                else "STREAM TIME DIFFERENCE"
            )
        else:
            status = "MISSING IN STREAM"
            reason = "STREAM TIME MISMATCH"

        result = _result_from_dat(dat_row, status, reason, selected)
        if status == "MATCHED":
            matched_records.append(result)
        elif status == "NEED REVIEW":
            need_review_records.append(result)
        else:
            missing_records.append(result)
        audit_records.append(result.copy())

        for evaluation in evaluations:
            if evaluation is selected:
                continue
            audit_row = _result_from_dat(
                dat_row,
                "STREAM CANDIDATE NOT SELECTED",
                "LOWER PRIORITY STREAM CANDIDATE",
                evaluation,
                validasi=str(result["VALIDASI"]),
            )
            audit_row["STREAM VALIDATION RESULT"] = (
                "NOT SELECTED / " + str(evaluation["validation_result"])
            )
            audit_records.append(audit_row)

    matched = _records_frame(matched_records)
    missing = _records_frame(missing_records)
    need_review = _records_frame(need_review_records)

    extra_rows = stream_prepared.loc[
        ~stream_prepared["SOURCE ROW"].astype(int).isin(candidate_stream_source_rows)
    ]
    extra_records = [_result_from_stream(row) for _, row in extra_rows.iterrows()]
    extra = _records_frame(extra_records)
    audit_records.extend(record.copy() for record in extra_records)
    for _, duplicate_row in duplicate_dat.iterrows():
        audit_record = {
            column: duplicate_row.get(column, "") for column in RESULT_COLUMNS
        }
        audit_record.update(
            {
                "SOURCE DATA": duplicate_row.get("SOURCE DATA", ""),
                "SOURCE ROW": duplicate_row.get("SOURCE ROW", ""),
                "STATUS": "DUPLICATE DAT",
                "MATCH_REASON": duplicate_row.get("DUPLICATE_REASON", ""),
                "STREAM STATUS FLIGHT": "",
                "DAT MOVEMENT DATETIME": format_datetime(
                    duplicate_row.get("DAT_MOVEMENT_DATETIME")
                ),
                "STREAM MOVEMENT DATETIME": "",
                "TIME DIFFERENCE MINUTES": "",
                "DUPLICATE GROUP KEY": duplicate_row.get(
                    "DUPLICATE_GROUP_KEY", ""
                ),
                "SELECTED DAT RECORD": False,
                "STREAM MATCH KEY": "",
                "STREAM VALIDATION RESULT": "NOT APPLICABLE",
                "STREAM MATCH DATE USED": "",
                "STREAM MATCH MODE": "NO STREAM MATCH",
                "BILLING CATEGORY": "BILLING REVIEW",
                "COMPLETENESS_SCORE": duplicate_row.get("COMPLETENESS_SCORE", ""),
                "HAS_MOVEMENT_TIME": duplicate_row.get("HAS_MOVEMENT_TIME", ""),
                "DAT_RECOVERY_USED": bool(
                    duplicate_row.get("DAT_RECOVERY_USED", False)
                ),
                "DAT_RECOVERY_REASON": duplicate_row.get(
                    "DAT_RECOVERY_REASON", ""
                ),
                "DAT_RECOVERY_SOURCE_DATE": duplicate_row.get(
                    "DAT_RECOVERY_SOURCE_DATE", ""
                ),
                "DAT_RECOVERY_SOURCE_ROW": duplicate_row.get(
                    "DAT_RECOVERY_SOURCE_ROW", ""
                ),
                "ORIGINAL_DAT_DATE": duplicate_row.get("ORIGINAL_DAT_DATE", ""),
                "RECOVERED_DAT_DATE": duplicate_row.get("RECOVERED_DAT_DATE", ""),
                "USED_FOR_RECOVERY": bool(
                    duplicate_row.get("USED_FOR_RECOVERY", False)
                ),
                "FLIGHT_INSTANCE_KEY": duplicate_row.get(
                    "FLIGHT_INSTANCE_KEY", ""
                ),
                "MOVEMENT_TIME_BUCKET": duplicate_row.get(
                    "MOVEMENT_TIME_BUCKET", ""
                ),
            }
        )
        audit_records.append(audit_record)
    audit_detail = _records_frame(audit_records)

    missing_billing = missing.copy().reset_index(drop=True)
    missing_non_billable = pd.DataFrame(columns=missing.columns)
    validasi = missing.loc[
        missing["VALIDASI"].eq(VALIDATION_NOT_FOUND)
    ].copy().reset_index(drop=True)

    unique_dat_total = len(dat_unique)
    matched_total = len(matched)
    validation_found_total = int(
        matched["VALIDASI"].eq(VALIDATION_FOUND).sum()
    )
    validation_review_total = int(
        missing["VALIDASI"].eq(VALIDATION_REVIEW).sum()
        + need_review["VALIDASI"].eq(VALIDATION_REVIEW).sum()
    )
    accuracy = (matched_total / unique_dat_total * 100) if unique_dat_total else 0.0
    summary = {
        "total_dat_dep": len(dep),
        "total_dat_arr": len(arr),
        "total_dat_combined": len(dat_combined),
        "total_dat_unique": unique_dat_total,
        "total_stream": len(stream_standard),
        "matched": matched_total,
        "missing_in_stream": len(missing),
        "missing_billing_review": len(missing_billing),
        "missing_non_billable_review": len(missing_non_billable),
        "need_review": len(need_review),
        "total_ada_di_stream": validation_found_total,
        "total_validasi": len(validasi),
        "total_perlu_review_stream": validation_review_total,
        "extra_in_stream": len(extra),
        "duplicate_dat": len(duplicate_dat),
        "excluded_dat_non_billable": len(dat_excluded),
        "excluded_stream_non_billable": len(stream_excluded),
        "accuracy_percentage": accuracy,
        "incomplete_dat_keys": int((~dat_unique["__complete_key"]).sum()),
        "incomplete_stream_keys": int((~stream_prepared["__complete_key"]).sum()),
    }
    summary_table = pd.DataFrame(
        [
            ("Total DAT DEP", summary["total_dat_dep"]),
            ("Total DAT ARR", summary["total_dat_arr"]),
            ("Total DAT Combined", summary["total_dat_combined"]),
            ("Total DAT Unique", summary["total_dat_unique"]),
            ("Total STREAM", summary["total_stream"]),
            ("Matched", summary["matched"]),
            ("Missing in Stream", summary["missing_in_stream"]),
            ("Missing Billing Review", summary["missing_billing_review"]),
            ("Need Review", summary["need_review"]),
            ("Total Ada di STREAM", summary["total_ada_di_stream"]),
            ("Total Validasi", summary["total_validasi"]),
            (
                "Total Perlu Review STREAM",
                summary["total_perlu_review_stream"],
            ),
            ("Extra in Stream", summary["extra_in_stream"]),
            ("Duplicate DAT", summary["duplicate_dat"]),
            ("Excluded DAT Non-Billable", summary["excluded_dat_non_billable"]),
            ("Excluded STREAM Non-Billable", summary["excluded_stream_non_billable"]),
            ("Accuracy Percentage", summary["accuracy_percentage"] / 100),
        ],
        columns=["METRIC", "VALUE"],
    )
    return {
        "summary": summary,
        "summary_table": summary_table,
        "missing": missing,
        "missing_billing": missing_billing,
        "missing_non_billable": missing_non_billable,
        "validasi": validasi,
        "matched": matched,
        "need_review": need_review,
        "extra": extra,
        "duplicates": duplicate_dat,
        "audit_detail": audit_detail,
        "excluded_non_billable": excluded_non_billable,
        "dat_unique": dat_unique,
        "dat_dep": dep,
        "dat_arr": arr,
        "stream": stream_standard,
        "raw_dat_normalized": raw_dat_normalized,
        "dat_recovered": dat_recovered,
        "settings": {
            "time_tolerance_minutes": time_tolerance_minutes,
            "invalid_stream_statuses": sorted(invalid_statuses),
            "treat_invalid_stream_status_as_missing": treat_invalid_stream_status_as_missing,
        },
    }
