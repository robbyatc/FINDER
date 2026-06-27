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
    is_non_billable_flight,
    is_present,
    movement_datetime,
    normalize_code,
    normalize_date,
    normalize_display_text,
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
    "BILLING CATEGORY",
]

DETAIL_COLUMNS = RESULT_COLUMNS + AUDIT_COLUMNS

DUPLICATE_AUDIT_COLUMNS = [
    "DUPLICATE_GROUP_KEY",
    "SELECTED_RECORD_FLAG",
    "DUPLICATE_REASON",
    "COMPLETENESS_SCORE",
    "HAS_MOVEMENT_TIME",
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
]


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
            "ATD_DATETIME": atd_datetime,
            "ATA_DATETIME": ata_datetime,
            "MOVEMENT_DATETIME": move_datetime,
            "MOVEMENT_TIME_DISPLAY": normalized_times[
                "ATA_TIME_DISPLAY" if move == "A" else "ATD_TIME_DISPLAY"
            ],
            "MOVEMENT_DATE_DIFFERS_FROM_FLIGHT_DATE": movement_date_differs,
            "RAW ATD": clean_text(raw["atd"]),
            "RAW ATA": clean_text(raw["ata"]),
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


def deduplicate_dat(dat_prepared: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Select exactly one best DAT record per complete base key."""
    working = dat_prepared.copy()
    working["COMPLETENESS_SCORE"] = working[COMPLETENESS_FIELDS].apply(
        lambda row: sum(is_present(value) for value in row), axis=1
    )
    working["HAS_MOVEMENT_TIME"] = working["DAT_MOVEMENT_DATETIME"].notna()
    working["__priority_timestamp"] = working["TIMESTAMP"].map(parse_priority_timestamp)
    working["__priority_timestamp"] = working["__priority_timestamp"].map(
        lambda value: value if value is not None else pd.Timestamp.min
    )
    working["__message_number"] = working["MESSAGE NUM"].map(parse_message_number)

    key_columns = [f"__key_{field}" for _, field in MATCH_KEY_SPECS]
    complete = working.loc[working["__complete_key"]].sort_values(
        key_columns
        + [
            "HAS_MOVEMENT_TIME",
            "COMPLETENESS_SCORE",
            "__priority_timestamp",
            "__message_number",
            "SOURCE ROW",
        ],
        ascending=[True] * len(key_columns) + [False, False, False, False, False],
        kind="stable",
    )
    selected_complete_indices = complete.groupby(
        key_columns, dropna=False, sort=False
    ).head(1).index
    selected_indices = set(selected_complete_indices) | set(
        working.index[~working["__complete_key"]]
    )
    working["SELECTED_RECORD_FLAG"] = working.index.map(
        lambda index: index in selected_indices
    )

    dat_unique = working.loc[working["SELECTED_RECORD_FLAG"]].copy()
    duplicate_dat = working.loc[~working["SELECTED_RECORD_FLAG"]].copy()
    selected_by_key = {
        str(row["__match_key"]): row for _, row in dat_unique.iterrows()
    }
    if not duplicate_dat.empty:
        duplicate_dat["DUPLICATE_REASON"] = duplicate_dat.apply(
            lambda row: _duplicate_reason(
                row, selected_by_key[str(row["__match_key"])]
            ),
            axis=1,
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


def _candidate_evaluation(
    dat_row: pd.Series,
    stream_row: pd.Series,
    invalid_statuses: set[str],
    treat_invalid_status_as_missing: bool,
) -> dict[str, object]:
    dat_datetime = _timestamp(dat_row["DAT_MOVEMENT_DATETIME"])
    stream_datetime = _timestamp(stream_row["STREAM_MOVEMENT_DATETIME"])
    stream_status = normalize_code(stream_row["STREAM STATUS FLIGHT"])
    invalid_status = bool(
        treat_invalid_status_as_missing and stream_status in invalid_statuses
    )
    time_difference: float | None = None
    date_mismatch = False
    if dat_datetime is not None and stream_datetime is not None:
        time_difference = abs(
            (stream_datetime - dat_datetime).total_seconds() / 60.0
        )
        date_mismatch = dat_datetime.date() != stream_datetime.date()

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
        time_difference if time_difference is not None else float("inf"),
        priority_value,
        -parse_message_number(stream_row["MESSAGE NUM"]),
        -int(stream_row["SOURCE ROW"]),
    )
    validation_result = "VALID STREAM CANDIDATE"
    if invalid_reasons:
        validation_result = " / ".join(invalid_reasons)
    return {
        "stream_row": stream_row,
        "stream_status": stream_status,
        "dat_datetime": dat_datetime,
        "stream_datetime": stream_datetime,
        "time_difference": time_difference,
        "date_mismatch": date_mismatch,
        "invalid_reasons": invalid_reasons,
        "validation_result": validation_result,
        "selection_sort": selection_sort,
    }


def _result_from_dat(
    dat_row: pd.Series,
    status: str,
    reason: str,
    selected_evaluation: Mapping[str, object] | None,
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
    result.update(
        {
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
            "BILLING CATEGORY": (
                "NON-BILLABLE/INTERNAL REVIEW"
                if is_non_billable_flight(dat_row["FLIGHT NUMBER"])
                else "BILLING REVIEW"
            ),
        }
    )
    if isinstance(stream_row, pd.Series):
        result["STREAM SOURCE ROW"] = int(stream_row["SOURCE ROW"])
    return result


def _result_from_stream(stream_row: pd.Series) -> dict[str, object]:
    result = {column: stream_row.get(column, "") for column in RESULT_COLUMNS}
    result.update(
        {
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
            "BILLING CATEGORY": "",
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
    dep = standardize_dataset(dat_dep, dep_mapping, "DAT DEP", "D")
    arr = standardize_dataset(dat_arr, arr_mapping, "DAT ARR", "A")
    stream_standard = standardize_dataset(stream, stream_mapping, "STREAM")
    if dep.empty:
        dat_combined = arr.copy().reset_index(drop=True)
    elif arr.empty:
        dat_combined = dep.copy().reset_index(drop=True)
    else:
        dat_combined = pd.DataFrame.from_records(
            dep.to_dict("records") + arr.to_dict("records"),
            columns=STANDARDIZED_COLUMNS,
        )

    dat_prepared = _prepare_keys(dat_combined, "DAT")
    stream_prepared = _prepare_keys(stream_standard, "STREAM")
    dat_unique, duplicate_raw = deduplicate_dat(dat_prepared)
    duplicate_dat = _duplicate_output(duplicate_raw)

    stream_groups = {
        str(key): group.copy()
        for key, group in stream_prepared.loc[
            stream_prepared["__complete_key"]
        ].groupby("__match_key", sort=False)
    }
    dat_keys = set(dat_unique.loc[dat_unique["__complete_key"], "__match_key"])

    matched_records: list[dict[str, object]] = []
    missing_records: list[dict[str, object]] = []
    need_review_records: list[dict[str, object]] = []
    audit_records: list[dict[str, object]] = []

    for _, dat_row in dat_unique.iterrows():
        candidates = stream_groups.get(str(dat_row["__match_key"]))
        if candidates is None or candidates.empty:
            result = _result_from_dat(
                dat_row, "MISSING IN STREAM", "STREAM NOT FOUND", None
            )
            missing_records.append(result)
            audit_records.append(result.copy())
            continue

        evaluations = [
            _candidate_evaluation(
                dat_row,
                stream_row,
                invalid_statuses,
                treat_invalid_stream_status_as_missing,
            )
            for _, stream_row in candidates.iterrows()
        ]
        selected = min(evaluations, key=lambda item: item["selection_sort"])
        invalid_reasons = list(selected["invalid_reasons"])
        difference = selected["time_difference"]
        if invalid_reasons:
            status = "MISSING IN STREAM"
            reason = " / ".join(invalid_reasons)
        elif difference is None:
            status = "MISSING IN STREAM"
            reason = "STREAM INVALID MOVEMENT TIME"
        elif float(difference) <= time_tolerance_minutes:
            status = "MATCHED"
            reason = "VALID STREAM MATCH"
        elif float(difference) <= 120:
            status = "NEED REVIEW"
            reason = "STREAM TIME DIFFERENCE"
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
            )
            audit_row["STREAM VALIDATION RESULT"] = (
                "NOT SELECTED / " + str(evaluation["validation_result"])
            )
            audit_records.append(audit_row)

    matched = _records_frame(matched_records)
    missing = _records_frame(missing_records)
    need_review = _records_frame(need_review_records)

    extra_rows = stream_prepared.loc[
        (~stream_prepared["__complete_key"])
        | (~stream_prepared["__match_key"].isin(dat_keys))
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
                "DAT MOVEMENT DATETIME": "",
                "STREAM MOVEMENT DATETIME": "",
                "TIME DIFFERENCE MINUTES": "",
                "DUPLICATE GROUP KEY": duplicate_row.get(
                    "DUPLICATE_GROUP_KEY", ""
                ),
                "SELECTED DAT RECORD": False,
                "STREAM MATCH KEY": "",
                "STREAM VALIDATION RESULT": "NOT APPLICABLE",
                "BILLING CATEGORY": (
                    "NON-BILLABLE/INTERNAL REVIEW"
                    if is_non_billable_flight(duplicate_row.get("FLIGHT NUMBER", ""))
                    else "BILLING REVIEW"
                ),
                "COMPLETENESS_SCORE": duplicate_row.get("COMPLETENESS_SCORE", ""),
                "HAS_MOVEMENT_TIME": duplicate_row.get("HAS_MOVEMENT_TIME", ""),
            }
        )
        audit_records.append(audit_record)
    audit_detail = _records_frame(audit_records)

    missing_billing = missing.loc[
        missing["BILLING CATEGORY"] == "BILLING REVIEW"
    ].reset_index(drop=True)
    missing_non_billable = missing.loc[
        missing["BILLING CATEGORY"] == "NON-BILLABLE/INTERNAL REVIEW"
    ].reset_index(drop=True)

    unique_dat_total = len(dat_unique)
    matched_total = len(matched)
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
        "extra_in_stream": len(extra),
        "duplicate_dat": len(duplicate_dat),
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
            ("Missing Non-Billable/Internal Review", summary["missing_non_billable_review"]),
            ("Need Review", summary["need_review"]),
            ("Extra in Stream", summary["extra_in_stream"]),
            ("Duplicate DAT", summary["duplicate_dat"]),
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
        "matched": matched,
        "need_review": need_review,
        "extra": extra,
        "duplicates": duplicate_dat,
        "audit_detail": audit_detail,
        "dat_unique": dat_unique,
        "dat_dep": dep,
        "dat_arr": arr,
        "stream": stream_standard,
        "settings": {
            "time_tolerance_minutes": time_tolerance_minutes,
            "invalid_stream_statuses": sorted(invalid_statuses),
            "treat_invalid_stream_status_as_missing": treat_invalid_stream_status_as_missing,
        },
    }
