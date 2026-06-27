from __future__ import annotations

import csv
import io
import re
import unicodedata
import zipfile
from collections import OrderedDict
from datetime import date, datetime, time, timedelta
from html.parser import HTMLParser
from typing import Iterable, Mapping

import pandas as pd


FIELD_DEFINITIONS = OrderedDict(
    [
        (
            "flight",
            {
                "label": "Flight Number / Callsign",
                "aliases": [
                    "flight number",
                    "flight no",
                    "flight",
                    "flightnumber",
                    "flightnum",
                    "flight nr",
                    "flt no",
                    "flt",
                    "callsign",
                    "call sign",
                    "acid",
                ],
            },
        ),
        (
            "adep",
            {
                "label": "Aerodrome / ADEP / From",
                "aliases": [
                    "aerodrome",
                    "adep",
                    "from",
                    "origin",
                    "departure aerodrome",
                    "departure airport",
                    "airport from",
                ],
            },
        ),
        (
            "ades",
            {
                "label": "To / ADES",
                "aliases": [
                    "to",
                    "to from",
                    "to/from",
                    "ades",
                    "destination",
                    "arrival aerodrome",
                    "arrival airport",
                    "airport to",
                ],
            },
        ),
        (
            "eobd",
            {
                "label": "EOBD / Date of Flight",
                "aliases": [
                    "eobd",
                    "date of flight",
                    "flight date",
                    "date flight",
                    "dof",
                    "date",
                    "tanggal penerbangan",
                    "tanggal",
                ],
            },
        ),
        (
            "eobt",
            {
                "label": "EOBT",
                "aliases": [
                    "eobt",
                    "estimated off block time",
                    "estimated departure time",
                    "schedule time",
                    "scheduled time",
                    "std",
                ],
            },
        ),
        (
            "atd",
            {
                "label": "ATD",
                "aliases": [
                    "atd",
                    "actual time departure",
                    "actual departure time",
                    "off block time",
                ],
            },
        ),
        (
            "ata",
            {
                "label": "ATA",
                "aliases": [
                    "ata",
                    "actual time arrival",
                    "actual arrival time",
                    "on block time",
                ],
            },
        ),
        (
            "register",
            {
                "label": "Register",
                "aliases": [
                    "register",
                    "registration",
                    "aircraft registration",
                    "aircraft register",
                    "ac register",
                    "reg",
                    "tail number",
                    "tail no",
                ],
            },
        ),
        (
            "movement",
            {
                "label": "Movement Type (D/A/L/O)",
                "aliases": [
                    "d/a/l/o",
                    "d a l o",
                    "movement",
                    "movement type",
                    "movement indicator",
                ],
            },
        ),
        (
            "arrival_gate",
            {
                "label": "Arrival Gate",
                "aliases": ["arrival gate", "arrivalgate", "arr gate", "gate arrival"],
            },
        ),
        (
            "departure_gate",
            {
                "label": "Departure Gate",
                "aliases": [
                    "departure gate",
                    "departuregate",
                    "dep gate",
                    "gate departure",
                ],
            },
        ),
        (
            "arrival_runway",
            {
                "label": "Arrival Runway",
                "aliases": [
                    "arrival runway",
                    "arrivalrunway",
                    "arr runway",
                    "landing runway",
                ],
            },
        ),
        (
            "departure_runway",
            {
                "label": "Departure Runway",
                "aliases": [
                    "departure runway",
                    "departurerunway",
                    "dep runway",
                    "takeoff runway",
                ],
            },
        ),
        (
            "parking",
            {
                "label": "Parking / Gate",
                "aliases": ["parking", "parking stand", "stand", "gate"],
            },
        ),
        (
            "runway",
            {
                "label": "Runway",
                "aliases": [
                    "runway",
                    "rwy",
                    "runway used",
                    "departure runway",
                    "arrival runway",
                ],
            },
        ),
    ]
)

FIELD_KEYS = list(FIELD_DEFINITIONS)
OUTPUT_LABELS = {key: value["label"] for key, value in FIELD_DEFINITIONS.items()}
TIME_FIELDS = {"eobt", "atd", "ata"}
CODE_FIELDS = {
    "flight",
    "adep",
    "ades",
    "register",
    "movement",
    "runway",
    "arrival_runway",
    "departure_runway",
}
UNMAPPED = "— Tidak dipetakan —"


def clean_column_name(value: object) -> str:
    """Return a stable display name and make duplicate/blank headers usable."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "Kolom Tanpa Nama"
    text = str(value).strip()
    return text or "Kolom Tanpa Nama"


def make_unique_columns(columns: Iterable[object]) -> list[str]:
    seen: dict[str, int] = {}
    result: list[str] = []
    for raw in columns:
        base = clean_column_name(raw)
        seen[base] = seen.get(base, 0) + 1
        result.append(base if seen[base] == 1 else f"{base} ({seen[base]})")
    return result


def normalize_header(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", "", text.lower())


def detect_mapping(columns: Iterable[object]) -> dict[str, str | None]:
    """Best-effort alias detection with exact matches preferred over partial ones."""
    display_columns = [str(column) for column in columns]
    normalized_columns = {column: normalize_header(column) for column in display_columns}
    mapping: dict[str, str | None] = {}
    used: set[str] = set()

    for field, definition in FIELD_DEFINITIONS.items():
        aliases = [normalize_header(alias) for alias in definition["aliases"]]
        candidates: list[tuple[int, int, str]] = []
        for column, normalized in normalized_columns.items():
            if column in used or not normalized:
                continue
            score = 0
            if normalized in aliases:
                score = 100
            else:
                for alias in aliases:
                    # Partial detection is deliberately conservative for short
                    # aliases such as "to", "reg", or "date".
                    if len(alias) >= 5 and (alias in normalized or normalized in alias):
                        score = max(score, 60 + min(len(alias), len(normalized)))
            if score:
                candidates.append((score, -display_columns.index(column), column))

        if candidates:
            column = max(candidates)[2]
            mapping[field] = column
            used.add(column)
        else:
            mapping[field] = None
    return mapping


def read_source_csv(data: bytes) -> tuple[pd.DataFrame, str]:
    """Read CSV with common Indonesian/Windows encodings and separator detection."""
    errors: list[str] = []
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            text = data.decode(encoding)
            sample = text[:8192]
            try:
                delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
            except csv.Error:
                delimiter = ","
            frame = pd.read_csv(
                io.StringIO(text),
                sep=delimiter,
                dtype=object,
                keep_default_na=False,
            )
            if frame.shape[1] == 1 and delimiter != ";" and ";" in sample:
                frame = pd.read_csv(
                    io.StringIO(text), sep=";", dtype=object, keep_default_na=False
                )
            frame.columns = make_unique_columns(frame.columns)
            return frame, encoding
        except (UnicodeDecodeError, pd.errors.ParserError, ValueError) as exc:
            errors.append(f"{encoding}: {exc}")
    raise ValueError("CSV tidak dapat dibaca. " + " | ".join(errors[-2:]))


class _HTMLTableExtractor(HTMLParser):
    """Small dependency-free parser for reports exported as HTML with an .xls suffix."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tables: list[dict[str, list]] = []
        self._table_depth = 0
        self._table: dict[str, list] | None = None
        self._row: list[str] | None = None
        self._row_has_header = False
        self._cell_parts: list[str] | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "table":
            if self._table_depth == 0:
                self._table = {"headers": [], "rows": []}
            self._table_depth += 1
        elif self._table_depth == 1 and tag == "tr":
            self._row = []
            self._row_has_header = False
        elif self._table_depth == 1 and self._row is not None and tag in {"th", "td"}:
            self._cell_parts = []
            if tag == "th":
                self._row_has_header = True
        elif self._cell_parts is not None and tag == "br":
            self._cell_parts.append(" ")

    def handle_data(self, data: str) -> None:
        if self._cell_parts is not None:
            self._cell_parts.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if self._table_depth == 1 and tag in {"th", "td"} and self._cell_parts is not None:
            value = re.sub(r"\s+", " ", "".join(self._cell_parts)).strip()
            if self._row is not None:
                self._row.append(value)
            self._cell_parts = None
        elif self._table_depth == 1 and tag == "tr":
            if self._table is not None and self._row:
                key = "headers" if self._row_has_header else "rows"
                self._table[key].append(self._row)
            self._row = None
            self._row_has_header = False
        elif tag == "table" and self._table_depth:
            self._table_depth -= 1
            if self._table_depth == 0 and self._table is not None:
                self.tables.append(self._table)
                self._table = None


def actual_file_format(data: bytes) -> str:
    """Detect the physical format; report exports are often HTML named .xls."""
    if data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return "xls"
    if data.startswith(b"PK\x03\x04"):
        return "xlsx"
    head = data[:65536].lstrip().lower()
    if b"<table" in head or head.startswith((b"<!doctype html", b"<html")):
        return "html"
    return "unknown"


def actual_file_format_label(data: bytes) -> str:
    labels = {
        "xls": "Excel 97–2003 (XLS)",
        "xlsx": "Excel Workbook (XLSX)",
        "html": "Tabel HTML berformat .xls",
        "unknown": "Format tidak dikenal",
    }
    return labels[actual_file_format(data)]


def _decode_html_report(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("Teks HTML di dalam file tidak dapat didekode.")


def read_html_report(data: bytes) -> pd.DataFrame:
    parser = _HTMLTableExtractor()
    parser.feed(_decode_html_report(data))
    parser.close()

    candidates = [table for table in parser.tables if table["headers"]]
    if not candidates:
        raise ValueError("Tabel data dengan header tidak ditemukan di dalam file HTML.")

    table = max(
        candidates,
        key=lambda item: (
            len(item["rows"]),
            max((len(row) for row in item["headers"]), default=0),
        ),
    )
    headers = max(table["headers"], key=len)
    if not headers:
        raise ValueError("Header tabel pada file HTML kosong.")

    width = len(headers)
    rows: list[list[str]] = []
    for raw_row in table["rows"]:
        row = raw_row[:width] + [""] * max(0, width - len(raw_row))
        if any(value.strip() for value in row):
            rows.append(row)

    frame = pd.DataFrame(rows, columns=make_unique_columns(headers))
    return frame


def excel_sheet_names(data: bytes) -> list[str]:
    kind = actual_file_format(data)
    if kind == "html":
        return ["Tabel Data"]
    if kind not in {"xls", "xlsx"}:
        raise ValueError(
            "Format file aktual tidak dikenali. Gunakan XLS, XLSX, atau laporan HTML berekstensi XLS."
        )
    engine = "xlrd" if kind == "xls" else "openpyxl"
    with pd.ExcelFile(io.BytesIO(data), engine=engine) as workbook:
        return list(workbook.sheet_names)


def read_actual_excel(data: bytes, sheet_name: str | int = 0) -> pd.DataFrame:
    kind = actual_file_format(data)
    if kind == "html":
        return read_html_report(data)
    if kind not in {"xls", "xlsx"}:
        raise ValueError(
            "Format file aktual tidak dikenali. Gunakan XLS, XLSX, atau laporan HTML berekstensi XLS."
        )
    engine = "xlrd" if kind == "xls" else "openpyxl"
    frame = pd.read_excel(
        io.BytesIO(data),
        sheet_name=sheet_name,
        dtype=object,
        keep_default_na=False,
        engine=engine,
    )
    frame.columns = make_unique_columns(frame.columns)
    return frame


def clean_scalar_text(value: object) -> str:
    text = str(value).strip()
    while len(text) >= 2 and (text[0], text[-1]) in {("'", "'"), ('"', '"')}:
        text = text[1:-1].strip()
    return text


def _is_empty(value: object) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return clean_scalar_text(value) == ""


def display_value(value: object, field: str) -> str:
    if _is_empty(value):
        return ""
    raw_text = clean_scalar_text(value)
    if raw_text.upper() in {"-", "—", "N/A", "NA", "NULL", "NONE"}:
        return ""
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}\s*[-—]", raw_text):
        return ""
    if field == "eobd":
        normalized = normalize_value(value, field)
        return normalized or raw_text
    if field in TIME_FIELDS:
        normalized = normalize_value(value, field)
        return normalized or raw_text
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return raw_text


def _normalize_date(value: object) -> str:
    if _is_empty(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        if 1 <= number <= 100000:
            try:
                return (datetime(1899, 12, 30) + timedelta(days=number)).strftime(
                    "%Y-%m-%d"
                )
            except (OverflowError, ValueError):
                pass

    text = clean_scalar_text(value)
    compact_short_date = re.fullmatch(r"\d{6}", text)
    if compact_short_date:
        try:
            return datetime.strptime(text, "%y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            pass
    compact_date = re.fullmatch(r"\d{8}", text)
    if compact_date:
        try:
            if text.startswith(("19", "20")):
                parsed_compact = datetime.strptime(text, "%Y%m%d")
            else:
                parsed_compact = datetime.strptime(text, "%d%m%Y")
            return parsed_compact.strftime("%Y-%m-%d")
        except ValueError:
            pass
    iso_match = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?", text)
    if iso_match:
        try:
            return datetime(
                int(iso_match.group(1)),
                int(iso_match.group(2)),
                int(iso_match.group(3)),
            ).strftime("%Y-%m-%d")
        except ValueError:
            pass
    if re.fullmatch(r"\d{5}(?:\.0+)?", text):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(text))).strftime(
                "%Y-%m-%d"
            )
        except (OverflowError, ValueError):
            pass
    try:
        parsed = pd.to_datetime(text, errors="raise", dayfirst=True)
        return parsed.strftime("%Y-%m-%d")
    except (ValueError, TypeError, OverflowError):
        return re.sub(r"\s+", " ", text).upper()


def _normalize_time(value: object) -> str:
    if _is_empty(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%H:%M")
    if isinstance(value, datetime):
        return value.strftime("%H:%M")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    if isinstance(value, timedelta):
        minutes = int(value.total_seconds() // 60) % (24 * 60)
        return f"{minutes // 60:02d}:{minutes % 60:02d}"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        if 0 <= number < 1:
            minutes = int(round(number * 24 * 60)) % (24 * 60)
            return f"{minutes // 60:02d}:{minutes % 60:02d}"
        if 0 <= number <= 2359 and float(number).is_integer():
            digits = f"{int(number):04d}"
            hour, minute = int(digits[:-2]), int(digits[-2:])
            if hour < 24 and minute < 60:
                return f"{hour:02d}:{minute:02d}"

    text = clean_scalar_text(value).upper()
    embedded_time = re.search(r"(?:^|\s)(\d{1,2}):(\d{2})(?::\d{2})?(?:\s|$)", text)
    if embedded_time:
        hour, minute = int(embedded_time.group(1)), int(embedded_time.group(2))
        if hour < 24 and minute < 60:
            return f"{hour:02d}:{minute:02d}"
    text = text.replace(".", ":")
    if re.fullmatch(r"\d{1,4}(?:\.0+)?", text):
        digits = text.split(".")[0].zfill(4)
        hour, minute = int(digits[:-2]), int(digits[-2:])
        if hour < 24 and minute < 60:
            return f"{hour:02d}:{minute:02d}"
    match = re.match(r"^(\d{1,2}):(\d{2})", text)
    if match:
        hour, minute = int(match.group(1)), int(match.group(2))
        if hour < 24 and minute < 60:
            return f"{hour:02d}:{minute:02d}"
    return re.sub(r"\s+", " ", text)


def normalize_value(value: object, field: str) -> str:
    if _is_empty(value):
        return ""
    if clean_scalar_text(value).upper() in {"-", "—", "N/A", "NA", "NULL", "NONE"}:
        return ""
    if field == "eobd":
        return _normalize_date(value)
    if field in TIME_FIELDS:
        return _normalize_time(value)

    text = unicodedata.normalize("NFKC", clean_scalar_text(value)).strip().upper()
    if field in CODE_FIELDS:
        return re.sub(r"[^A-Z0-9]+", "", text)
    return re.sub(r"\s+", " ", text)


def canonicalize(
    frame: pd.DataFrame, mapping: Mapping[str, str | None]
) -> pd.DataFrame:
    result = pd.DataFrame(index=frame.index)
    for field in FIELD_KEYS:
        column = mapping.get(field)
        if column and column in frame.columns:
            result[field] = frame[column].map(lambda value, f=field: display_value(value, f))
        else:
            result[field] = ""
    result["__row"] = range(2, len(result) + 2)
    return result.reset_index(drop=True)


def _prepare_for_matching(
    frame: pd.DataFrame,
    key_fields: list[str],
    side: str,
    allow_incomplete_keys: bool,
) -> pd.DataFrame:
    prepared = frame.copy()
    normalized_columns: list[str] = []
    for field in key_fields:
        column = f"__key_{field}"
        prepared[column] = prepared[field].map(lambda value, f=field: normalize_value(value, f))
        normalized_columns.append(column)

    prepared["__complete_key"] = prepared[normalized_columns].ne("").all(axis=1)
    if allow_incomplete_keys:
        group_columns = normalized_columns
        prepared["__occurrence"] = prepared.groupby(
            group_columns, dropna=False, sort=False
        ).cumcount()
        prepared["__match_id"] = (
            prepared[normalized_columns].astype(str).agg("␟".join, axis=1)
            + "␟"
            + prepared["__occurrence"].astype(str)
        )
    else:
        complete = prepared["__complete_key"]
        prepared["__occurrence"] = 0
        prepared.loc[complete, "__occurrence"] = prepared.loc[complete].groupby(
            normalized_columns, dropna=False, sort=False
        ).cumcount()
        prepared.loc[complete, "__match_id"] = (
            prepared.loc[complete, normalized_columns].astype(str).agg("␟".join, axis=1)
            + "␟"
            + prepared.loc[complete, "__occurrence"].astype(str)
        )
        prepared.loc[~complete, "__match_id"] = prepared.loc[~complete, "__row"].map(
            lambda row: f"__INCOMPLETE__{side}__{row}"
        )
    return prepared


def _unmatched_output(
    rows: pd.DataFrame, origin: str, note: str
) -> pd.DataFrame:
    data: dict[str, object] = {
        "Keterangan": [note] * len(rows),
        "Asal Data": [origin] * len(rows),
        "Baris Asli": rows["__row"].astype(int).tolist() if len(rows) else [],
    }
    for field in FIELD_KEYS:
        data[OUTPUT_LABELS[field]] = rows[field].tolist() if len(rows) else []
    return pd.DataFrame(data)


def _matched_output(
    matches: pd.DataFrame, comparable_fields: list[str]
) -> pd.DataFrame:
    records: list[dict[str, object]] = []
    for _, row in matches.iterrows():
        different: list[str] = []
        for field in comparable_fields:
            source_value = normalize_value(row[f"source_{field}"], field)
            actual_value = normalize_value(row[f"actual_{field}"], field)
            if source_value != actual_value:
                different.append(OUTPUT_LABELS[field])

        record: dict[str, object] = {
            "Status Perbandingan": "Ada Perbedaan" if different else "Sama",
            "Field Berbeda": ", ".join(different),
            "Baris Sumber": int(row["source___row"]),
            "Baris Aktual": int(row["actual___row"]),
        }
        for field in FIELD_KEYS:
            label = OUTPUT_LABELS[field]
            record[f"{label} (Sumber)"] = row[f"source_{field}"]
            record[f"{label} (Aktual)"] = row[f"actual_{field}"]
        records.append(record)
    return pd.DataFrame(records)


def compare_datasets(
    source_frame: pd.DataFrame,
    actual_frame: pd.DataFrame,
    source_mapping: Mapping[str, str | None],
    actual_mapping: Mapping[str, str | None],
    key_fields: Iterable[str],
    allow_incomplete_keys: bool = False,
) -> dict[str, object]:
    keys = list(dict.fromkeys(key_fields))
    if not keys:
        raise ValueError("Pilih minimal satu kunci pencocokan.")
    invalid = [field for field in keys if field not in FIELD_DEFINITIONS]
    if invalid:
        raise ValueError(f"Kunci tidak dikenal: {', '.join(invalid)}")
    for field in keys:
        if not source_mapping.get(field) or not actual_mapping.get(field):
            raise ValueError(
                f"Kolom {OUTPUT_LABELS[field]} harus dipetakan pada kedua data."
            )

    source = canonicalize(source_frame, source_mapping)
    actual = canonicalize(actual_frame, actual_mapping)
    source = _prepare_for_matching(source, keys, "SOURCE", allow_incomplete_keys)
    actual = _prepare_for_matching(actual, keys, "ACTUAL", allow_incomplete_keys)

    source_ids = set(source["__match_id"])
    actual_ids = set(actual["__match_id"])
    shared_ids = source_ids & actual_ids

    source_only_raw = source.loc[~source["__match_id"].isin(actual_ids)].copy()
    actual_only_raw = actual.loc[~actual["__match_id"].isin(source_ids)].copy()
    source_match = source.loc[source["__match_id"].isin(shared_ids)].copy()
    actual_match = actual.loc[actual["__match_id"].isin(shared_ids)].copy()

    matched_raw = source_match.merge(
        actual_match,
        on="__match_id",
        how="inner",
        suffixes=("_source", "_actual"),
        validate="one_to_one",
    )
    rename: dict[str, str] = {}
    for field in FIELD_KEYS:
        rename[f"{field}_source"] = f"source_{field}"
        rename[f"{field}_actual"] = f"actual_{field}"
    rename["__row_source"] = "source___row"
    rename["__row_actual"] = "actual___row"
    matched_raw = matched_raw.rename(columns=rename)

    comparable = [
        field
        for field in FIELD_KEYS
        if field not in keys
        and source_mapping.get(field)
        and actual_mapping.get(field)
    ]
    matched = _matched_output(matched_raw, comparable)
    if matched.empty:
        matched = pd.DataFrame(
            columns=["Status Perbandingan", "Field Berbeda", "Baris Sumber", "Baris Aktual"]
            + [
                f"{OUTPUT_LABELS[field]} ({side})"
                for field in FIELD_KEYS
                for side in ("Sumber", "Aktual")
            ]
        )

    source_only = _unmatched_output(
        source_only_raw,
        "Data Sumber",
        "Hanya ada di Data Sumber; tidak ditemukan di Data Aktual",
    )
    actual_only = _unmatched_output(
        actual_only_raw,
        "Data Aktual",
        "Hanya ada di Data Aktual; tidak ditemukan di Data Sumber",
    )
    combined_unmatched = pd.concat([source_only, actual_only], ignore_index=True)
    differences = matched.loc[matched["Status Perbandingan"] == "Ada Perbedaan"].copy()

    source_incomplete = int((~source["__complete_key"]).sum())
    actual_incomplete = int((~actual["__complete_key"]).sum())
    summary = {
        "source_total": len(source),
        "actual_total": len(actual),
        "matched_total": len(matched),
        "source_only_total": len(source_only),
        "actual_only_total": len(actual_only),
        "differences_total": len(differences),
        "source_incomplete_keys": source_incomplete,
        "actual_incomplete_keys": actual_incomplete,
    }
    return {
        "summary": summary,
        "source_only": source_only,
        "actual_only": actual_only,
        "unmatched": combined_unmatched,
        "matched": matched,
        "differences": differences,
    }


def dataframe_to_csv_bytes(frame: pd.DataFrame) -> bytes:
    # UTF-8 BOM makes Indonesian text and punctuation open cleanly in Excel.
    return frame.to_csv(index=False).encode("utf-8-sig")


def results_to_zip(results: Mapping[str, object]) -> bytes:
    files = {
        "01_hanya_data_sumber.csv": results["source_only"],
        "02_hanya_data_aktual.csv": results["actual_only"],
        "03_semua_tidak_cocok.csv": results["unmatched"],
        "04_data_cocok.csv": results["matched"],
        "05_perbedaan_nilai.csv": results["differences"],
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        for filename, frame in files.items():
            assert isinstance(frame, pd.DataFrame)
            archive.writestr(filename, dataframe_to_csv_bytes(frame))
    return buffer.getvalue()
