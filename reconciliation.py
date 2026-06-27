"""Backward-compatible public API for FINDER's reconciliation modules."""

from utils.exporter import MISSING_EXPORT_COLUMNS, build_excel_report
from utils.reconciliation import (
    AUDIT_COLUMNS,
    DEFAULT_INVALID_STREAM_STATUSES,
    DETAIL_COLUMNS,
    DUPLICATE_AUDIT_COLUMNS,
    MATCH_KEY_SPECS,
    RESULT_COLUMNS,
    deduplicate_dat,
    detected_mapping,
    read_uploaded_table,
    reconcile_dat_vs_stream,
    standardize_dataset,
    validate_required_columns,
)

__all__ = [
    "AUDIT_COLUMNS",
    "DEFAULT_INVALID_STREAM_STATUSES",
    "DETAIL_COLUMNS",
    "DUPLICATE_AUDIT_COLUMNS",
    "MATCH_KEY_SPECS",
    "MISSING_EXPORT_COLUMNS",
    "RESULT_COLUMNS",
    "build_excel_report",
    "deduplicate_dat",
    "detected_mapping",
    "read_uploaded_table",
    "reconcile_dat_vs_stream",
    "standardize_dataset",
    "validate_required_columns",
]
