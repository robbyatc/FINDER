from __future__ import annotations

import hashlib
import html
from datetime import date

import pandas as pd
import streamlit as st

from comparison import OUTPUT_LABELS
from reconciliation import (
    AUDIT_COLUMNS,
    DISPLAY_RESULT_COLUMNS,
    DUPLICATE_AUDIT_COLUMNS,
    build_excel_report,
    detected_mapping,
    read_uploaded_table,
    reconcile_dat_vs_stream,
    validate_required_columns,
)


st.set_page_config(
    page_title="FINDER — DAT vs STREAM Reconciliation System",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# Clear reconciliation output created with an older result schema. Without this,
# a long-lived Streamlit session can keep displaying a cached table that predates
# newly added result columns even after the application has been redeployed.
RESULT_SCHEMA_VERSION = "2026-06-30-actual-movement-validation-v5"
if st.session_state.get("finder_result_schema_version") != RESULT_SCHEMA_VERSION:
    for stale_key in (
        "finder_results",
        "finder_report",
        "finder_result_signature",
    ):
        st.session_state.pop(stale_key, None)
    st.session_state["finder_result_schema_version"] = RESULT_SCHEMA_VERSION


st.markdown(
    """
    <style>
      :root {
        --navy-950:#07111f; --navy-900:#0b1930; --navy-800:#102744;
        --blue-600:#1769e0; --blue-500:#2f80ed; --blue-100:#eaf2ff;
        --slate-900:#172033; --slate-600:#5f6d83; --slate-300:#d7dfec;
        --green:#1f9d68; --red:#e84c3d; --orange:#f28b30; --purple:#7567c7;
        --yellow:#d99a1b; --surface:#ffffff; --canvas:#f4f7fb;
      }
      .stApp { background:var(--canvas); }
      .block-container { max-width:1540px; padding-top:1.3rem; padding-bottom:4rem; }
      [data-testid="stHeader"] { background:transparent; }
      [data-testid="stAppDeployButton"], footer { display:none; }
      [data-testid="stSidebar"] {
        background:linear-gradient(180deg,var(--navy-950),var(--navy-900));
        border-right:1px solid rgba(93,156,255,.16);
      }
      [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
      [data-testid="stSidebar"] label,
      [data-testid="stSidebar"] .stCaption { color:#c7d7ef !important; }
      [data-testid="stSidebar"] hr { border-color:rgba(255,255,255,.12); }
      [data-testid="stSidebar"] div[role="radiogroup"] label {
        border-radius:11px; padding:8px 10px; margin:3px 0;
        transition:all .15s ease;
      }
      [data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        background:rgba(47,128,237,.14);
      }
      .sidebar-brand { display:flex; gap:12px; align-items:center; padding:6px 2px 18px; }
      .sidebar-logo {
        width:48px; height:48px; border-radius:14px; display:grid; place-items:center;
        background:linear-gradient(145deg,#1769e0,#44a5ff); color:white; font-size:1.45rem;
        box-shadow:0 10px 28px rgba(23,105,224,.35); border:1px solid rgba(255,255,255,.24);
      }
      .sidebar-title { color:#fff; font-weight:800; letter-spacing:.08em; font-size:1.1rem; }
      .sidebar-subtitle { color:#87a6cf; font-size:.73rem; margin-top:2px; line-height:1.35; }
      .sidebar-status { margin-top:18px; padding:12px; border-radius:12px; color:#b9cbe4;
        background:rgba(255,255,255,.045); border:1px solid rgba(255,255,255,.08); font-size:.8rem; }

      .finder-hero {
        position:relative; overflow:hidden; border-radius:22px; padding:29px 34px;
        background:linear-gradient(125deg,#07111f 0%,#0c2343 62%,#103a68 100%);
        box-shadow:0 18px 42px rgba(7,17,31,.18); border:1px solid rgba(93,156,255,.24);
        margin-bottom:24px;
      }
      .finder-hero:after {
        content:""; position:absolute; width:260px; height:260px; right:-65px; top:-92px;
        border-radius:50%; border:1px solid rgba(80,167,255,.26);
        box-shadow:0 0 0 38px rgba(80,167,255,.05),0 0 0 78px rgba(80,167,255,.035);
      }
      .hero-kicker { color:#6eb6ff; font-size:.73rem; font-weight:800; letter-spacing:.16em; }
      .hero-title { color:#fff; font-size:2.55rem; line-height:1; font-weight:850;
        letter-spacing:.04em; margin:8px 0 6px; }
      .hero-subtitle { color:#cce1ff; font-size:1.06rem; font-weight:650; }
      .hero-description { color:#91a9c8; max-width:850px; margin-top:9px; font-size:.92rem; }
      .hero-chip { position:absolute; right:34px; bottom:27px; z-index:2; color:#a9d2ff;
        border:1px solid rgba(110,182,255,.25); background:rgba(20,71,120,.32);
        border-radius:999px; padding:7px 12px; font-size:.72rem; letter-spacing:.06em; }

      .section-eyebrow { color:var(--blue-600); font-size:.71rem; font-weight:800; letter-spacing:.13em; }
      .section-title { color:var(--slate-900); font-size:1.45rem; font-weight:800; margin:3px 0 2px; }
      .section-copy { color:var(--slate-600); font-size:.88rem; margin-bottom:14px; }
      .upload-heading { display:flex; align-items:center; justify-content:space-between; gap:8px; }
      .upload-icon { width:35px; height:35px; border-radius:10px; display:grid; place-items:center;
        color:white; background:linear-gradient(145deg,#1769e0,#4da3ff); }
      .upload-name { color:var(--slate-900); font-weight:800; font-size:.96rem; }
      .upload-hint { color:#7a879b; font-size:.75rem; margin-top:2px; }
      div[data-testid="stVerticalBlockBorderWrapper"] {
        background:var(--surface); border-color:#e0e7f1 !important; border-radius:16px !important;
        box-shadow:0 8px 25px rgba(17,39,68,.055); padding:4px;
      }
      div[data-testid="stFileUploaderDropzone"] {
        border:1px dashed #b8c9df; background:#f7faff; border-radius:12px;
      }
      .status-badge { display:inline-flex; align-items:center; gap:6px; border-radius:999px;
        padding:5px 10px; font-size:.7rem; font-weight:800; letter-spacing:.04em; }
      .status-ok { color:#087a4d; background:#e6f8f0; border:1px solid #bfead7; }
      .status-off { color:#6e7c91; background:#f0f3f7; border:1px solid #dbe2eb; }
      .status-danger { color:#b42318; background:#fff0ee; border:1px solid #fac9c3; }
      .file-stats { display:grid; grid-template-columns:1fr 1fr; gap:8px; margin:9px 0 1px; }
      .file-stat { padding:9px 10px; border-radius:10px; background:#f6f8fc; border:1px solid #e8edf4; }
      .file-stat-value { color:#1a2a42; font-weight:800; font-size:1rem; }
      .file-stat-label { color:#79869a; font-size:.67rem; text-transform:uppercase; letter-spacing:.06em; }
      .validation-item { display:flex; align-items:center; gap:10px; padding:9px 0;
        border-bottom:1px solid #edf1f6; }
      .validation-item:last-child { border-bottom:none; }
      .validation-dot { width:23px; height:23px; border-radius:50%; display:grid; place-items:center;
        font-size:.72rem; font-weight:900; }
      .valid { color:white; background:var(--green); }
      .invalid { color:#7a879b; background:#e9edf3; }
      .validation-label { color:#273750; font-weight:650; font-size:.87rem; }
      .validation-detail { color:#8490a3; font-size:.72rem; }
      .metric-card { min-height:116px; padding:16px 17px; background:#fff; border-radius:15px;
        border:1px solid #e0e7f1; border-top:4px solid var(--accent); box-shadow:0 8px 22px rgba(17,39,68,.055); }
      .metric-priority { box-shadow:0 12px 30px rgba(232,76,61,.16); background:linear-gradient(145deg,#fff,#fff6f3); }
      .metric-label { color:#66758b; font-size:.7rem; font-weight:750; letter-spacing:.06em; text-transform:uppercase; }
      .metric-value { color:#15243a; font-size:1.72rem; font-weight:850; margin-top:7px; line-height:1; }
      .metric-note { color:#95a0b1; font-size:.67rem; margin-top:8px; }
      .result-banner { padding:16px 18px; border-radius:14px; margin-bottom:14px;
        background:linear-gradient(90deg,#fff2ef,#fff8f4); border:1px solid #ffd2ca; }
      .result-banner-title { color:#b42318; font-weight:850; font-size:1rem; }
      .result-banner-copy { color:#8d5148; font-size:.78rem; margin-top:2px; }
      .workflow { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:15px 0; }
      .workflow-step { border-radius:14px; padding:16px; background:#fff; border:1px solid #e0e7f1; }
      .workflow-num { color:#fff; background:var(--blue-600); width:26px; height:26px;
        border-radius:8px; display:grid; place-items:center; font-weight:800; font-size:.75rem; }
      .workflow-title { color:#1d2e48; font-weight:750; margin-top:10px; font-size:.86rem; }
      .workflow-copy { color:#7e8a9d; font-size:.72rem; margin-top:4px; }
      .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
        background:linear-gradient(90deg,#1769e0,#2f80ed); border:none; color:white;
        border-radius:12px; min-height:48px; font-weight:800; letter-spacing:.025em;
        box-shadow:0 9px 22px rgba(23,105,224,.22);
      }
      .stButton > button, .stDownloadButton > button { border-radius:11px; }
      [data-testid="stMetric"] { background:white; border:1px solid #e0e7f1; border-radius:14px; padding:12px; }
      .stTabs [data-baseweb="tab-list"] { gap:5px; background:#eaf0f8; padding:5px; border-radius:12px; }
      .stTabs [data-baseweb="tab"] { border-radius:9px; padding:8px 15px; }
      .stTabs [aria-selected="true"] { background:white; box-shadow:0 2px 8px rgba(17,39,68,.08); }
      @media (max-width:900px) {
        .hero-chip { display:none; } .workflow { grid-template-columns:1fr 1fr; }
        .finder-hero { padding:24px; } .hero-title { font-size:2rem; }
      }
    </style>
    """,
    unsafe_allow_html=True,
)


def file_digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


@st.cache_data(show_spinner=False)
def cached_read_table(data: bytes, filename: str):
    return read_uploaded_table(data, filename)


def render_header() -> None:
    st.markdown(
        """
        <div class="finder-hero">
          <div class="hero-kicker">AIR TRAFFIC BILLING VALIDATION</div>
          <div class="hero-title">FINDER</div>
          <div class="hero-subtitle">DAT vs STREAM Reconciliation System</div>
          <div class="hero-description">
            Mendeteksi data penerbangan yang ada di DAT DEP dan DAT ARR tetapi belum masuk ke STREAM
            untuk validasi billing yang cepat, akurat, dan dapat diaudit.
          </div>
          <div class="hero-chip">● OPERATIONAL MONITORING</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_heading(eyebrow: str, title: str, copy: str) -> None:
    st.markdown(
        f"""
        <div class="section-eyebrow">{html.escape(eyebrow)}</div>
        <div class="section-title">{html.escape(title)}</div>
        <div class="section-copy">{html.escape(copy)}</div>
        """,
        unsafe_allow_html=True,
    )


def reconciliation_settings() -> dict[str, object]:
    tolerance_label = st.session_state.get(
        "finder_time_tolerance", "30 minutes - Recommended"
    )
    tolerance_map = {
        "15 minutes": 15,
        "30 minutes - Recommended": 30,
        "60 minutes": 60,
    }
    invalid_status_text = st.session_state.get("finder_invalid_statuses", "OTHER")
    invalid_statuses = [
        item.strip().upper()
        for item in str(invalid_status_text).replace("\n", ",").split(",")
        if item.strip()
    ]
    return {
        "time_tolerance_minutes": tolerance_map.get(str(tolerance_label), 30),
        "invalid_stream_statuses": invalid_statuses or ["OTHER"],
        "treat_invalid_stream_status_as_missing": bool(
            st.session_state.get("finder_treat_invalid_status", True)
        ),
        "show_audit_columns": bool(
            st.session_state.get("finder_show_audit_columns", False)
        ),
    }


def render_sidebar() -> str:
    with st.sidebar:
        st.markdown(
            """
            <div class="sidebar-brand">
              <div class="sidebar-logo">✈</div>
              <div>
                <div class="sidebar-title">FINDER</div>
                <div class="sidebar-subtitle">DAT vs STREAM<br>Reconciliation System</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        page = st.radio(
            "Navigation",
            [
                "📤  Upload Data",
                "📊  Reconciliation Result",
                "📥  Export Report",
                "ℹ️  About",
            ],
            label_visibility="collapsed",
            key="finder_navigation",
        )
        st.divider()
        st.markdown("#### ⚙️ Reconciliation settings")
        st.selectbox(
            "Time tolerance",
            [
                "15 minutes",
                "30 minutes - Recommended",
                "60 minutes",
            ],
            index=1,
            key="finder_time_tolerance",
        )
        st.text_input(
            "Invalid STREAM status",
            value="OTHER",
            help="Pisahkan beberapa status dengan koma, misalnya OTHER, CANCELLED.",
            key="finder_invalid_statuses",
        )
        st.checkbox(
            "Treat invalid STREAM status as Missing in Stream",
            value=True,
            key="finder_treat_invalid_status",
        )
        st.checkbox(
            "Show audit columns",
            value=False,
            key="finder_show_audit_columns",
        )
        st.divider()
        has_result = bool(st.session_state.get("finder_results"))
        status_class = "status-ok" if has_result else "status-off"
        status_text = "RESULT READY" if has_result else "WAITING FOR DATA"
        st.markdown(
            f"""
            <div class="sidebar-status">
              <span class="status-badge {status_class}">{status_text}</span><br><br>
              Upload 3 file → Process → Review Missing → Export Excel
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.caption("FINDER · Operational Reconciliation")
    return page


def render_upload_card(title: str, subtitle: str, key: str, icon: str) -> dict[str, object]:
    with st.container(border=True):
        st.markdown(
            f"""
            <div class="upload-heading">
              <div>
                <div class="upload-name">{icon} &nbsp;{html.escape(title)}</div>
                <div class="upload-hint">{html.escape(subtitle)}</div>
              </div>
              <div class="upload-icon">⇧</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        uploaded = st.file_uploader(
            title,
            type=["csv", "tsv", "xls", "xlsx"],
            key=key,
            label_visibility="collapsed",
            help="CSV, XLS, XLSX, dan laporan tabel HTML berekstensi XLS didukung.",
        )
        if uploaded is None:
            st.markdown(
                '<span class="status-badge status-off">○ NOT UPLOADED</span>',
                unsafe_allow_html=True,
            )
            st.caption("Belum ada file dipilih")
            return {"loaded": False, "error": None}

        data = uploaded.getvalue()
        try:
            with st.spinner(f"Membaca {title}…"):
                frame, format_label = cached_read_table(data, uploaded.name)
        except Exception as exc:
            st.markdown(
                '<span class="status-badge status-danger">! READ FAILED</span>',
                unsafe_allow_html=True,
            )
            st.error(str(exc))
            return {"loaded": False, "error": str(exc), "name": uploaded.name}

        st.markdown(
            '<span class="status-badge status-ok">✓ UPLOADED</span>',
            unsafe_allow_html=True,
        )
        st.caption(uploaded.name)
        st.markdown(
            f"""
            <div class="file-stats">
              <div class="file-stat"><div class="file-stat-value">{len(frame):,}</div><div class="file-stat-label">Rows</div></div>
              <div class="file-stat"><div class="file-stat-value">{len(frame.columns):,}</div><div class="file-stat-label">Columns</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.caption(format_label)
        with st.expander("Preview data", expanded=False):
            st.dataframe(frame.head(10), width="stretch", hide_index=True)
        return {
            "loaded": True,
            "error": None,
            "name": uploaded.name,
            "data": data,
            "frame": frame,
            "format": format_label,
        }


def validation_item(valid: bool, label: str, detail: str) -> str:
    icon = "✓" if valid else "•"
    css_class = "valid" if valid else "invalid"
    return f"""
      <div class="validation-item">
        <div class="validation-dot {css_class}">{icon}</div>
        <div><div class="validation-label">{html.escape(label)}</div>
        <div class="validation-detail">{html.escape(detail)}</div></div>
      </div>
    """


def render_validation(
    dep: dict[str, object], arr: dict[str, object], stream: dict[str, object]
) -> tuple[bool, dict, dict, dict, list[str]]:
    all_loaded = bool(dep.get("loaded") and arr.get("loaded") and stream.get("loaded"))
    dep_mapping = detected_mapping(dep["frame"]) if dep.get("loaded") else {}
    arr_mapping = detected_mapping(arr["frame"]) if arr.get("loaded") else {}
    stream_mapping = detected_mapping(stream["frame"]) if stream.get("loaded") else {}
    issues = (
        validate_required_columns(dep_mapping, arr_mapping, stream_mapping)
        if all_loaded
        else []
    )
    columns_valid = all_loaded and not issues
    ready = all_loaded and columns_valid

    with st.container(border=True):
        st.markdown("#### 🛡️ Validation checklist")
        validation_html = "".join(
            [
                validation_item(bool(dep.get("loaded")), "DAT DEP loaded", "File departure siap dibaca"),
                validation_item(bool(arr.get("loaded")), "DAT ARR loaded", "File arrival siap dibaca"),
                validation_item(bool(stream.get("loaded")), "STREAM loaded", "Data pembanding STREAM siap"),
                validation_item(columns_valid, "Required columns valid", "Flight, date, route, dan movement tersedia"),
                validation_item(ready, "Ready for reconciliation", "Seluruh validasi berhasil"),
            ]
        )
        st.markdown(validation_html, unsafe_allow_html=True)

    if not all_loaded:
        st.warning("Harap upload file DAT DEP, DAT ARR, dan STREAM terlebih dahulu.")
    elif issues:
        st.error("Kolom wajib belum lengkap: " + " | ".join(issues))

    if all_loaded:
        with st.expander("🔎 Detected column mapping", expanded=False):
            fields = [
                "flight", "eobd", "adep", "ades", "movement", "register",
                "eobt", "atd", "ata", "departure_gate", "arrival_gate",
                "departure_runway", "arrival_runway", "parking", "runway",
                "status_flight", "timestamp", "message_num",
            ]
            mapping_rows = []
            for field in fields:
                mapping_rows.append(
                    {
                        "Field": OUTPUT_LABELS.get(field, field),
                        "DAT DEP": "D (from file type)" if field == "movement" else dep_mapping.get(field) or "—",
                        "DAT ARR": "A (from file type)" if field == "movement" else arr_mapping.get(field) or "—",
                        "STREAM": stream_mapping.get(field) or "—",
                    }
                )
            st.dataframe(pd.DataFrame(mapping_rows), width="stretch", hide_index=True)
    return ready, dep_mapping, arr_mapping, stream_mapping, issues


def metric_card(label: str, value: str, color: str, note: str, priority: bool = False) -> None:
    priority_class = " metric-priority" if priority else ""
    st.markdown(
        f"""
        <div class="metric-card{priority_class}" style="--accent:{color}">
          <div class="metric-label">{html.escape(label)}</div>
          <div class="metric-value">{html.escape(value)}</div>
          <div class="metric-note">{html.escape(note)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_summary_dashboard(results: dict[str, object]) -> None:
    summary = results["summary"]
    section_heading(
        "RECONCILIATION OVERVIEW",
        "Summary dashboard",
        "Ringkasan cakupan data, hasil pencocokan, dan potensi data billing yang belum tercatat.",
    )
    row_one = st.columns(6)
    with row_one[0]:
        metric_card("Total DAT DEP", f"{summary['total_dat_dep']:,}", "#1769e0", "Departure records")
    with row_one[1]:
        metric_card("Total DAT ARR", f"{summary['total_dat_arr']:,}", "#1769e0", "Arrival records")
    with row_one[2]:
        metric_card("Total DAT Combined", f"{summary['total_dat_combined']:,}", "#2f80ed", "After hard exclude")
    with row_one[3]:
        metric_card("Total DAT Unique", f"{summary['total_dat_unique']:,}", "#13a8bd", "Best records selected")
    with row_one[4]:
        metric_card("Total STREAM", f"{summary['total_stream']:,}", "#2f80ed", "Comparison records")
    with row_one[5]:
        metric_card("Matched", f"{summary['matched']:,}", "#1f9d68", "Validated in STREAM")

    st.write("")
    row_two = st.columns(5)
    with row_two[0]:
        metric_card("Missing in Stream", f"{summary['missing_in_stream']:,}", "#e84c3d", "Potential unbilled flights", True)
    with row_two[1]:
        metric_card("Need Review", f"{summary['need_review']:,}", "#f28b30", "Time or register validation")
    with row_two[2]:
        metric_card("Extra in Stream", f"{summary['extra_in_stream']:,}", "#7567c7", "Not found in DAT")
    with row_two[3]:
        metric_card("Duplicate DAT", f"{summary['duplicate_dat']:,}", "#d99a1b", "Non-selected copies")
    with row_two[4]:
        metric_card("Accuracy Percentage", f"{summary['accuracy_percentage']:.1f}%", "#13a8bd", "Matched ÷ unique DAT")

    st.write("")
    validation_cols = st.columns(3)
    with validation_cols[0]:
        metric_card(
            "Total Ada di STREAM",
            f"{summary['total_ada_di_stream']:,}",
            "#1f9d68",
            "Actual movement matched",
        )
    with validation_cols[1]:
        metric_card(
            "Total Validasi",
            f"{summary['total_validasi']:,}",
            "#e84c3d",
            "DAT without any STREAM candidate",
            True,
        )
    with validation_cols[2]:
        metric_card(
            "Total Perlu Review STREAM",
            f"{summary['total_perlu_review_stream']:,}",
            "#f28b30",
            "Candidate found with discrepancy",
        )

    st.write("")
    exclusion_cols = st.columns([1, 1, 3])
    with exclusion_cols[0]:
        metric_card(
            "Excluded DAT Non-Billable",
            f"{summary['excluded_dat_non_billable']:,}",
            "#64748b",
            "Removed before recovery",
        )
    with exclusion_cols[1]:
        metric_card(
            "Excluded STREAM Non-Billable",
            f"{summary['excluded_stream_non_billable']:,}",
            "#64748b",
            "Removed before matching",
        )


def filtered_missing_table(
    frame: pd.DataFrame,
    key_prefix: str = "missing",
) -> pd.DataFrame:
    if frame.empty:
        st.info("Tidak ada data MISSING IN STREAM. Semua data DAT sudah ditemukan di STREAM.")
        return frame

    parsed_dates = pd.to_datetime(frame["DATE OF FLIGHT"], errors="coerce").dt.date
    valid_dates = parsed_dates.dropna()
    min_date = min(valid_dates) if len(valid_dates) else date.today()
    max_date = max(valid_dates) if len(valid_dates) else date.today()

    filter_row = st.columns([1.25, 1.2, 1, 1, .9])
    with filter_row[0]:
        selected_dates = st.date_input(
            "Date filter",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            key=f"{key_prefix}_date_filter",
        )
    with filter_row[1]:
        flight_search = st.text_input(
            "Flight number search", placeholder="e.g. LNI200", key=f"{key_prefix}_flight_search"
        )
    with filter_row[2]:
        aerodromes = sorted(value for value in frame["AERODROME"].dropna().astype(str).unique() if value)
        selected_aerodromes = st.multiselect("Aerodrome", aerodromes, key=f"{key_prefix}_aerodrome")
    with filter_row[3]:
        destinations = sorted(value for value in frame["TO FROM"].dropna().astype(str).unique() if value)
        selected_destinations = st.multiselect("TO FROM", destinations, key=f"{key_prefix}_to_from")
    with filter_row[4]:
        movements = sorted(value for value in frame["D/A/L/O"].dropna().astype(str).unique() if value)
        selected_movements = st.multiselect("Movement A/D", movements, default=movements, key=f"{key_prefix}_movement")

    filtered = frame.copy()
    if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
        filtered = filtered.loc[
            parsed_dates.between(selected_dates[0], selected_dates[1], inclusive="both")
        ]
    if flight_search.strip():
        filtered = filtered.loc[
            filtered["FLIGHT NUMBER"].astype(str).str.contains(
                flight_search.strip(), case=False, regex=False, na=False
            )
        ]
    if selected_aerodromes:
        filtered = filtered.loc[filtered["AERODROME"].isin(selected_aerodromes)]
    if selected_destinations:
        filtered = filtered.loc[filtered["TO FROM"].isin(selected_destinations)]
    if selected_movements:
        filtered = filtered.loc[filtered["D/A/L/O"].isin(selected_movements)]
    return filtered


def result_table(
    frame: pd.DataFrame,
    height: int = 480,
    extra_default_columns: list[str] | None = None,
) -> None:
    settings = reconciliation_settings()
    columns = [
        column for column in DISPLAY_RESULT_COLUMNS if column in frame.columns
    ]
    if extra_default_columns:
        columns.extend(
            column
            for column in extra_default_columns
            if column in frame.columns and column not in columns
        )
    if settings["show_audit_columns"]:
        columns.extend(
            column
            for column in AUDIT_COLUMNS
            if column in frame.columns and column not in columns
        )
    st.dataframe(
        frame[columns],
        width="stretch",
        hide_index=True,
        height=height,
        column_config={
            "DATE OF FLIGHT": st.column_config.DateColumn("DATE OF FLIGHT", format="YYYY-MM-DD"),
            "ACTUAL MOVEMENT DATE": st.column_config.DateColumn(
                "ACTUAL MOVEMENT DATE", format="YYYY-MM-DD"
            ),
            "STATUS": st.column_config.TextColumn("STATUS", width="medium"),
            "VALIDASI": st.column_config.TextColumn("VALIDASI", width="large"),
        },
    )


def render_missing_reason_summary(frame: pd.DataFrame) -> None:
    if frame.empty or "MATCH_REASON" not in frame.columns:
        return
    monitored_reasons = [
        "STREAM NOT FOUND",
        "STREAM STATUS OTHER",
        "STREAM INVALID ATA DATE",
        "STREAM INVALID ATD DATE",
        "STREAM TIME MISMATCH",
    ]
    reason_text = frame["MATCH_REASON"].fillna("").astype(str)
    summaries = []
    for reason in monitored_reasons:
        count = int(reason_text.str.contains(reason, regex=False).sum())
        if count:
            summaries.append(f"{reason}: {count:,}")
    if summaries:
        st.warning("Audit reason summary — " + " · ".join(summaries))


def render_export_download(results: dict[str, object], report_bytes: bytes | None) -> None:
    if report_bytes is None:
        with st.spinner("Generating Excel report…"):
            report_bytes = build_excel_report(results)
            st.session_state["finder_report"] = report_bytes
    st.download_button(
        "📥 Download Excel Report",
        data=report_bytes,
        file_name="FINDER_DAT_vs_STREAM_Reconciliation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        width="stretch",
    )


def render_results_page(results: dict[str, object]) -> None:
    render_summary_dashboard(results)
    st.write("")
    section_heading(
        "DETAILED RESULT",
        "Reconciliation result",
        "Prioritaskan pemeriksaan MISSING IN STREAM sebelum meninjau kategori lainnya.",
    )
    summary = results["summary"]
    tabs = st.tabs(
        [
            f"🔎 Validasi ({summary['total_validasi']:,})",
            f"🚨 Missing — Billing Review ({summary['missing_billing_review']:,})",
            f"✅ Matched ({summary['matched']:,})",
            f"🟠 Need Review ({summary['need_review']:,})",
            f"◈ Extra in Stream ({summary['extra_in_stream']:,})",
            f"⚠ Duplicate DAT ({summary['duplicate_dat']:,})",
            "▦ Summary",
        ]
    )
    with tabs[0]:
        st.error(
            "VALIDASI — Hanya data DAT yang tidak memiliki kandidat STREAM setelah pencarian actual movement, original date, dan recovered date."
        )
        validated_missing = filtered_missing_table(
            results["validasi"], key_prefix="validasi"
        )
        st.caption(
            f"Showing {len(validated_missing):,} of {len(results['validasi']):,} validation records"
        )
        result_table(validated_missing)
    with tabs[1]:
        st.markdown(
            """
            <div class="result-banner">
              <div class="result-banner-title">MISSING IN STREAM — Billing Review / Potential Unbilled Flights</div>
              <div class="result-banner-copy">Prioritas billing: STREAM tidak ditemukan, status tidak valid, tanggal salah, atau time mismatch.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        render_missing_reason_summary(results["missing"])
        filtered = filtered_missing_table(
            results["missing"], key_prefix="missing_billing"
        )
        st.caption(
            f"Showing {len(filtered):,} of {len(results['missing']):,} total missing records"
        )
        result_table(filtered)
    with tabs[2]:
        st.success("MATCHED — Data DAT berhasil ditemukan di STREAM.")
        result_table(results["matched"])
    with tabs[3]:
        st.warning(
            "NEED REVIEW — STREAM valid, tetapi selisih waktu melebihi tolerance atau AC REGISTER berbeda."
        )
        result_table(results["need_review"])
    with tabs[4]:
        st.info("EXTRA IN STREAM — Data STREAM tidak memiliki pasangan di DAT DEP/ARR.")
        result_table(results["extra"])
    with tabs[5]:
        st.warning(
            "DUPLICATE DAT — Hanya record yang tidak terpilih. Record terbaik tetap menjadi DAT Unique dan dipakai untuk reconciliation."
        )
        result_table(results["duplicates"], extra_default_columns=DUPLICATE_AUDIT_COLUMNS)
    with tabs[6]:
        st.dataframe(results["summary_table"], width="stretch", hide_index=True)
        if summary["incomplete_dat_keys"] or summary["incomplete_stream_keys"]:
            st.warning(
                f"Incomplete reconciliation keys: DAT {summary['incomplete_dat_keys']:,}, "
                f"STREAM {summary['incomplete_stream_keys']:,}."
            )

    st.write("")
    section_heading(
        "EXPORT",
        "Export reconciliation report",
        "Workbook berisi sheet Validasi, hasil utama, audit recovery, dan Excluded Non-Billable.",
    )
    download_left, download_right = st.columns([1.2, 2])
    with download_left:
        render_export_download(results, st.session_state.get("finder_report"))
    with download_right:
        st.info("Gunakan laporan Excel sebagai lampiran pemeriksaan dan validasi billing.")


def render_upload_page() -> None:
    section_heading(
        "STEP 1 · DATA INGESTION",
        "Upload operational data",
        "Unggah DAT DEP, DAT ARR, dan STREAM. FINDER akan membaca CSV, XLS, XLSX, maupun laporan HTML berformat XLS.",
    )
    upload_columns = st.columns(3, gap="large")
    with upload_columns[0]:
        dep = render_upload_card("Upload DAT DEP", "Departure movement data", "finder_dat_dep", "🛫")
    with upload_columns[1]:
        arr = render_upload_card("Upload DAT ARR", "Arrival movement data", "finder_dat_arr", "🛬")
    with upload_columns[2]:
        stream = render_upload_card("Upload STREAM", "Billing comparison data", "finder_stream", "📡")

    st.write("")
    section_heading(
        "STEP 2 · VALIDATION",
        "Pre-reconciliation validation",
        "FINDER memastikan semua file dan kolom kunci tersedia sebelum proses dijalankan.",
    )
    ready, dep_mapping, arr_mapping, stream_mapping, _ = render_validation(dep, arr, stream)
    settings = reconciliation_settings()

    input_signature = None
    if ready:
        input_signature = "-".join(
            [file_digest(package["data"]) for package in (dep, arr, stream)]
            + [
                str(settings["time_tolerance_minutes"]),
                ",".join(settings["invalid_stream_statuses"]),
                str(settings["treat_invalid_stream_status_as_missing"]),
            ]
        )
        if st.session_state.get("finder_result_signature") not in {None, input_signature}:
            st.info("Input berubah. Jalankan rekonsiliasi kembali untuk memperbarui hasil.")

    st.write("")
    section_heading(
        "STEP 3 · RECONCILIATION",
        "Process DAT vs STREAM",
        "Normalisasi, hard exclude, midnight recovery, deduplikasi flight instance, dan pencocokan dilakukan otomatis.",
    )
    left, center, right = st.columns([1, 2.2, 1])
    with center:
        process_clicked = st.button(
            "▶  PROCESS RECONCILIATION",
            type="primary",
            width="stretch",
            disabled=not ready,
        )

    if process_clicked:
        progress = st.progress(8, text="Reading files")
        try:
            progress.progress(22, text="Normalizing data")
            progress.progress(34, text="Excluding non-billable/internal movement")
            progress.progress(46, text="Recovering adjacent-date midnight movement")
            progress.progress(58, text="Deduplicating DAT flight instances")
            results = reconcile_dat_vs_stream(
                dep["frame"],
                arr["frame"],
                stream["frame"],
                dep_mapping,
                arr_mapping,
                stream_mapping,
                time_tolerance_minutes=int(settings["time_tolerance_minutes"]),
                invalid_stream_statuses=settings["invalid_stream_statuses"],
                treat_invalid_stream_status_as_missing=bool(
                    settings["treat_invalid_stream_status_as_missing"]
                ),
            )
            progress.progress(78, text="Comparing DAT vs STREAM")
            report = build_excel_report(results)
            progress.progress(94, text="Generating result")
            st.session_state["finder_results"] = results
            st.session_state["finder_report"] = report
            st.session_state["finder_result_signature"] = input_signature
            progress.progress(100, text="Reconciliation completed")
            st.success("Reconciliation completed successfully.")
        except Exception as exc:
            st.error(f"Reconciliation failed: {exc}")

    results = st.session_state.get("finder_results")
    if (
        results
        and input_signature
        and st.session_state.get("finder_result_signature") == input_signature
    ):
        st.write("")
        render_summary_dashboard(results)
        st.info("Buka menu **Reconciliation Result** untuk filter dan pemeriksaan detail.")


def render_export_page(results: dict[str, object] | None) -> None:
    section_heading(
        "REPORT DELIVERY",
        "Export report",
        "Unduh satu workbook terstruktur untuk validasi billing dan tindak lanjut operasional.",
    )
    if not results:
        st.warning("Belum ada hasil. Upload tiga file dan jalankan PROCESS RECONCILIATION terlebih dahulu.")
        return

    summary = results["summary"]
    cols = st.columns(3)
    with cols[0]:
        metric_card("Report Status", "READY", "#1f9d68", "Nine worksheets generated")
    with cols[1]:
        metric_card("Missing Records", f"{summary['missing_in_stream']:,}", "#e84c3d", "Priority billing review", True)
    with cols[2]:
        metric_card("Accuracy", f"{summary['accuracy_percentage']:.1f}%", "#13a8bd", "Reconciliation coverage")
    st.write("")
    with st.container(border=True):
        st.markdown("#### 📘 Excel workbook contents")
        st.markdown(
            """
            - **Summary** — reconciliation metrics and accuracy
            - **Validasi** — hanya DAT yang tidak ditemukan sama sekali di STREAM
            - **Missing in Stream** — potential unbilled flights
            - **Matched** — validated DAT records
            - **Need Review** — valid STREAM with time difference above tolerance
            - **Extra in Stream** — STREAM records without DAT pair
            - **Duplicate DAT** — non-selected DAT copies only
            - **Audit Detail** — status, reason, movement datetime, and selected candidate audit
            - **Excluded Non-Billable** — internal DAT/STREAM records removed before reconciliation
            """
        )
        render_export_download(results, st.session_state.get("finder_report"))


def render_about_page() -> None:
    section_heading(
        "ABOUT FINDER",
        "Operational reconciliation made simple",
        "FINDER membantu operator non-IT menemukan potensi penerbangan belum tertagih tanpa formula manual.",
    )
    st.markdown(
        """
        <div class="workflow">
          <div class="workflow-step"><div class="workflow-num">1</div><div class="workflow-title">Upload 3 files</div><div class="workflow-copy">DAT DEP, DAT ARR, dan STREAM.</div></div>
          <div class="workflow-step"><div class="workflow-num">2</div><div class="workflow-title">Process</div><div class="workflow-copy">Normalisasi dan deduplikasi otomatis.</div></div>
          <div class="workflow-step"><div class="workflow-num">3</div><div class="workflow-title">Review missing</div><div class="workflow-copy">Fokus pada potential unbilled flights.</div></div>
          <div class="workflow-step"><div class="workflow-num">4</div><div class="workflow-title">Export Excel</div><div class="workflow-copy">Laporan siap untuk validasi billing.</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    about_left, about_right = st.columns(2)
    with about_left:
        with st.container(border=True):
            st.markdown("#### 🎯 Matching logic")
            st.markdown(
                "Record diprioritaskan berdasarkan **tanggal/waktu actual movement ATD/ATA + Flight Number + ADEP/Aerodrome + ADES/TO FROM + Movement Type**. Date of Flight dan recovered date digunakan sebagai fallback."
            )
    with about_right:
        with st.container(border=True):
            st.markdown("#### 🔒 Data handling")
            st.markdown(
                "Pada deployment cloud, file diproses oleh sesi aplikasi Streamlit dan tidak dikomit ke repository GitHub oleh FINDER."
            )


page = render_sidebar()
render_header()

if "Upload Data" in page:
    render_upload_page()
elif "Reconciliation Result" in page:
    results = st.session_state.get("finder_results")
    if results:
        render_results_page(results)
    else:
        section_heading(
            "NO RESULT YET",
            "Reconciliation result",
            "Jalankan proses rekonsiliasi dari menu Upload Data terlebih dahulu.",
        )
        st.warning("Belum ada hasil rekonsiliasi.")
elif "Export Report" in page:
    render_export_page(st.session_state.get("finder_results"))
else:
    render_about_page()
