import streamlit as st
import pandas as pd
import re
import datetime as _dt
from io import BytesIO

st.set_page_config(page_title="Header Fixer", layout="wide")

# ---------------------------------------------------------------------------
# Standard/target headers for each report type. Edit these lists directly if
# a standard header ever changes.
# ---------------------------------------------------------------------------
DAILY_REMARK_HEADER = [
    "S.No", "Date", "Time", "Debtor", "Account No.", "Card No.", "Service No.",
    "DPD", "Call Status", "Status", "Remark", "Remark By", "Remark Type",
    "Field Visit Date", "Collector", "Client", "Product Description",
    "Product Type", "Batch No", "Account Type", "Relation", "PTP Amount",
    "PTP Date", "Next Call", "Claim Paid Amount", "Claim Paid Date",
    "Dialed Number", "Days Past Write Off", "Balance", "Contact Type",
    "Cycle", "Old IC", "I.C Issue Date", "Bank Code", "Over Limit Amount",
    "Min Payment", "Due Date", "Monthly Installment", "30 Days", "MIA",
    "Area", "Call Duration", "Talk Time Duration", "Debtor ID",
    "Black Case No.", "Red Case No.", "Court Name", "Lawyer", "Legal Stage",
    "Legal Status", "Next Legal Follow up",
]

AUTOSTAT_HEADER = [
    "CH CODE", "ACCOUNT NUMBER", "STATUS CODE", "REMARKS", "REMARKS BY",
    "REMARKS DATE", "PTP DATE", "PTP AMOUNT",
]

# Autostat date/time columns need special reformatting (the system expects
# these exact patterns to recognize the file): REMARKS DATE -> MM/DD/YYYY
# HH:MM:SS (24-hour), PTP DATE -> MM/DD/YYYY (date only).
AUTOSTAT_DATETIME_COLS = {"REMARKS DATE": False, "PTP DATE": True}  # value = date_only?

REPORT_CONFIGS = {
    "Daily Remark": {
        "target_header": DAILY_REMARK_HEADER,
        "datetime_cols": {},  # handled generically for this report type
    },
    "Autostat": {
        "target_header": AUTOSTAT_HEADER,
        "datetime_cols": AUTOSTAT_DATETIME_COLS,
    },
}


def load_file(uploaded_file):
    """Read an uploaded CSV or Excel file into a DataFrame.

    Deliberately does NOT force dtype=str at read time, because doing so
    mangles real Excel date/time cells (they get parsed into datetime
    objects by openpyxl *before* the dtype cast happens, so forcing str
    afterwards turns a proper date like 09/09/2026 into the ugly ISO form
    '2026-09-09 00:00:00'). Instead we read with native types and format
    each cell ourselves in `normalize_dataframe`.
    """
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, dtype=str, keep_default_na=False)
    else:
        return pd.read_excel(uploaded_file)


def _format_cell(val):
    """Turn any cell value into clean text, preserving real dates/times
    correctly instead of letting them turn into ISO strings.
    """
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass  # some values (e.g. arrays) can't be checked with pd.isna; fall through
    if isinstance(val, pd.Timestamp):
        val = val.to_pydatetime()
    if isinstance(val, _dt.datetime):
        if val.time() == _dt.time(0, 0):
            return val.strftime("%d/%m/%Y")  # pure date -> dd/mm/yyyy
        return val.strftime("%d/%m/%Y %I:%M:%S %p")  # date+time, just in case
    if isinstance(val, _dt.date):
        return val.strftime("%d/%m/%Y")
    if isinstance(val, _dt.time):
        return val.strftime("%I:%M:%S %p")
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
    return str(val).strip()


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Apply _format_cell to every cell so Date/Time (and everything else)
    come out as clean, correct text instead of whatever pandas inferred.
    """
    return df.apply(lambda col: col.map(_format_cell))


_TRAILING_AMPM_RE = re.compile(r"\s*(AM|PM)\s*$", re.IGNORECASE)


def parse_autostat_datetime(raw, date_only: bool = False) -> str:
    """Reformat an Autostat date/time value into the exact pattern the
    system expects: MM/DD/YYYY HH:MM:SS (24-hour), or MM/DD/YYYY for
    date-only columns.

    Handles the quirky source format seen in raw exports, e.g.
    '9/14/2026  18:00:00 PM' (double space, 24-hour time with a bogus
    trailing AM/PM tag) by stripping the bogus AM/PM marker and keeping
    the 24-hour value as-is, rather than misapplying an AM/PM conversion.
    """
    if raw is None:
        return ""
    s = str(raw).strip()
    if not s or s.lower() == "nan":
        return ""
    s = re.sub(r"\s+", " ", s)          # collapse repeated spaces
    s = _TRAILING_AMPM_RE.sub("", s)     # drop bogus trailing AM/PM tag

    formats = ["%m/%d/%Y"] if date_only else [
        "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%m/%d/%Y",
    ]
    for fmt in formats:
        try:
            dt = _dt.datetime.strptime(s, fmt)
            return dt.strftime("%m/%d/%Y") if date_only else dt.strftime("%m/%d/%Y %H:%M:%S")
        except ValueError:
            continue
    # last resort: let pandas try to parse it
    try:
        dt = pd.to_datetime(s)
        return dt.strftime("%m/%d/%Y") if date_only else dt.strftime("%m/%d/%Y %H:%M:%S")
    except Exception:
        return s  # give up, return original text rather than losing data


def apply_datetime_formatting(df: pd.DataFrame, datetime_cols: dict) -> pd.DataFrame:
    """Reformat specific columns using parse_autostat_datetime, per the
    report config's datetime_cols mapping of {column_name: date_only_bool}.
    """
    for col, date_only in datetime_cols.items():
        if col in df.columns:
            df[col] = df[col].map(lambda v: parse_autostat_datetime(v, date_only=date_only))
    return df


def fix_header(df: pd.DataFrame, target_header: list[str]):
    """Reorder/select df's columns to match target_header.
    Returns (fixed_df, missing_cols, extra_cols).
    """
    col_lookup = {str(c).strip().lower(): c for c in df.columns}

    missing_cols = []
    fixed = pd.DataFrame()

    for target_col in target_header:
        key = target_col.strip().lower()
        if key in col_lookup:
            fixed[target_col] = df[col_lookup[key]]
        else:
            fixed[target_col] = ""  # fill blank if this file doesn't have it
            missing_cols.append(target_col)

    used_keys = {t.strip().lower() for t in target_header}
    extra_cols = [c for c in df.columns if str(c).strip().lower() not in used_keys]

    return fixed, missing_cols, extra_cols


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="sheet")
    return buffer.getvalue()


st.title("📋 Header Fixer")
st.caption(
    "Upload the newly downloaded file and this will automatically reorder/align "
    "its columns (and fix date/time formats) to match your standard working header."
)

report_type = st.radio("Which file are you fixing?", list(REPORT_CONFIGS.keys()), horizontal=True)
config = REPORT_CONFIGS[report_type]
target_header = config["target_header"]
datetime_cols = config["datetime_cols"]

with st.expander(f"Standard header currently in use for {report_type} (edit app.py to change it)"):
    st.write(f"{len(target_header)} columns:")
    st.code(", ".join(target_header))

uploaded_file = st.file_uploader(
    "Upload the new data file (.csv or .xlsx)", type=["csv", "xlsx", "xls"], key=report_type
)

if uploaded_file is not None:
    try:
        raw_df = load_file(uploaded_file)
    except Exception as e:
        st.error(f"Couldn't read that file: {e}")
        st.stop()

    st.success(f"Loaded **{uploaded_file.name}** — {raw_df.shape[0]} rows, {raw_df.shape[1]} columns")

    raw_df = normalize_dataframe(raw_df)
    fixed_df, missing_cols, extra_cols = fix_header(raw_df, target_header)

    if datetime_cols:
        fixed_df = apply_datetime_formatting(fixed_df, datetime_cols)

    col1, col2 = st.columns(2)
    with col1:
        if missing_cols:
            st.warning(
                f"⚠️ {len(missing_cols)} column(s) from your standard header were "
                f"NOT found in this file (left blank):\n\n- " + "\n- ".join(missing_cols)
            )
        else:
            st.info("✅ All standard columns were found in this file.")
    with col2:
        if extra_cols:
            st.info(
                f"ℹ️ {len(extra_cols)} extra column(s) in this file were dropped "
                f"(not part of your standard header):\n\n- " + "\n- ".join(extra_cols)
            )
        else:
            st.info("No extra columns found.")

    st.subheader("Preview (first 20 rows)")
    st.dataframe(fixed_df.head(20), use_container_width=True)

    excel_bytes = to_excel_bytes(fixed_df)
    st.download_button(
        label="⬇️ Download fixed file (.xlsx)",
        data=excel_bytes,
        file_name=f"{report_type.replace(' ', '_')}_Fixed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload a file above to get started.")
