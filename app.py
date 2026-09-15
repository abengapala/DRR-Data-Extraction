import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Daily Remark Header Fixer", layout="wide")

# ---------------------------------------------------------------------------
# Your standard/target header — the order you normally work with.
# Edit this list directly if your standard header ever changes.
# ---------------------------------------------------------------------------
TARGET_HEADER = [
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


import datetime as _dt


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
    """Turn any cell value into the exact text format we want, preserving
    real dates/times correctly instead of letting them turn into ISO strings.
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
        # avoid turning whole numbers like 1.0 into "1.0"
        if val.is_integer():
            return str(int(val))
        return str(val)
    return str(val).strip()


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Apply _format_cell to every cell so Date/Time (and everything else)
    come out as clean, correct text instead of whatever pandas inferred.
    """
    return df.apply(lambda col: col.map(_format_cell))


def fix_header(df: pd.DataFrame, target_header: list[str]):
    """Reorder/select df's columns to match target_header.
    Returns (fixed_df, missing_cols, extra_cols).
    """
    # normalize for matching (trim spaces, case-insensitive) but keep original target names
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


st.title("📋 Daily Remark Header Fixer")
st.caption(
    "Upload the newly downloaded file and this will automatically reorder/align "
    "its columns to match your standard working header — no more manual copy-pasting."
)

with st.expander("Standard header currently in use (edit app.py to change it)"):
    st.write(f"{len(TARGET_HEADER)} columns:")
    st.code(", ".join(TARGET_HEADER))

uploaded_file = st.file_uploader(
    "Upload the new data file (.csv or .xlsx)", type=["csv", "xlsx", "xls"]
)

if uploaded_file is not None:
    try:
        raw_df = load_file(uploaded_file)
    except Exception as e:
        st.error(f"Couldn't read that file: {e}")
        st.stop()

    st.success(f"Loaded **{uploaded_file.name}** — {raw_df.shape[0]} rows, {raw_df.shape[1]} columns")

    raw_df = normalize_dataframe(raw_df)
    fixed_df, missing_cols, extra_cols = fix_header(raw_df, TARGET_HEADER)

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
        file_name="Daily_Remark_Fixed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload a file above to get started.")
