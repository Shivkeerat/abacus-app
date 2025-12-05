# ============================================================
# SECTION 1 — Imports & Global Config
# ============================================================
import io
from pathlib import Path
import pandas as pd
import streamlit as st

# Streamlit Page Setup
st.set_page_config(
    page_title="Abacus File Loader & XLOOKUP Tool",
    layout="wide",
)

# --------- CONFIG (change only if folder changes) ----------
ABACUS_FOLDER = Path(r"\\vgd-FILES\DJM\DjmTemp")  
TEXT_ENCODING = "latin-1"
TEXT_SEPARATOR = "|"


# Helper for headings
def section_title(emoji, text):
    st.markdown(f"### {emoji} {text}")


# ============================================================
# SECTION 2 — Cached Helpers (Listing, Loading, Cleaning)
# ============================================================
@st.cache_data(show_spinner=False)
def list_abacus_files(folder: Path):
    """List all .TXT files in the Abacus folder."""
    files = []
    if not folder.exists():
        return pd.DataFrame(columns=["name", "modified"])

    for p in folder.glob("*.TXT"):
        files.append({
            "name": p.name,
            "modified": pd.to_datetime(p.stat().st_mtime, unit="s"),
        })

    df = pd.DataFrame(files)
    if not df.empty:
        df = df.sort_values("name").reset_index(drop=True)
    return df


@st.cache_data(show_spinner=False)
def parse_text_file(path: Path, nrows=None):
    """
    Ultra-safe Abacus TXT parser.
    - Reads raw lines
    - Splits manually on |
    - Fixes unbalanced quotes
    - Skips corrupt lines automatically
    - Trims spaces + removes double quotes
    """
    rows = []
    col_count = None
    line_no = 0

    with open(path, "r", encoding="latin-1", errors="ignore") as f:
        for line in f:
            line_no += 1
            line = line.strip("\n").strip()

            # Stop early for preview
            if nrows and len(rows) >= nrows:
                break

            # Skip empty lines
            if not line:
                continue

            # Fix broken or unbalanced quotes
            if line.count('"') % 2 != 0:
                line = line.replace('"', '')

            # Split by |
            parts = line.split("|")

            # First row defines column count
            if col_count is None:
                col_count = len(parts)
            else:
                # If row has fewer/more columns → fix it
                if len(parts) < col_count:
                    parts += [""] * (col_count - len(parts))
                elif len(parts) > col_count:
                    parts = parts[:col_count]

            # Clean each field
            cleaned = [p.strip().strip('"') for p in parts]
            rows.append(cleaned)

    # Build DataFrame
    df = pd.DataFrame(rows)

    # Use first row as header
    df.columns = df.iloc[0].tolist()
    df = df[1:].reset_index(drop=True)

    return df


# SECTION 8 — EXPORT UTILITIES (UPDATED TO USE openpyxl)

def export_df(df, filename_prefix="export", key_prefix="export"):
    """
    CSV + Excel export with unique widget keys.
    """

    st.markdown("### ⬇ Download Output")

    # ---- CSV ----
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download CSV",
        data=csv,
        file_name=f"{filename_prefix}.csv",
        mime="text/csv",
        key=f"{key_prefix}_csv"
    )

    # ---- Excel (openpyxl) ----
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")

    st.download_button(
        label="📘 Download Excel",
        data=excel_buffer.getvalue(),
        file_name=f"{filename_prefix}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"{key_prefix}_excel"
    )


# ============================================================
# SECTION 3 — Sidebar (Folder, Sorting, File Selection)
# ============================================================
st.sidebar.header("📂 Checking folder:")
st.sidebar.code(str(ABACUS_FOLDER))

if ABACUS_FOLDER.exists():
    st.sidebar.success("UNC folder is accessible.")
else:
    st.sidebar.error("Folder not found.")

sort_choice = st.sidebar.radio(
    "Sort files by:",
    ["Name (A–Z)", "Date modified (Newest first)"]
)

files_df = list_abacus_files(ABACUS_FOLDER)

if not files_df.empty:
    if sort_choice == "Date modified (Newest first)":
        files_df = files_df.sort_values("modified", ascending=False)
    else:
        files_df = files_df.sort_values("name")

    st.sidebar.subheader("📑 Available Abacus Files")
    st.sidebar.dataframe(
        files_df.rename(columns={"name": "Name", "modified": "Modified"}),
        use_container_width=True,
        hide_index=True,
        height=260,
    )

file_names = files_df["name"].tolist() if not files_df.empty else []

st.sidebar.subheader("📌 Select files to load:")
selected_files = st.sidebar.multiselect("Choose files:", file_names)

file_paths = {name: ABACUS_FOLDER / name for name in file_names}


# ============================================================
# SECTION 4 — Main Title + Table Loader
# ============================================================
st.title("Abacus File Loader & XLOOKUP Tool")

if not selected_files:
    st.info("Choose one or more files from the left to continue.")
else:
    section_title("📊", "Tables")


for fname in selected_files:
    path = file_paths[fname]

    with st.container(border=True):
        st.markdown(f"**{fname} — Table Setup**")
        st.caption(f"Path: `{path}`")

        try:
            df_preview = parse_text_file(path, nrows=50)
        except Exception as e:
            st.error(f"Error loading preview: {e}")
            continue

        if df_preview.empty:
            st.warning("This file appears empty.")
            continue

        all_cols = df_preview.columns.tolist()

        default_keep = st.session_state.get(f"keep_{fname}", all_cols)
        keep_cols = st.multiselect(
            "Columns to use:",
            options=all_cols,
            default=default_keep,
            key=f"keep_{fname}",
        )

        if not keep_cols:
            st.warning("Select at least one column.")
            continue

        st.write("Preview (first 50 rows):")
        st.dataframe(df_preview[keep_cols], use_container_width=True, height=260)

        if st.button(f"⬇ Download full cleaned file ({fname})", key=f"download_{fname}"):
            try:
                df_full = parse_text_file(path)
                df_full = df_full[keep_cols]
                export_df(df_full, fname.replace(".TXT", "") + "_FULL", f"full_{fname}")
            except Exception as e:
                st.error(f"Download error: {e}")


# ============================================================
# SECTION 5 — XLOOKUP Builder (MULTI-RULE)
# ============================================================
st.markdown("---")
section_title("🧩", "XLOOKUP Builder (Multi-Rule)")

if len(selected_files) < 2:
    st.info("Select at least 2 files to enable XLOOKUP.")
    st.stop()

# Select primary + secondary table
colA, colB = st.columns(2)

with colA:
    primary_file = st.selectbox("Primary table (rows kept):", selected_files)

with colB:
    secondary_candidates = [f for f in selected_files if f != primary_file]
    secondary_file = st.selectbox("Secondary table (lookup source):", secondary_candidates)

# Load full tables
try:
    primary_df = parse_text_file(file_paths[primary_file])
    secondary_df = parse_text_file(file_paths[secondary_file])
except Exception as e:
    st.error(f"Error loading full tables: {e}")
    st.stop()

p_cols = primary_df.columns.tolist()
s_cols = secondary_df.columns.tolist()

st.success(f"Loaded Primary: {len(primary_df):,} rows — Secondary: {len(secondary_df):,} rows")

# Number of XLOOKUP rules
num_rules = st.number_input("Number of lookup rules:", min_value=1, max_value=10, value=1)

rules = []

for i in range(int(num_rules)):
    st.markdown(f"#### 🔎 Lookup Rule #{i+1}")
    c1, c2, c3 = st.columns(3)

    with c1:
        pk = st.selectbox("Primary lookup column", p_cols, key=f"pk_{i}")
    with c2:
        sk = st.selectbox("Secondary match column", s_cols, key=f"sk_{i}")
    with c3:
        so = st.selectbox("Secondary return column", s_cols, key=f"so_{i}")

    rules.append((pk, sk, so))

run_lookup = st.button("🚀 Run XLOOKUP", type="primary")

if run_lookup:
    result_df = primary_df.copy()
    used_columns = set(result_df.columns)

    for idx, (pk, sk, so) in enumerate(rules, start=1):

        base_name = f"{pk}_FROM_{so}"
        new_col = base_name
        c = 2
        while new_col in used_columns:
            new_col = f"{base_name}_{c}"
            c += 1

        used_columns.add(new_col)

        try:
            map_series = secondary_df.set_index(sk)[so]
            result_df[new_col] = result_df[pk].map(map_series)
        except Exception as e:
            st.error(f"Rule #{idx} failed ({pk} → {sk} → {so}): {e}")

    st.success("XLOOKUP completed!")
    st.dataframe(result_df.head(50), use_container_width=True)

    export_df(result_df,
              f"{primary_file.replace('.TXT','')}_XLOOKUP_RESULT",
              key_prefix="xlookup_final")
