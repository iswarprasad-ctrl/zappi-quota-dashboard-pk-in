import streamlit as st
import pandas as pd
import io
import re
import requests

st.set_page_config(layout="wide", page_title="Zappi Multi-Market Dashboard")

st.title("📊 ZAPPI SERVICES MARKET PERFORMANCE REVIEW - DASHBOARD")
st.markdown("---")

# =========================================================================
# 0. RAW-DATA COLUMN MAPPING -- confirmed against the live raw data export
# =========================================================================
# Confirmed by inspecting the actual raw file (same Google Sheet as FILE_ID
# below): there is NO dedicated "City_Code", "ISEC", or "SEC" column. The
# schema is a single flat sheet shared by every country and includes:
#   Survey Country, Survey Language, Project Name, Device, Supplier Group,
#   Gender, Age, Region, Region Code, Division, Division Code, State,
#   Buyer Url (contains extra query-string params per respondent)
#
# The India socio-economic band (ISEC) is NOT its own column -- it's a
# query parameter inside "Buyer Url", e.g.
#   .../start_project/446309?age=24&...&india_socio_economic_classification=114&...
# CONFIRMED: param name is "india_socio_economic_classification".
#
# The Pakistan SEC band equivalent was NOT visible in the rows we could
# inspect (only India rows were in the sample). We assume the analogous
# param name below by pattern -- TODO: confirm once you can see a Pakistan
# row's Buyer Url, and update SEC_URL_PARAM_PK if it's named differently.
# =========================================================================

REGION_COLUMN_PK = "Region"                 # confirmed raw column; used for Pakistan's Bal/Khy/Pak, Pun/Isl, Sindh rows
REGION_COLUMN_FALLBACK_PK = "State"         # tried if "Region" doesn't contain matching values
ISEC_URL_COLUMN = "Buyer Url"               # confirmed raw column holding query-string params
ISEC_URL_PARAM_IN = "india_socio_economic_classification"   # CONFIRMED present in raw data
SEC_URL_PARAM_PK = "pakistan_socio_economic_classification"  # TODO: UNCONFIRMED -- verify against a real Pakistan row

# Map each quota-sheet region label -> list of possible raw-data values that should count toward it.
# These are reasonable province-name guesses; confirm against actual Pakistan rows once available.
REGION_VALUE_MAP_PK = {
    "Bal/Khy/Pak": ["Balochistan", "Khyber Pakhtunkhwa", "KPK", "Bal/Khy/Pak"],
    "Pun/Isl": ["Punjab", "Islamabad", "Pun/Isl"],
    "Sindh": ["Sindh"],
}

# =========================================================================
# 1. CLOUD FILE CONFIGURATION (Direct Cloud Sync)
# =========================================================================
# If India and Pakistan raw data live in TWO SEPARATE files, put each file's
# Google Drive ID below. If they're still the same single combined file,
# just set both to the same ID -- everything else works unchanged either way.
FILE_ID_INDIA = "1fFcmQFcKUYGtr5_IpMQ7V6cal0k15vzd"      # TODO: replace with India's file ID if separate
FILE_ID_PAKISTAN = "1fFcmQFcKUYGtr5_IpMQ7V6cal0k15vzd"   # TODO: replace with Pakistan's file ID if separate


def drive_url(file_id):
    return f"https://drive.google.com/uc?id={file_id}&export=download"


@st.cache_data(ttl=600)
def load_excel_from_cloud(url):
    try:
        session = requests.Session()
        response = session.get(url, stream=True)

        if response.status_code == 404:
            st.warning("⚠️ The source raw-data file has been deleted or moved from Google Drive.")
            return pd.DataFrame()

        token = None
        for key, value in response.cookies.items():
            if key.startswith("download_warning"):
                token = value
                break

        if token:
            url = url + f"&confirm={token}"
            response = session.get(url, stream=True)

        response.raise_for_status()
        return pd.read_excel(io.BytesIO(response.content), engine="openpyxl")
    except Exception:
        st.warning("⚠️ Source data file not found on Google Drive.")
        return pd.DataFrame()


def load_country_data(file_id):
    df = load_excel_from_cloud(drive_url(file_id))
    if df.empty:
        return df
    df.columns = df.columns.astype(str).str.strip()
    if "Project Name" in df.columns:
        df = df.dropna(subset=["Project Name"])
    return df


# Load once per country. If FILE_ID_INDIA == FILE_ID_PAKISTAN, @st.cache_data
# means this is still only ONE network fetch -- no extra cost either way.
raw_df_by_country = {
    "India": load_country_data(FILE_ID_INDIA),
    "Pakistan": load_country_data(FILE_ID_PAKISTAN),
}

if all(df.empty for df in raw_df_by_country.values()):
    st.info("💡 Please upload the raw data file(s) back to the Google Drive folder to resume tracking.")
    st.stop()

# =========================================================================
# 2. QUOTA CONFIG -- built directly from Zappi_-_PK_and_IN_quotas.xlsx
# =========================================================================
# Each row dict:
#   label      -> row label shown in the table
#   type       -> Device / Region / SEC / ISEC / Age-Gender / AllData / Total
#   match      -> value(s) to match in the raw data for this row (ignored for Total rows)
#   target_total          -> TOTAL target
#   targets    -> dict of {supplier_column_name: target or None for "no cap"}
#   sum_of     -> (Total rows only) list of row labels whose collected counts should be summed

COUNTRY_CONFIGS = {
    "India": {
        "supplier_columns": ["GROUP MP (ONLINE)", "MARKETEXCEL (OFFLINE)"],
        # None = "everything not matched by the other columns" (i.e. the remainder)
        "supplier_match": {
            "GROUP MP (ONLINE)": ["Group MP"],
            "MARKETEXCEL (OFFLINE)": None,
        },
        "rows": [
            {"label": "Region", "type": "AllData", "target_total": 200,
             "targets": {"GROUP MP (ONLINE)": None, "MARKETEXCEL (OFFLINE)": None}},

            {"label": "Desktop", "type": "Device", "match": "Desktop", "target_total": 200,
             "targets": {"GROUP MP (ONLINE)": None, "MARKETEXCEL (OFFLINE)": 0}},
            {"label": "Mobile", "type": "Device", "match": "Mobile", "target_total": 200,
             "targets": {"GROUP MP (ONLINE)": None, "MARKETEXCEL (OFFLINE)": 120}},
            {"label": "Device Total", "type": "Total", "target_total": 400,
             "targets": {"GROUP MP (ONLINE)": None, "MARKETEXCEL (OFFLINE)": 120},
             "sum_of": ["Desktop", "Mobile"]},

            {"label": "Male - 16-24", "type": "Age-Gender", "match": "Male:16-24", "target_total": 25,
             "targets": {"GROUP MP (ONLINE)": 5, "MARKETEXCEL (OFFLINE)": 20}},
            {"label": "Female - 16-24", "type": "Age-Gender", "match": "Female:16-24", "target_total": 25,
             "targets": {"GROUP MP (ONLINE)": 5, "MARKETEXCEL (OFFLINE)": 20}},
            {"label": "Male 25-44", "type": "Age-Gender", "match": "Male:25-44", "target_total": 45,
             "targets": {"GROUP MP (ONLINE)": 9, "MARKETEXCEL (OFFLINE)": 36}},
            {"label": "Female 25-44", "type": "Age-Gender", "match": "Female:25-44", "target_total": 45,
             "targets": {"GROUP MP (ONLINE)": 9, "MARKETEXCEL (OFFLINE)": 36}},
            {"label": "Male 45-75", "type": "Age-Gender", "match": "Male:45-75", "target_total": 30,
             "targets": {"GROUP MP (ONLINE)": 6, "MARKETEXCEL (OFFLINE)": 24}},
            {"label": "Female 45-75", "type": "Age-Gender", "match": "Female:45-75", "target_total": 30,
             "targets": {"GROUP MP (ONLINE)": 6, "MARKETEXCEL (OFFLINE)": 24}},
            {"label": "Gender-Age Total", "type": "Total", "target_total": 200,
             "targets": {"GROUP MP (ONLINE)": 40, "MARKETEXCEL (OFFLINE)": 160},
             "sum_of": ["Male - 16-24", "Female - 16-24", "Male 25-44", "Female 25-44",
                        "Male 45-75", "Female 45-75"]},

            {"label": "ISEC 1-3", "type": "ISEC", "match": "1-3", "target_total": 40,
             "targets": {"GROUP MP (ONLINE)": 40, "MARKETEXCEL (OFFLINE)": 0}},
            {"label": "ISEC 4-5", "type": "ISEC", "match": "4-5", "target_total": 40,
             "targets": {"GROUP MP (ONLINE)": 40, "MARKETEXCEL (OFFLINE)": 40}},
            {"label": "ISEC 6-7", "type": "ISEC", "match": "6-7", "target_total": 60,
             "targets": {"GROUP MP (ONLINE)": 0, "MARKETEXCEL (OFFLINE)": 60}},
            {"label": "ISEC 8-12", "type": "ISEC", "match": "8-12", "target_total": 60,
             "targets": {"GROUP MP (ONLINE)": 0, "MARKETEXCEL (OFFLINE)": 60}},
            {"label": "ISEC Total", "type": "Total", "target_total": 200,
             "targets": {"GROUP MP (ONLINE)": 40, "MARKETEXCEL (OFFLINE)": 160},
             "sum_of": ["ISEC 1-3", "ISEC 4-5", "ISEC 6-7", "ISEC 8-12"]},
        ],
    },
    "Pakistan": {
        "supplier_columns": ["GROUPMP", "CPX", "MARKETEXCEL"],
        "supplier_match": {
            "GROUPMP": ["Group MP", "GroupMP"],
            "CPX": ["CPX"],
            "MARKETEXCEL": None,  # remainder
        },
        "rows": [
            {"label": "TOTAL", "type": "AllData", "target_total": 200,
             "targets": {"GROUPMP": 28, "CPX": 40, "MARKETEXCEL": 132}},

            {"label": "Mobile/Tablet", "type": "Device", "match": ["Mobile", "Tablet"], "target_total": 180,
             "targets": {"GROUPMP": None, "CPX": None, "MARKETEXCEL": None}},
            {"label": "Desktop (10% min)", "type": "Device", "match": "Desktop", "target_total": 20,
             "targets": {"GROUPMP": None, "CPX": None, "MARKETEXCEL": None}},
            {"label": "Device Total", "type": "Total", "target_total": 200,
             "targets": {"GROUPMP": None, "CPX": None, "MARKETEXCEL": None},
             "sum_of": ["Mobile/Tablet", "Desktop (10% min)"]},

            {"label": "Bal/Khy/Pak", "type": "Region", "match": "Bal/Khy/Pak", "target_total": 42,
             "targets": {"GROUPMP": 42, "CPX": 42, "MARKETEXCEL": 42}},
            {"label": "Pun/Isl", "type": "Region", "match": "Pun/Isl", "target_total": 110,
             "targets": {"GROUPMP": 110, "CPX": 110, "MARKETEXCEL": 110}},
            {"label": "Sindh", "type": "Region", "match": "Sindh", "target_total": 48,
             "targets": {"GROUPMP": 48, "CPX": 48, "MARKETEXCEL": 48}},
            {"label": "Region Total", "type": "Total", "target_total": 200,
             "targets": {"GROUPMP": 200, "CPX": 200, "MARKETEXCEL": 200},
             "sum_of": ["Bal/Khy/Pak", "Pun/Isl", "Sindh"]},

            {"label": "SEC A", "type": "SEC", "match": "A", "target_total": 24,
             "targets": {"GROUPMP": 10, "CPX": 15, "MARKETEXCEL": 0}},
            {"label": "SEC B", "type": "SEC", "match": "B", "target_total": 44,
             "targets": {"GROUPMP": 19, "CPX": 26, "MARKETEXCEL": 0}},
            {"label": "SEC C", "type": "SEC", "match": "C", "target_total": 68,
             "targets": {"GROUPMP": 0, "CPX": 0, "MARKETEXCEL": 68}},
            {"label": "SEC D", "type": "SEC", "match": "D", "target_total": 64,
             "targets": {"GROUPMP": 0, "CPX": 0, "MARKETEXCEL": 64}},
            {"label": "SEC Total", "type": "Total", "target_total": 200,
             "targets": {"GROUPMP": 29, "CPX": 41, "MARKETEXCEL": 132},
             "sum_of": ["SEC A", "SEC B", "SEC C", "SEC D"]},

            {"label": "Male 16-24", "type": "Age-Gender", "match": "Male:16-24", "target_total": 32,
             "targets": {"GROUPMP": 5, "CPX": 6, "MARKETEXCEL": 21}},
            {"label": "Female 16-25", "type": "Age-Gender", "match": "Female:16-25", "target_total": 32,
             "targets": {"GROUPMP": 5, "CPX": 6, "MARKETEXCEL": 21}},
            {"label": "Male 25-44", "type": "Age-Gender", "match": "Male:25-44", "target_total": 41,
             "targets": {"GROUPMP": 6, "CPX": 8, "MARKETEXCEL": 27}},
            {"label": "Female 25-44", "type": "Age-Gender", "match": "Female:25-44", "target_total": 41,
             "targets": {"GROUPMP": 6, "CPX": 8, "MARKETEXCEL": 27}},
            {"label": "Male 45+", "type": "Age-Gender", "match": "Male:45-120", "target_total": 27,
             "targets": {"GROUPMP": 4, "CPX": 6, "MARKETEXCEL": 18}},
            {"label": "Female 45+", "type": "Age-Gender", "match": "Female:45-120", "target_total": 27,
             "targets": {"GROUPMP": 4, "CPX": 6, "MARKETEXCEL": 18}},
            {"label": "Age-Gender Total", "type": "Total", "target_total": 200,
             "targets": {"GROUPMP": 30, "CPX": 40, "MARKETEXCEL": 128},
             "sum_of": ["Male 16-24", "Female 16-25", "Male 25-44", "Female 25-44", "Male 45+", "Female 45+"]},
        ],
    },
}

# =========================================================================
# 3. FILTERS -- now per-country, since each country can have its own file
# =========================================================================
# (Language/Project filters moved inside render_country_tab below, because
#  India and Pakistan may now come from two different raw files with two
#  different sets of languages/projects.)


# =========================================================================
# 4. MATCHING + COUNTING LOGIC
# =========================================================================
def extract_url_param(series, param_name):
    """Pull a query-string parameter's value out of a URL column, e.g. '...&sec=114&...' -> '114'."""
    pattern = re.escape(param_name) + r"=([^&]+)"
    return series.astype(str).str.extract(pattern, expand=False)


def filter_for_row(df, row):
    """Return the slice of df matching this row's criteria. None if the needed column is missing."""
    if row["type"] == "AllData":
        return df.copy()

    if row["type"] == "Device":
        if "Device" not in df.columns:
            return None
        matches = row["match"] if isinstance(row["match"], list) else [row["match"]]
        matches = [m.strip().lower() for m in matches]
        return df[df["Device"].astype(str).str.strip().str.lower().isin(matches)]

    if row["type"] == "Region":
        valid_values = [v.strip().lower() for v in REGION_VALUE_MAP_PK.get(row["match"], [row["match"]])]
        region_col = None
        if REGION_COLUMN_PK in df.columns and df[REGION_COLUMN_PK].astype(str).str.strip().str.lower().isin(valid_values).any():
            region_col = REGION_COLUMN_PK
        elif REGION_COLUMN_FALLBACK_PK in df.columns:
            region_col = REGION_COLUMN_FALLBACK_PK
        if region_col is None:
            return None
        return df[df[region_col].astype(str).str.strip().str.lower().isin(valid_values)]

    if row["type"] == "SEC":
        # SEC band is embedded as a URL query param (no dedicated raw column) -- see SEC_URL_PARAM_PK note at top.
        if ISEC_URL_COLUMN not in df.columns:
            return None
        temp = df.copy()
        temp["_SEC_VALUE"] = extract_url_param(temp[ISEC_URL_COLUMN], SEC_URL_PARAM_PK)
        return temp[temp["_SEC_VALUE"].astype(str).str.strip().str.upper() == row["match"].upper()]

    if row["type"] == "ISEC":
        # ISEC band is embedded as a URL query param inside Buyer Url (confirmed).
        if ISEC_URL_COLUMN not in df.columns:
            return None
        low, high = map(int, row["match"].split("-"))
        temp = df.copy()
        temp["_ISEC_VALUE"] = extract_url_param(temp[ISEC_URL_COLUMN], ISEC_URL_PARAM_IN)
        numeric = pd.to_numeric(temp["_ISEC_VALUE"], errors="coerce")
        return temp[(numeric >= low) & (numeric <= high)]

    if row["type"] == "Age-Gender":
        if "Gender" not in df.columns or "Age" not in df.columns:
            return None
        gender, age_range = row["match"].split(":")
        low, high = map(int, age_range.split("-"))
        temp = df[df["Gender"].astype(str).str.strip().str.upper() == gender.upper()].copy()
        age_num = pd.to_numeric(temp["Age"], errors="coerce")
        return temp[(age_num >= low) & (age_num <= high)]

    return None


def supplier_counts(df, supplier_match):
    """Split a dataframe slice into counts per supplier column, using the match rules for that country."""
    if df is None:
        return {col: 0 for col in supplier_match}

    counts = {}
    if "Supplier Group" not in df.columns:
        # Can't split -> put everything in the first defined column
        first_col = list(supplier_match.keys())[0]
        counts = {col: 0 for col in supplier_match}
        counts[first_col] = len(df)
        return counts

    remainder_cols = []
    assigned_mask = pd.Series(False, index=df.index)
    for col, matches in supplier_match.items():
        if matches is None:
            remainder_cols.append(col)
            continue
        mask = df["Supplier Group"].astype(str).str.strip().str.lower().isin([m.lower() for m in matches])
        counts[col] = int(mask.sum())
        assigned_mask = assigned_mask | mask

    remainder_count = int((~assigned_mask).sum())
    for col in remainder_cols:
        counts[col] = remainder_count  # if there were >1 remainder cols this would double count; configs above use only one

    return counts


def build_report(project_df, config):
    supplier_cols = config["supplier_columns"]
    supplier_match = config["supplier_match"]
    rows = config["rows"]

    collected_by_label = {}  # label -> {supplier_col: count, "TOTAL": count}
    display_rows = []
    labels = []

    for row in rows:
        labels.append(row["label"])
        if row["type"] == "Total":
            # collected computed after the loop from sum_of; placeholder now
            collected_by_label[row["label"]] = None
            display_rows.append(row)
            continue

        matched_df = filter_for_row(project_df, row)
        counts = supplier_counts(matched_df, supplier_match)
        counts["TOTAL"] = sum(counts.values())
        collected_by_label[row["label"]] = counts
        display_rows.append(row)

    # resolve Total rows now that all non-total rows are computed
    for row in rows:
        if row["type"] == "Total":
            summed = {col: 0 for col in supplier_cols}
            summed["TOTAL"] = 0
            for src_label in row.get("sum_of", []):
                src = collected_by_label.get(src_label)
                if src:
                    for col in supplier_cols:
                        summed[col] += src.get(col, 0)
                    summed["TOTAL"] += src.get("TOTAL", 0)
            collected_by_label[row["label"]] = summed

    # assemble final dataframe
    col_tuples = [("TOTAL", "Target"), ("TOTAL", "Collected")]
    for col in supplier_cols:
        col_tuples.append((col, "Target"))
        col_tuples.append((col, "Collected"))
        col_tuples.append((col, "Pending"))
    columns = pd.MultiIndex.from_tuples(col_tuples)

    data = []
    for row in rows:
        target_total = row.get("target_total")
        collected = collected_by_label[row["label"]]
        line = [target_total, collected["TOTAL"]]
        for col in supplier_cols:
            tgt = row["targets"].get(col)
            got = collected.get(col, 0)
            pending = (tgt - got) if tgt is not None else None
            line.extend([tgt if tgt is not None else "–", got, pending if pending is not None else "–"])
        data.append(line)

    report_df = pd.DataFrame(data, index=labels, columns=columns)
    return report_df


def highlight_collected(val):
    return "background-color: #FFFF99; color: black;"


def render_country_tab(country_name):
    st.markdown(f"### {country_name} — Quota Performance Summary")

    raw = raw_df_by_country.get(country_name, pd.DataFrame())
    if raw.empty:
        st.warning(f"No raw data loaded for {country_name}. Check its Google Drive file ID at the top of the script.")
        return

    if "Survey Country" in raw.columns:
        country_df = raw[raw["Survey Country"].astype(str).str.strip().str.lower() == country_name.lower()]
    else:
        country_df = raw  # single-country file with no Survey Country column -- assume it's all this country

    # --- per-country filters (own widget keys so India/Pakistan don't collide) ---
    with st.expander("🔍 Filter Parameters", expanded=False):
        available_langs = list(country_df["Survey Language"].unique()) if "Survey Language" in country_df.columns else []
        selected_langs = st.multiselect("Language", available_langs, default=available_langs, key=f"lang_{country_name}")
        filtered = country_df[country_df["Survey Language"].isin(selected_langs)] if selected_langs else country_df

        available_projects = list(filtered["Project Name"].unique()) if "Project Name" in filtered.columns else []
        selected_projects = st.multiselect("Project Name", available_projects, default=available_projects, key=f"proj_{country_name}")
        project_df = filtered[filtered["Project Name"].isin(selected_projects)] if selected_projects else filtered

    n_projects = project_df["Project Name"].nunique() if "Project Name" in project_df.columns else 0
    st.markdown(f"📊 **Currently analyzing {len(project_df)} completes across {n_projects} project(s) for {country_name}**")

    if project_df.empty:
        st.info(f"No rows found for {country_name} with the current filters.")
        return

    config = COUNTRY_CONFIGS[country_name]
    report_df = build_report(project_df, config)

    collected_cols = [(col, "Collected") for col in config["supplier_columns"]]
    styled = report_df.style.map(highlight_collected, subset=collected_cols)
    st.dataframe(styled, use_container_width=True, height=750)


# =========================================================================
# 5. RENDER TABS
# =========================================================================
tab_india, tab_pakistan = st.tabs(["🇮🇳 India", "🇵🇰 Pakistan"])

with tab_india:
    render_country_tab("India")

with tab_pakistan:
    render_country_tab("Pakistan")
