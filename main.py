import streamlit as st
import pandas as pd
import os
import re
import io
import time
import openpyxl

# ==================================================
# PAGE CONFIG
# ==================================================
st.set_page_config(
    page_title="Fleet Compliance Dashboard",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==================================================
# CSS
# ==================================================
st.markdown("""
<style>
:root { --primary-color: #0ea5e9 !important; }

html, body, [class*="css"] {
    font-family: 'Inter', 'Segoe UI', -apple-system, sans-serif;
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header[data-testid="stHeader"] {
    background: transparent !important;
    box-shadow: none !important;
}
button[data-testid="baseButton-headerNoPadding"] {
    visibility: visible !important;
    opacity: 1 !important;
    color: #0f172a !important;
}

.block-container {
    padding-top: 2rem !important;
    padding-bottom: 0.5rem !important;
    max-width: 100% !important;
}
section.main > div { padding-bottom: 0 !important; }
div[data-testid="stAppViewContainer"] > .main { padding-bottom: 0 !important; }
.element-container:has(div[data-testid="stDataFrame"]) { margin-bottom: 0 !important; }
div[data-testid="stDataFrame"] { margin-bottom: 0 !important; }

[data-testid="stMetric"] {
    background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}
[data-testid="stMetric"]:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
}
[data-testid="stMetricLabel"] {
    font-size: 13px !important;
    font-weight: 600;
    color: #64748b !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
[data-testid="stMetricValue"] {
    font-size: 28px !important;
    font-weight: 700;
    color: #0f172a !important;
}

div[data-testid="stDialog"] > div:first-child {
    border-radius: 16px !important;
    box-shadow: 0 20px 60px rgba(0,0,0,0.15) !important;
    border: 1px solid #e2e8f0 !important;
    background: #ffffff !important;
    padding: 8px 4px !important;
}
div[data-testid="stDialog"] h2 {
    font-size: 20px !important;
    font-weight: 700 !important;
    color: #0f172a !important;
    padding-bottom: 12px;
    border-bottom: 1px solid #f1f5f9;
    margin-bottom: 16px !important;
}
div[data-testid="stDialog"] h3 {
    font-size: 15px !important;
    font-weight: 700 !important;
    color: #334155 !important;
    text-transform: uppercase;
    letter-spacing: 0.6px;
    margin-bottom: 16px !important;
    padding-bottom: 6px;
    border-bottom: 2px solid #0ea5e9;
    display: inline-block;
}
div[data-testid="stDialog"] input,
div[data-testid="stDialog"] textarea {
    background-color: #f8fafc !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 8px !important;
    color: #0f172a !important;
    font-size: 14px !important;
    font-weight: 500 !important;
    padding: 10px 14px !important;
    -webkit-text-fill-color: #0f172a !important;
    opacity: 1 !important;
}
div[data-testid="stDialog"] label {
    font-size: 11px !important;
    font-weight: 600 !important;
    color: #64748b !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 4px !important;
}
div[data-testid="stDialog"] input:disabled {
    color: #0f172a !important;
    -webkit-text-fill-color: #0f172a !important;
    background-color: #f8fafc !important;
    opacity: 1 !important;
    cursor: default !important;
}
div[data-testid="stDialog"] .stButton > button,
div[data-testid="stDialog"] .stLinkButton > a {
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    padding: 10px 16px !important;
    border: 1px solid #e2e8f0 !important;
    background: #ffffff !important;
    color: #0f172a !important;
    transition: all 0.15s ease !important;
    text-align: center !important;
    text-decoration: none !important;
}
div[data-testid="stDialog"] .stButton > button:hover,
div[data-testid="stDialog"] .stLinkButton > a:hover {
    background: #0ea5e9 !important;
    color: #ffffff !important;
    border-color: #0ea5e9 !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(14,165,233,0.3);
}

div[data-testid="stDataFrame"] {
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid #e2e8f0;
}
div[data-testid="stDataFrame"] td:first-child {
    color: #0ea5e9 !important;
    font-weight: 600 !important;
}
div[data-testid="stDataFrame"] input[type="checkbox"] {
    accent-color: #0ea5e9 !important;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #f8fafc 0%, #ffffff 60%, #f8fafc 100%);
    border-right: 1px solid #e2e8f0;
}
section[data-testid="stSidebar"] h3 {
    font-size: 18px !important;
    font-weight: 800 !important;
    color: #0f172a !important;
    letter-spacing: -0.4px;
    margin-top: 8px !important;
    margin-bottom: 20px !important;
    padding-bottom: 12px;
    border-bottom: 2px solid #0ea5e9;
    display: inline-block;
}
section[data-testid="stSidebar"] label {
    font-size: 11px !important;
    font-weight: 700 !important;
    color: #64748b !important;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    margin-bottom: 6px !important;
}
section[data-testid="stSidebar"] div[role="radiogroup"] {
    display: flex;
    gap: 8px;
    margin-top: 4px;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label {
    flex: 1;
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 10px 8px !important;
    text-align: center;
    cursor: pointer;
    transition: all 0.18s ease;
    font-size: 12px !important;
    font-weight: 700 !important;
    color: #334155 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.5px !important;
    margin: 0 !important;
    white-space: nowrap !important;
    overflow: hidden !important;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
    border-color: #0ea5e9 !important;
    background: #f0f9ff !important;
    color: #0284c7 !important;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked),
section[data-testid="stSidebar"] div[role="radiogroup"] > label[data-checked="true"] {
    background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%) !important;
    border-color: #0ea5e9 !important;
    color: #ffffff !important;
    box-shadow: 0 4px 12px rgba(14,165,233,0.3) !important;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child,
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child * {
    display: none !important;
    width: 0 !important;
    height: 0 !important;
    margin: 0 !important;
}
section[data-testid="stSidebar"] input[type="radio"] {
    display: none !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 10px !important;
    transition: all 0.18s ease;
    font-weight: 500;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] > div:hover {
    border-color: #0ea5e9 !important;
    box-shadow: 0 0 0 3px rgba(14,165,233,0.1);
}
section[data-testid="stSidebar"] input[type="text"] {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 10px !important;
    color: #0f172a !important;
    font-size: 14px !important;
    padding: 10px 14px !important;
    -webkit-text-fill-color: #0f172a !important;
}
section[data-testid="stSidebar"] input[type="text"]:focus {
    border-color: #0ea5e9 !important;
    box-shadow: 0 0 0 3px rgba(14,165,233,0.15) !important;
}
section[data-testid="stSidebar"] .stButton > button {
    background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%) !important;
    border: none !important;
    color: #ffffff !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 13px !important;
    padding: 12px 16px !important;
    letter-spacing: 0.3px;
    transition: all 0.2s ease !important;
    box-shadow: 0 4px 12px rgba(14,165,233,0.25);
}
section[data-testid="stSidebar"] .stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(14,165,233,0.4) !important;
    background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important;
}
section[data-testid="stSidebar"] hr {
    border-color: #e2e8f0 !important;
    margin: 20px 0 !important;
}
section[data-testid="stSidebar"] img { margin-bottom: 4px; }

h1 {
    font-size: 28px !important;
    font-weight: 800 !important;
    color: #0f172a !important;
    letter-spacing: -0.5px;
}
button[kind="primary"],
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%) !important;
    border-color: #0ea5e9 !important;
    color: #ffffff !important;
}
</style>
""", unsafe_allow_html=True)


# ==================================================
# CONFIG
# ==================================================
EXCEL_FILE = "Steelworks_Fleet_Compliance.xlsm"

COMPLIANCE_SHEETS = {
    "SWPE": "SWPE",
    "SWIN": "SWIN",
}

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
FILE_PATH = os.path.join(BASE_DIR, EXCEL_FILE)

# Column letters in Excel
RC_COLUMN_LETTER = "H"

# SharePoint search fallback
SHAREPOINT_SEARCH_BASE = (
    "https://steelworkspower-my.sharepoint.com/personal/"
    "utkarsh_kashyap_steelworks_in/_layouts/15/onedrive.aspx?q="
)


# ==================================================
# SESSION STATE
# ==================================================
if "dialog_open" not in st.session_state:
    st.session_state.dialog_open = False
if "dialog_row_idx" not in st.session_state:
    st.session_state.dialog_row_idx = None
if "table_reset_counter" not in st.session_state:
    st.session_state.table_reset_counter = 0
if "search_query" not in st.session_state:
    st.session_state.search_query = ""


# ==================================================
# LOAD HELPERS
# ==================================================
@st.cache_data(ttl=60)
def load_sheet(sheet):
    if not os.path.exists(FILE_PATH):
        st.error(f"❌ File not found: {FILE_PATH}")
        st.info(f"📁 Files here: {os.listdir(BASE_DIR)}")
        st.stop()

    for attempt in range(3):
        try:
            with open(FILE_PATH, "rb") as f:
                data = f.read()
            df = pd.read_excel(
                io.BytesIO(data),
                sheet_name=sheet,
                header=0,
                engine="openpyxl",
            )
            df.columns = df.columns.str.strip()
            return df
        except PermissionError:
            time.sleep(1)
        except Exception as e:
            st.error(f"❌ Error reading sheet '{sheet}': {e}")
            st.stop()

    st.error("❌ File is locked. Close Excel and click Rerun.")
    st.stop()


@st.cache_data(ttl=60)
def load_hyperlinks(sheet):
    if not os.path.exists(FILE_PATH):
        return {}
    try:
        with open(FILE_PATH, "rb") as f:
            data = f.read()
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=False, keep_vba=True)
        ws = wb[sheet]

        links = {}
        for row in ws.iter_rows():
            for cell in row:
                url = None
                if cell.hyperlink:
                    url = cell.hyperlink.target
                elif isinstance(cell.value, str) and cell.value.upper().startswith("=HYPERLINK"):
                    try:
                        url = cell.value.split('"')[1]
                    except Exception:
                        pass
                if url:
                    links.setdefault(cell.row, {})[cell.column_letter] = url
        return links
    except Exception as e:
        st.warning(f"Could not read hyperlinks: {e}")
        return {}


def clean_days_left(series):
    def parse(v):
        if pd.isna(v): return None
        if isinstance(v, (int, float)): return int(v)
        s = str(v).strip().upper()
        if "EXPIRED" in s:       return -1
        if s == "LTT":           return 99999
        if s == "ACTIVE":        return 9999
        if s == "EXPIRING SOON": return 15
        m = re.search(r"-?\d+", s)
        return int(m.group()) if m else None
    return series.apply(parse)


def status_icon(days):
    if days is None or (isinstance(days, float) and pd.isna(days)):
        return "⚪ N/A"
    if not isinstance(days, (int, float)):
        s = str(days).strip().upper()
        if "EXPIRED" in s: return "🔴 EXPIRED"
        if "LTT" in s:     return "🟢 LTT"
        if "ACTIVE" in s:  return "🟢 ACTIVE"
        m = re.search(r"-?\d+", s)
        if not m: return "⚪ N/A"
        days = int(m.group())
    d = int(days)
    if d < 0:   return "🔴 EXPIRED"
    if d <= 30: return "🟡 EXPIRING"
    return "🟢 ACTIVE"


# ==================================================
# DOC CONFIG
# ==================================================
DOC_CONFIG = {
    "Insurance": {
        "expiry": "Insurance End Date",
        "days":   "Insurance Days Left",
        "link":   "Insurance Document Link",
        "col":    "Q",
    },
    "Fitness": {
        "expiry": "Fitness Expiry Date",
        "days":   "Fitness Days Left",
        "link":   "Fitness Document Link",
        "col":    "V",
    },
    "MV Tax": {
        "expiry": "MV Tax Expiry Date",
        "days":   "MV Tax Days Left",
        "link":   "MV Tax Document Link",
        "col":    "AA",
    },
    "Permit": {
        "expiry": "Permit Expiry Date",
        "days":   "Permit Days Left",
        "link":   "Permit Document Link",
        "col":    "AE",
    },
    "TP": {
        "expiry": "TP Expiry Date",
        "days":   "TP Days Left",
        "link":   "TP Document Link",
        "col":    "AL",
    },
}


# ==================================================
# UNIVERSAL DOC URL RESOLVER
# ==================================================
def get_doc_url(doc_label, col_letter, hyperlinks, excel_row, vehicle, veh_col):
    # 1. From pre-loaded hyperlinks
    url = hyperlinks.get(excel_row, {}).get(col_letter)
    if url:
        return url

    # 2. Re-read fresh for native hyperlink
    try:
        with open(FILE_PATH, "rb") as f:
            _data = f.read()
        _wb = openpyxl.load_workbook(io.BytesIO(_data), data_only=False, keep_vba=True)
        for _sheet_name in COMPLIANCE_SHEETS.values():
            if _sheet_name in _wb.sheetnames:
                _ws = _wb[_sheet_name]
                cell = _ws[f"{col_letter}{excel_row}"]
                if cell.hyperlink and cell.hyperlink.target:
                    return cell.hyperlink.target
    except Exception:
        pass

    # 3. SharePoint search fallback
    try:
        veh_no = str(vehicle[veh_col]).strip().replace(" ", "%20")
        return f"{SHAREPOINT_SEARCH_BASE}{doc_label}%20{veh_no}"
    except Exception:
        return None


# ==================================================
# SIDEBAR — FILTERS + SEARCH
# ==================================================
with st.sidebar:
    logo_path = os.path.join(BASE_DIR, "SteelworksLogo.png")
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    st.markdown("### 🔍  Filters")

    search_query = st.text_input(
        "🔎 Search Vehicle",
        value=st.session_state.search_query,
        placeholder="Enter vehicle no. or name...",
        key="search_input",
    )
    st.session_state.search_query = search_query

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    entity = st.radio("Select Entity", ["SWPE", "SWIN"], horizontal=True, key="entity_filter")

    df_all     = load_sheet(COMPLIANCE_SHEETS[entity])
    hyperlinks = load_hyperlinks(COMPLIANCE_SHEETS[entity])

    if search_query.strip():
        mask = (
            df_all.iloc[:, 0].astype(str).str.contains(search_query, case=False, na=False) |
            df_all.iloc[:, 1].astype(str).str.contains(search_query, case=False, na=False)
        )
        df_all = df_all[mask]

    doc_type = st.selectbox("Document Type", list(DOC_CONFIG.keys()), key="doc_filter")
    cfg = DOC_CONFIG[doc_type]

    if cfg["expiry"] in df_all.columns:
        df_all["_Expiry"]   = pd.to_datetime(df_all[cfg["expiry"]], errors="coerce")
        df_all["_DaysLeft"] = (df_all["_Expiry"] - pd.Timestamp.today()).dt.days
    else:
        df_all["_Expiry"]   = pd.NaT
        df_all["_DaysLeft"] = None

    if df_all["_DaysLeft"].isna().all() and cfg["days"] in df_all.columns:
        df_all["_DaysLeft"] = clean_days_left(df_all[cfg["days"]])

    df_all["_DaysLeft"] = pd.to_numeric(df_all["_DaysLeft"], errors="coerce")

    df_all["_Year"]  = df_all["_Expiry"].dt.year
    df_all["_Month"] = df_all["_Expiry"].dt.month

    years = ["All"] + sorted(df_all["_Year"].dropna().unique().astype(int).tolist())
    year  = st.selectbox("Year", years, key="year_filter")

    month_names = {1:"January", 2:"February", 3:"March", 4:"April", 5:"May", 6:"June",
                   7:"July", 8:"August", 9:"September", 10:"October", 11:"November", 12:"December"}
    month_sel = st.selectbox("Month", ["All"] + list(month_names.values()), key="month_filter")
    month = 0 if month_sel == "All" else [k for k, v in month_names.items() if v == month_sel][0]

    df_filtered = df_all.copy()
    if year != "All":
        df_filtered = df_filtered[df_filtered["_Year"] == year]
    if month != 0:
        df_filtered = df_filtered[df_filtered["_Month"] == month]

    st.markdown(f"""
    <div style="
        background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
        border: 1px solid #bae6fd;
        border-radius: 10px;
        padding: 12px 16px;
        margin-top: 12px;
        text-align: center;
    ">
        <div style="font-size: 10px; color:#0369a1; text-transform:uppercase; letter-spacing:1px; font-weight:700;">
            Results
        </div>
        <div style="font-size: 24px; color:#0284c7; font-weight:800; letter-spacing:-0.5px; margin-top:2px;">
            {len(df_filtered)}
        </div>
        <div style="font-size: 11px; color:#64748b; margin-top:2px;">vehicles shown</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    st.divider()

    if st.button("🔄  Reset", use_container_width=True):
        st.cache_data.clear()
        st.session_state.dialog_open       = False
        st.session_state.dialog_row_idx    = None
        st.session_state.table_reset_counter += 1
        st.session_state.search_query      = ""
        st.rerun()


# ==================================================
# HERO HEADER
# ==================================================
st.markdown(f"""
<div style="
    background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
    border-radius: 14px;
    padding: 24px 28px;
    color: white;
    margin-bottom: 20px;
    box-shadow: 0 4px 16px rgba(15,23,42,0.25);
">
    <div style="
        font-size: 12px;
        font-weight: 600;
        letter-spacing: 1.5px;
        opacity: 0.7;
        text-transform: uppercase;
        margin-bottom: 6px;
    ">Fleet Compliance</div>
    <div style="
        font-size: 30px;
        font-weight: 800;
        letter-spacing: -0.8px;
    ">{entity} Dashboard</div>
</div>
""", unsafe_allow_html=True)


# ==================================================
# KPI CARDS
# ==================================================
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Vehicles", len(df_filtered))
c2.metric("🟢 Active",        int((df_filtered["_DaysLeft"] > 30).sum()))
c3.metric("🟡 Expiring Soon", int(((df_filtered["_DaysLeft"] >= 0) & (df_filtered["_DaysLeft"] <= 30)).sum()))
c4.metric("🔴 Expired",       int((df_filtered["_DaysLeft"] < 0).sum()))

st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
st.subheader(f"📋 {doc_type} Records")


# ==================================================
# MAIN TABLE
# ==================================================
if len(df_filtered) == 0:
    st.warning("⚠️ No vehicles match the current filters. Try adjusting the search or filters.")
    st.stop()

VEH_COL  = df_filtered.columns[0]
NAME_COL = df_filtered.columns[1]

display_cols = [VEH_COL, NAME_COL, cfg["expiry"], "_DaysLeft"]
available    = [c for c in display_cols if c in df_filtered.columns]

display_df = df_filtered[available].copy()
display_df.columns = ["Vehicle No", "Vehicle Name", "Expiry Date", "Days Left"][:len(available)]
display_df["Status"] = display_df["Days Left"].apply(status_icon)
display_df = display_df[["Vehicle No", "Vehicle Name", "Expiry Date", "Days Left", "Status"]]

event = st.dataframe(
    display_df,
    use_container_width=True,
    hide_index=True,
    on_select="rerun",
    selection_mode="single-row",
    key=f"main_table_{st.session_state.table_reset_counter}",
    column_config={
        "Days Left":   st.column_config.NumberColumn("Days Left", format="%d"),
        "Expiry Date": st.column_config.DateColumn("Expiry Date", format="DD-MMM-YYYY"),
    },
)

if event.selection.rows:
    st.session_state.dialog_row_idx = event.selection.rows[0]
    st.session_state.dialog_open    = True


# ==================================================
# VEHICLE DETAILS DIALOG
# ==================================================
if st.session_state.dialog_open and st.session_state.dialog_row_idx is not None:
    idx = st.session_state.dialog_row_idx

    if not isinstance(idx, int) or idx < 0 or idx >= len(df_filtered):
        st.session_state.dialog_open    = False
        st.session_state.dialog_row_idx = None
        st.rerun()

    vehicle   = df_filtered.iloc[idx]
    excel_row = idx + 2

    @st.dialog(
        f"🚛  {vehicle.get(NAME_COL, '')}  •  {vehicle[VEH_COL]}",
        width="large",
    )
    def show_vehicle_details():

        st.markdown(f"""
        <div style="
            background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%);
            border-radius: 12px;
            padding: 20px 24px;
            margin-bottom: 24px;
            color: white;
            box-shadow: 0 4px 14px rgba(14,165,233,0.3);
        ">
            <div style="
                font-size: 11px;
                font-weight: 600;
                opacity: 0.85;
                text-transform: uppercase;
                letter-spacing: 1.2px;
                margin-bottom: 4px;
            ">Vehicle</div>
            <div style="
                font-size: 24px;
                font-weight: 700;
                letter-spacing: -0.5px;
            ">{vehicle[VEH_COL]}</div>
            <div style="
                font-size: 14px;
                opacity: 0.95;
                margin-top: 4px;
            ">{vehicle.get(NAME_COL, '')}</div>
        </div>
        """, unsafe_allow_html=True)

        left, right = st.columns([1.7, 1])

        with left:
            st.markdown("### 📝 Vehicle Details")
            info = [
                ("RTO Location",           "RTO Location"),
                ("Model No.",              "Model No."),
                ("Chassis No.",            "Chassis No."),
                ("Engine No",              "Engine No"),
                ("Model Year",             "Model Year"),
                ("Insurance Company",      "Insurance Company"),
                ("Policy No.",             "Policy No."),
                ("Insured Value (IDV)",    "Insured Value (IDV)"),
                ("Insurance Amount",       "Insurance Amount"),
                ("Fully Compliant Vehicle","Fully Compliant Vehicle"),
            ]
            c1, c2 = st.columns(2)
            for i, (label, col) in enumerate(info):
                if col in vehicle.index and pd.notna(vehicle[col]):
                    val = vehicle[col]
                    if isinstance(val, float):
                        if col in ("Insured Value (IDV)", "Insurance Amount"):
                            val = f"₹ {val:,.0f}"
                        else:
                            val = f"{val:g}"
                    with (c1 if i % 2 == 0 else c2):
                        st.text_input(label, str(val), disabled=True, key=f"fld_{i}_{idx}")

        with right:
            st.markdown("### 📄 Documents")

            # ---------- RC ----------
            rc_url = get_doc_url("RC", RC_COLUMN_LETTER, hyperlinks, excel_row, vehicle, VEH_COL)
            if rc_url:
                st.link_button("📋   View RC", rc_url, use_container_width=True)
            else:
                st.button("❌   RC — No file", disabled=True,
                          use_container_width=True, key=f"nodoc_rc_{idx}")

            # ---------- All other documents ----------
            for doc_name, doc_cfg in DOC_CONFIG.items():
                doc_url = get_doc_url(
                    doc_name,
                    doc_cfg["col"],
                    hyperlinks,
                    excel_row,
                    vehicle,
                    VEH_COL,
                )
                if doc_url:
                    st.link_button(f"📄   View {doc_name}", doc_url,
                                   use_container_width=True)
                else:
                    st.button(f"❌   {doc_name} — No file",
                              disabled=True,
                              use_container_width=True,
                              key=f"nodoc_{doc_name}_{idx}")

        st.markdown("""
        <div style="
            text-align:center;
            color:#94a3b8;
            font-size:11px;
            padding:16px 0 4px 0;
            letter-spacing:0.3px;
        ">
            Click outside or press <b>ESC</b> to close
        </div>
        """, unsafe_allow_html=True)

    show_vehicle_details()


# ==================================================
# EXPIRING SOON
# ==================================================
st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
st.divider()
st.subheader("⚠️ Expiring in 30 Days")

soon = df_filtered[(df_filtered["_DaysLeft"] >= 0) & (df_filtered["_DaysLeft"] <= 30)]

if soon.empty:
    st.success("✅ No documents expiring in the next 30 days.")
else:
    soon_cols  = [VEH_COL, NAME_COL, "_DaysLeft"]
    soon_avail = [c for c in soon_cols if c in soon.columns]
    soon_display = soon[soon_avail].copy()
    soon_display.columns = ["Vehicle No", "Vehicle Name", "Days Left"][:len(soon_avail)]
    st.dataframe(soon_display, use_container_width=True, hide_index=True)


# ==================================================
# FOOTER
# ==================================================
st.markdown(f"""
<div style="
    text-align:center;
    color:#94a3b8;
    font-size:12px;
    padding:16px 0 8px 0;
    margin-top:12px;
    border-top:1px solid #f1f5f9;
">
    © {pd.Timestamp.now().year} Steelworks • Fleet Compliance System
</div>
""", unsafe_allow_html=True)