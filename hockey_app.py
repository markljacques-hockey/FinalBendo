import streamlit as st
import pandas as pd
import urllib.parse
import urllib.request
import io
import re
from datetime import datetime, timedelta

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="Hockey Team Balancer", layout="wide")

if "shuffle_seed" not in st.session_state:
    st.session_state.shuffle_seed = 0

# --- 2. CACHED DATA LOADERS ---
@st.cache_data
def load_excel_data(uploaded_file):
    try:
        data = pd.read_excel(uploaded_file, sheet_name=None, header=1)
        return data, None
    except Exception as e:
        return None, str(e)

@st.cache_data(ttl=60)
def load_google_sheet_data(sheet_url):
    try:
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not match:
            return None, "Invalid Google Sheets URL."
        sheet_id = match.group(1)
        base = "https://docs.google.com/spreadsheets/d"
        export_url = f"{base}/{sheet_id}/export?format=xlsx"
        hdrs = {"User-Agent": "Mozilla/5.0"}
        req = urllib.request.Request(export_url, headers=hdrs)
        with urllib.request.urlopen(req, timeout=15) as resp:
            xlsx_bytes = resp.read()
        data = pd.read_excel(
            io.BytesIO(xlsx_bytes), sheet_name=None, header=1
        )
        return data, None
    except Exception as e:
        msg = (
            f"Could not read Google Sheet ({e}). Ensure Share access "
            "is set to 'Anyone with the link can view'."
        )
        return None, msg

# --- 3. HELPER FUNCTIONS ---
def true_snake_draft(players, pre_count_a=0, pre_count_b=0):
    players = players.reset_index(drop=True)
    mirror = pre_count_b > pre_count_a
    start_slot = pre_count_a + pre_count_b

    team_a_list, team_b_list = [], []
    for i in range(len(players)):
        slot = start_slot + i
        pick_a = (slot % 4 == 0) or (slot % 4 == 3)
        if mirror:
            pick_a = not pick_a

        if pick_a:
            team_a_list.append(players.iloc[i])
        else:
            team_b_list.append(players.iloc[i])

    cols = players.columns
    if team_a_list:
        df_a = pd.DataFrame(team_a_list, columns=cols)
    else:
        df_a = pd.DataFrame(columns=cols)

    if team_b_list:
        df_b = pd.DataFrame(team_b_list, columns=cols)
    else:
        df_b = pd.DataFrame(columns=cols)

    return df_a, df_b

def format_team_list(df, team_name):
    if df.empty:
        return f"{team_name} (0 players):\n"
    txt = f"{team_name} ({len(df)} players):\n"
    if "Position" in df.columns and "Full Name" in df.columns:
        df_sorted = df.sort_values(
            by=["Position", "Full Name"], ascending=[True, True]
        )
        for _, row in df_sorted.iterrows():
            txt += f"- {row['Full Name']} ({row['Position']})\n"
    return txt

def get_top_n_score(df, n):
    if df.empty or n <= 0:
        return 0
    top_df = df.sort_values(by="Score", ascending=False).head(n)
    return top_df["Score"].sum()

def clean_name_key(name):
    if pd.isna(name):
        return ""
    return str(name).lower().replace(" ", "").strip()

def find_col_case_insensitive(df, target_names):
    if isinstance(target_names, str):
        target_names = [target_names]
    targets = ["".join(t.lower().split()) for t in target_names]
    for col in df.columns:
        clean_col = "".join(str(col).lower().split())
        if clean_col in targets:
            return col
    return None

def extract_clean_sheet_df(raw_df):
    df = raw_df.copy()
    c_name = find_col_case_insensitive(
        df, ["name", "full name", "fullname"]
    )
    c_first = find_col_case_insensitive(
        df, ["first_name", "first name", "firstname"]
    )
    c_last = find_col_case_insensitive(
        df, ["last_name", "last name", "lastname"]
    )

    if c_name:
        df["Full Name"] = df[c_name].fillna("").astype(str).str.strip()
        if c_first and c_last:
            f_part = df[c_first].fillna("").astype(str).str.strip()
            l_part = df[c_last].fillna("").astype(str).str.strip()
            fallback = (f_part + " " + l_part).str.strip()
            df["Full Name"] = df["Full Name"].where(
                df["Full Name"] != "", fallback
            )
    elif c_first and c_last:
        f_part = df[c_first].fillna("").astype(str).str.strip()
        l_part = df[c_last].fillna("").astype(str).str.strip()
        df["Full Name"] = (f_part + " " + l_part).str.strip()
    else:
        df["Full Name"] = ""

    df = df[df["Full Name"] != ""].copy()
    df["Key"] = df["Full Name"].apply(clean_name_key)

    c_avail = find_col_case_insensitive(df, ["availability", "avail"])
    c_1st = find_col_case_insensitive(
        df, ["1st choice", "1stchoice", "position"]
    )
    c_2nd = find_col_case_insensitive(df, ["2nd choice", "2ndchoice"])
    c_score = find_col_case_insensitive(
        df, ["score", "rating", "skill"]
    )
    c_email = find_col_case_insensitive(df, ["email", "e-mail"])
    c_reg = find_col_case_insensitive(
        df, ["reg / spare", "reg/spare", "status"]
    )
    c_notes = find_col_case_insensitive(df, ["notes", "note"])
    c_phone = find_col_case_insensitive(df, ["phone", "phone number"])
    c_days = find_col_case_insensitive(df, ["days", "day"])
    c_bday = find_col_case_insensitive(
        df, ["b-day", "bday", "birthday", "dob"]
    )

    if c_avail:
        df["Availability"] = (
            df[c_avail].fillna("").astype(str).str.strip().str.upper()
        )
    else:
        df["Availability"] = "N"

    if c_1st:
        df["1st Choice"] = (
            df[c_1st].fillna("").astype(str).str.strip().str.upper()
        )
    else:
        df["1st Choice"] = ""

    if c_2nd:
        df["2nd Choice"] = (
            df[c_2nd].fillna("").astype(str).str.strip().str.upper()
        )
    else:
        df["2nd Choice"] = ""
