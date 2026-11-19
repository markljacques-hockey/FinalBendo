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

# --- 2. CACHED DATA LOADERS (EXCEL & GOOGLE SHEETS) ---
@st.cache_data
def load_excel_data(uploaded_file):
    try:
        return pd.read_excel(uploaded_file, sheet_name=None, header=1), None
    except Exception as e:
        return None, str(e)

@st.cache_data(ttl=60)
def load_google_sheet_data(sheet_url):
    try:
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not match:
            return None, "Invalid Google Sheets URL. Please paste the full link from your browser."
        sheet_id = match.group(1)
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
        req = urllib.request.Request(export_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=15) as response:
            xlsx_bytes = response.read()
        return pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=None, header=1), None
    except Exception as e:
        return None, (
            f"Could not read Google Sheet ({e}). "
            "Make sure the Google Sheet's Share settings are set to 'Anyone with the link can view'."
        )

# --- 3. HELPER FUNCTIONS ---
def true_snake_draft(players, pre_count_a=0, pre_count_b=0):
    players = players.reset_index(drop=True)
    mirror = pre_count_b > pre_count_a
    start_slot = pre_count_a + pre_count_b

    team_a_list, team_b_list = [], []
    for i in range(len(players)):
        slot = start_slot + i
        pick_a = (slot % 4 == 0 or slot % 4 == 3)
        if mirror:
            pick_a = not pick_a

        if pick_a:
            team_a_list.append(players.iloc[i])
        else:
            team_b_list.append(players.iloc[i])

    cols = players.columns
    df_a = pd.DataFrame(team_a_list, columns=cols) if team_a_list else pd.DataFrame(columns=cols)
    df_b = pd.DataFrame(team_b_list, columns=cols) if team_b_list else pd.DataFrame(columns=cols)
    return df_a, df_b

def format_team_list(df, team_name):
    if df.empty:
        return f"{team_name} (0 players):\n"
    txt = f"{team_name} ({len(df)} players):\n"
    if 'Position' in df.columns and 'Full Name' in df.columns:
        df_sorted = df.sort_values(by=['Position', 'Full Name'], ascending=[True, True])
        for _, row in df_sorted.iterrows():
            txt += f"- {row['Full Name']} ({row['Position']})\n"
    return txt

def get_top_n_score(df, n):
    if df.empty or n <= 0:
        return 0
    return df.sort_values(by='Score', ascending=False).head(n)['Score'].sum()

def clean_name_key(name):
    if pd.isna(name):
        return ""
    return str(name).lower().replace(" ", "").strip()

def find_col_case_insensitive(df, target_names):
    if isinstance(target_names, str):
        target_names = [target_names]
    normalized_targets = ["".join(t.lower().split()) for t in target_names]
    for col in df.columns:
        clean_col = "".join(str(col).lower().split())
        if clean_col in normalized_targets:
            return col
    return None

def extract_clean_sheet_df(raw_df):
    df = raw_df.copy()
    col_name = find_col_case_insensitive(df, ['name', 'full name', 'fullname'])
    col_first = find_col_case_insensitive(df, ['first_name', 'first name', 'firstname'])
    col_last = find_col_case_insensitive(df, ['last_name', 'last name', 'lastname'])

    if col_name:
        df['Full Name'] = df[col_name].fillna('').astype(str).str.strip()
        if col_first and col_last:
            fallback_names = (
                df[col_first].fillna('').astype(str).str.strip() + ' ' +
                df[col_last].fillna('').astype(str).str.strip()
            ).str.strip()
            df['Full Name'] = df['Full Name'].where(df['Full Name'] != '', fallback_names)
    elif col_first and col_last:
        df['Full Name'] = (
            df[col_first].fillna('').astype(str).str.strip() + ' ' +
            df[col_last].fillna('').astype(str).str.strip()
        ).str.strip()
    else:
        df['Full Name'] = ''

    df = df[df['Full Name'] != ''].copy()
    df['Key'] = df['Full Name'].apply(clean_name_key)

    col_avail = find_col_case_insensitive(df, ['availability', 'avail'])
    col_choice = find_col_case_insensitive(df, ['1st choice', '1stchoice', 'position'])
    col_2nd = find_col_case_insensitive(df, ['2nd choice', '2ndchoice'])
    col_score = find_col_case_insensitive(df, ['score', 'rating', 'skill'])
    col_email = find_col_case_insensitive(df, ['email', 'e-mail'])
    col_reg = find_col_case_insensitive(df, ['reg / spare', 'reg/spare', 'status'])
    col_notes = find_col_case_insensitive(df, ['notes', 'note'])
    col_phone = find_col_case_insensitive(df, ['phone', 'phone number'])
    col_days = find_col_case_insensitive(df, ['days', 'day'])
    col_bday = find_col_case_insensitive(df, ['b-day', 'bday', 'birthday', 'dob'])

    df['Availability'] = df[col_avail].fillna('').astype(str).str.strip().str.upper() if col_avail else 'N'
    df['1st Choice'] = df[col_choice].fillna('').astype(str).str.strip().str.upper() if col_choice else ''
    df['2nd Choice'] = df[col_2nd].fillna('').astype(str).str.strip().str.upper() if col_2nd else ''
    df['Score'] = pd.to_numeric(df[col_score], errors='coerce').fillna(0.0) if col_score else 0.0
    df['Email'] = df[col_email].fillna('').astype(str).str.strip() if col_email else ''
    df['Reg / Spare'] = df[col_reg].fillna('S').astype(str).str.strip().str.upper() if col_reg else 'S'
    df['Notes'] = df[col_notes].fillna('').astype(str).str.strip() if col_notes else ''
    df['Phone'] = df[col_phone].fillna('').astype(str).str.strip() if col_phone else ''
    df['Days'] = df[col_days].fillna('').astype(str).str.strip() if col_days else ''
    df['B-day'] = pd.to_datetime(df[col_bday], errors='coerce') if col_bday else pd.NaT

    return df

def build_summary_table(all_sheets, valid_sheets, master_sheet_name):
    players_dict = {}

    if master_sheet_name and master_sheet_name in all_sheets:
        df_m = extract_clean_sheet_df(all_sheets[master_sheet_name])
        for _, row in df_m.iterrows():
            k = row['Key']
            if k not in players_dict:
                players_dict[k] = {
                    'Key': k,
                    'Full Name': row['Full Name'],
                    'Reg / Spare': row['Reg / Spare'] if row['Reg / Spare'] in ['R', 'S'] else 'S',
                    '1st Choice': row['1st Choice'],
                    '2nd Choice': row['2nd Choice'],
                    'Score': float(row['Score']),
                    'Email': row['Email'],
                    'Phone': row['Phone'],
                    'Notes': row['Notes'],
                    'B-day': row['B-day'],
                    **{day: False for day in valid_sheets}
                }

    for day in valid_sheets:
        df_d = extract_clean_sheet_df(all_sheets[day])
        for _, row in df_d.iterrows():
            k = row['Key']
            is_avail = row['Availability'].startswith('Y')
            if k not in players_dict:
                players_dict[k] = {
                    'Key': k,
                    'Full Name': row['Full Name'],
                    'Reg / Spare': row['Reg / Spare'] if row['Reg / Spare'] in ['R', 'S'] else 'S',
                    '1st Choice': row['1st Choice'],
                    '2nd Choice': row['2nd Choice'],
                    'Score': float(row['Score']),
                    'Email': row['Email'],
                    'Phone': row['Phone'],
                    'Notes': row['Notes'],
                    'B-day': row['B-day'],
                    **{d: False for d in valid_sheets}
                }
            else:
                for col in ['1st Choice', '2nd Choice', 'Email', 'Phone', 'Notes']:
                    if not players_dict[k][col] and row[col]:
                        players_dict[k][col] = row[col]
                if players_dict[k]['Score'] == 0 and row['Score'] > 0:
                    players_dict[k]['Score'] = float(row['Score'])

            players_dict[k][day] = is_avail

    summary_df = pd.DataFrame(list(players_dict.values()))
    if not summary_df.empty:
        summary_df['Status_Sort'] = summary_df['Reg / Spare'].apply(lambda x: 0 if x == 'R' else 1)
        summary_df = (
            summary_df.sort_values(by=['Status_Sort', 'Full Name'])
            .drop(columns=['Status_Sort'])
            .reset_index(drop=True)
        )
    return summary_df

def get_birthday_message(players_df, bday_col='B-day'):
    if not bday_col or bday_col not in players_df.columns:
        return "", []

    today = datetime.now()
    if today.weekday() == 6:
        start_of_week = today + timedelta(days=1)
    else:
        start_of_week = today - timedelta(days=today.weekday())

    start_of_week = start_of_week.replace(hour=0, minute=0, second=0, microsecond=0)
    end_of_week = start_of_week + timedelta(days=6)
    end_of_week = end_of_week.replace(hour=23, minute=59, second=59, microsecond=999999)

    celebrants = []
    work_df = players_df.copy()
    work_df[bday_col] = pd.to_datetime(work_df[bday_col], errors='coerce')

    for _, row in work_df.iterrows():
        bday = row[bday_col]
        if pd.isna(bday):
            continue
        try:
            candidates = []
            years = [today.year, today.year - 1, today.year + 1]
            for y in years:
                try:
                    candidates.append(bday.replace(year=y))
                except ValueError:
                    candidates.append(bday.replace(year=y, month=3, day=1))

            if any(start_of_week <= c <= end_of_week for c in candidates):
                name = str(row.get('Full Name', row.get('Name', ''))).strip()
                if name and name.lower() != 'nan' and name not in celebrants:
                    celebrants.append(name)
        except Exception:
            continue

    if not celebrants:
        return "", []

    names_str = " and ".join(
        [", ".join(celebrants[:-1]), celebrants[-1]] if len(celebrants) > 2 else celebrants
    )
    verb = "is" if len(celebrants) == 1 else "are"
    noun = "birthday" if len(celebrants) == 1 else
