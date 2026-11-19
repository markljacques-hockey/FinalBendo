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
    """
    True snake draft (A, B, B, A, A, B, B, A) adjusted for any pre-assigned players
    at the same position. Ensures top-rated and weakest-rated players at each
    position are separated onto opposite teams rather than stacked together.
    """
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
    """Finds a column name ignoring case and extra internal/external spaces."""
    if isinstance(target_names, str):
        target_names = [target_names]
    normalized_targets = ["".join(t.lower().split()) for t in target_names]
    for col in df.columns:
        clean_col = "".join(str(col).lower().split())
        if clean_col in normalized_targets:
            return col
    return None

def extract_clean_sheet_df(raw_df):
    """Normalizes a single worksheet into standard columns and removes empty rows."""
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
    """
    Combines Master + all daily game sheets into one unified summary table
    with boolean checkboxes for each game day.
    """
    players_dict = {}

    # 1. Seed from Master sheet first so every league player is included
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

    # 2. Merge each game day's roster & availability
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
                # Fill in any missing fields if blank in Master
                for col in ['1st Choice', '2nd Choice', 'Email', 'Phone', 'Notes']:
                    if not players_dict[k][col] and row[col]:
                        players_dict[k][col] = row[col]
                if players_dict[k]['Score'] == 0 and row['Score'] > 0:
                    players_dict[k]['Score'] = float(row['Score'])

            players_dict[k][day] = is_avail

    summary_df = pd.DataFrame(list(players_dict.values()))
    if not summary_df.empty:
        summary_df['Status_Sort'] = summary_df['Reg / Spare'].apply(lambda x: 0 if x == 'R' else 1)
        summary_df = summary_df.sort_values(by=['Status_Sort', 'Full Name']).drop(columns=['Status_Sort']).reset_index(drop=True)
    return summary_df

def get_birthday_message(players_df, bday_col='B-day'):
    if not bday_col or bday_col not in players_df.columns:
        return "", []

    today = datetime.now()

    # SUNDAY LOOKAHEAD LOGIC: If run on Sunday (6), scan upcoming Mon-Sun
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

    names_str = " and ".join([", ".join(celebrants[:-1]), celebrants[-1]] if len(celebrants) > 2 else celebrants)
    verb = "is" if len(celebrants) == 1 else "are"
    noun = "birthday" if len(celebrants) == 1 else "birthdays"
    return f"🎉 Congratulations to {names_str} who {verb} celebrating their {noun} this week!\n\n", celebrants

def build_rosters(pool_d, pool_f, is_wednesday, allow_wed_d_pair=False):
    """Runs rival separation + true snake draft and returns teams, cuts, and notes."""
    pd_work = pool_d.copy()
    pf_work = pool_f.copy()
    pre_a, pre_b, rival_logs = [], [], []
    TARGET_D = 8
    MAX_PLAYERS = 30

    def extract(nm, pd_p, pf_p):
        key = clean_name_key(nm)
        pd_p['K'] = pd_p['Full Name'].apply(clean_name_key)
        pf_p['K'] = pf_p['Full Name'].apply(clean_name_key)
        md, mf = pd_p[pd_p['K'] == key], pf_p[pf_p['K'] == key]
        if not md.empty:
            row = md.iloc[0].copy()
            row['Assigned_Pos'] = 'D'
            return row, pd_p.drop(md.index), pf_p
        if not mf.empty:
            row = mf.iloc[0].copy()
            row['Assigned_Pos'] = 'F'
            return row, pd_p, pf_p.drop(mf.index)
        return None, pd_p, pf_p

    def return_to_pool(obj, pd_p, pf_p):
        if obj is not None:
            if obj['Assigned_Pos'] == 'D':
                pd_p = pd.concat([pd_p, obj.to_frame().T])
            else:
                pf_p = pd.concat([pf_p, obj.to_frame().T])
        return pd_p, pf_p

    rivals = [("Mike Tonietto", "Jamie Devin"), ("Mark Hicks", "Gary Fera")]
    pidx = 0
    for p1, p2 in rivals:
        o1, pd_work, pf_work = extract(p1, pd_work, pf_work)
        o2, pd_work, pf_work = extract(p2, pd_work, pf_work)

        if o1 is not None and o2 is not None:
            is_hicks_fera = {clean_name_key(p1), clean_name_key(p2)} == {clean_name_key("Mark Hicks"), clean_name_key("Gary Fera")}
            gary_obj = o1 if clean_name_key(o1['Full Name']) == clean_name_key("Gary Fera") else o2

            # Wednesday exception: only allowed when Gary is playing Defence AND separation caused a balance issue
            if is_wednesday and is_hicks_fera and gary_obj['Assigned_Pos'] == 'D' and allow_wed_d_pair:
                pd_work, pf_work = return_to_pool(o1, pd_work, pf_work)
                pd_work, pf_work = return_to_pool(o2, pd_work, pf_work)
                rival_logs.append("Wednesday Exception: Allowed Mark Hicks & Gary Fera (on Defence) into open draft for team balance")
            else:
                srt = sorted([o1, o2], key=lambda x: x['Score'], reverse=True)
                if pidx % 2 == 0:
                    pre_a.append(srt[0])
                    pre_b.append(srt[1])
                else:
                    pre_b.append(srt[0])
                    pre_a.append(srt[1])
                rival_logs.append(f"Separated {p1} & {p2}")
                pidx += 1
        else:
            pd_work, pf_work = return_to_pool(o1, pd_work, pf_work)
            pd_work, pf_work = return_to_pool(o2, pd_work, pf_work)

    for df_x in [pd_work, pf_work]:
        if 'K' in df_x.columns:
            del df_x['K']

    pd_work = pd_work.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False], kind='mergesort')
    pf_work = pf_work.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False], kind='mergesort')

    pre_d_a = len([p for p in pre_a if p['Assigned_Pos'] == 'D'])
    pre_d_b = len([p for p in pre_b if p['Assigned_Pos'] == 'D'])
    pre_f_a = len([p for p in pre_a if p['Assigned_Pos'] == 'F'])
    pre_f_b = len([p for p in pre_b if p['Assigned_Pos'] == 'F'])

    total_pre_d = pre_d_a + pre_d_b
    total_pre_f = pre_f_a + pre_f_b

    rem_d = max(0, min(TARGET_D, len(pd_work) + total_pre_d) - total_pre_d)
    sel_d = pd_work.head(rem_d).copy()
    cuts_d = pd_work.iloc[len(sel_d):].copy()

    max_f = MAX_PLAYERS - (len(sel_d) + total_pre_d)
    rem_f = max(0, max_f - total_pre_f)
    sel_f = pf_work.head(rem_f).copy()
    cuts_f = pf_work.iloc[len(sel_f):].copy()

    sel_d['Position'] = 'D'
    sel_f['Position'] = 'F'
    for p in pre_a + pre_b:
        p['Position'] = p['Assigned_Pos']

    da, db = true_snake_draft(sel_d, pre_d_a, pre_d_b)
    fa, fb = true_snake_draft(sel_f, pre_f_a, pre_f_b)

    cols = ['Full Name', 'Position', 'Score', 'Email', 'Phone', 'Notes']
    def to_df(lst):
        return pd.DataFrame(lst) if lst else pd.DataFrame(columns=cols)

    ta = pd.concat([to_df(pre_a), da, fa], ignore_index=True)
    tb = pd.concat([to_df(pre_b), db, fb], ignore_index=True)

    ta = ta.sort_values(by=['Position', 'Full Name']).reset_index(drop=True)
    tb = tb.sort_values(by=['Position', 'Full Name']).reset_index(drop=True)

    return ta, tb, cuts_d, cuts_f, rival_logs

# --- 4. MAIN APP ---
st.title("🏒 Hockey Team Generator")

source_type = st.radio(
    "Choose Roster Source:",
    ["📁 Upload Excel File", "🔗 Google Sheets Link"],
    horizontal=True
)

all_sheets = None
error_msg = None
source_signature = ""

if source_type == "📁 Upload Excel File":
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls', 'xlsm'])
    if uploaded_file is not None:
        all_sheets, error_msg = load_excel_data(uploaded_file)
        source_signature = f"file_{uploaded_file.name}_{uploaded_file.size}"
else:
    col_url, col_ref = st.columns([4, 1])
    with col_url:
        sheet_url = st.text_input(
            "Paste Google Sheets URL:",
            placeholder="https://docs.google.com/spreadsheets/d/..."
        )
    with col_ref:
        st.write("")
        st.write("")
        if st.button("🔄 Refresh"):
            load_google_sheet_data.clear()
            if "summary_df" in st.session_state:
                del st.session_state["summary_df"]
            st.rerun()

    if sheet_url.strip():
        all_sheets, error_msg = load_google_sheet_data(sheet_url.strip())
        source_signature = f"gsheet_{sheet_url.strip()}"

if error_msg:
    st.error(f"❌ {error_msg}")
    st.stop()

if all_sheets is not None:
    all_sheet_names = list(all_sheets.keys())
    master_sheet_name = next((s for s in all_sheet_names if "master" in s.lower()), None)

    block_list = ["master", "instructions", "instruction", "reference", "setup"]
    valid_sheets = [s for s in all_sheet_names if not any(b in s.strip().lower() for b in block_list)]

    if not valid_sheets:
        st.error("No valid game sheets found (Hidden: Master, Instructions).")
        st.stop()

    # Initialize or reload summary table when a new file/link is provided
    if st.session_state.get("source_signature") != source_signature or "summary_df" not in st.session_state:
        st.session_state.summary_df = build_summary_table(all_sheets, valid_sheets, master_sheet_name)
        st.session_state.source_signature = source_signature

    tab_summary, tab_lineups = st.tabs(["📋 Weekly Availability Summary", "🏒 Generate Day Lineups"])

    # --- TAB 1: INTERACTIVE MULTI-DAY SUMMARY CHART ---
    with tab_summary:
        st.subheader("Master Roster & Game-Day Availability")
        st.write("Check or uncheck who is playing on each day below. Any changes automatically update the **Generate Day Lineups** tab.")

        # Quick headcount metrics per day
        metric_cols = st.columns(len(valid_sheets))
        for idx, day in enumerate(valid_sheets):
            count_in = int(st.session_state.summary_df[day].sum())
            metric_cols[idx].metric(label=f"{day} Players In", value=count_in)

        display_cols = ['Full Name', 'Reg / Spare'] + valid_sheets + ['1st Choice', '2nd Choice', 'Score', 'Notes']

        col_config = {
            "Full Name": st.column_config.TextColumn("Player Name", disabled=True),
            "Reg / Spare": st.column_config.SelectboxColumn("Reg / Spare", options=["R", "S"], width="small"),
            "1st Choice": st.column_config.SelectboxColumn("1st Choice", options=["F", "D", ""], width="small"),
            "2nd Choice": st.column_config.SelectboxColumn("2nd Choice", options=["F", "D", "G", ""], width="small"),
            "Score": st.column_config.NumberColumn("Score", min_value=0.0, max_value=10.0, step=0.5, width="small"),
            "Notes": st.column_config.TextColumn("Notes"),
        }
        for day in valid_sheets:
            col_config[day] = st.column_config.CheckboxColumn(day, default=False, width="small")

        edited_view = st.data_editor(
            st.session_state.summary_df[display_cols],
            column_config=col_config,
            hide_index=True,
            use_container_width=True,
            height=600,
            key="attendance_editor"
        )

        # Sync user edits back into session_state.summary_df
        for col in display_cols:
            st.session_state.summary_df[col] = edited_view[col]

        if st.button("↩️ Reset Checkboxes to File Defaults"):
            st.session_state.summary_df = build_summary_table(all_sheets, valid_sheets, master_sheet_name)
            st.rerun()

    # --- TAB 2: LINEUP GENERATOR FOR SELECTED DAY ---
    with tab_lineups:
        selected_sheet = st.selectbox("Select Game Day to Generate Lineups:", valid_sheets)
        is_monday = "monday" in selected_sheet.strip().lower()
        is_wednesday = "wednesday" in selected_sheet.strip().lower()

        working_df = st.session_state.summary_df.copy()
        working_df['1st Choice'] = working_df['1st Choice'].fillna('').astype(str).str.strip().str.upper()
        working_df['2nd Choice'] = working_df['2nd Choice'].fillna('').astype(str).str.strip().str.upper()
        working_df['Reg / Spare'] = working_df['Reg / Spare'].fillna('S').astype(str).str.strip().str.upper()
        working_df['Score'] = pd.to_numeric(working_df['Score'], errors='coerce').fillna(0.0)

        # --- DAY-SPECIFIC POSITION OVERRIDE ---
        # Gary Fera regularly plays Defense on Wednesdays
        if is_wednesday:
            gary_mask = working_df['Full Name'].apply(clean_name_key) == clean_name_key("Gary Fera")
            if gary_mask.any():
                working_df.loc[gary_mask, '1st Choice'] = 'D'
                working_df.loc[gary_mask, '2nd Choice'] = 'F'

        # Filter to players checked 'True' for the chosen day
        available = working_df[working_df[selected_sheet] == True].copy()
        if available.empty:
            st.warning(f"No players are currently checked as available for **{selected_sheet}**. Go to the **Weekly Availability Summary** tab to check players in.")
            st.stop()

        total = len(available)
        TARGET_D = 8  # Always balance to 4 Defensemen per team (8 total)

        missing_pos = available[~available['1st Choice'].isin(['F', 'D'])]
        if not missing_pos.empty:
            st.warning(
                f"⚠️ Missing 1st Choice position for: **{', '.join(missing_pos['Full Name'])}** "
                "(assigned as Forward; you can set their position in the Summary tab)."
            )

        # --- BIRTHDAY CHECK (Checks all players in Master list) ---
        bday_msg, bday_names = get_birthday_message(working_df, 'B-day')
        with st.expander("🎂 Birthday Debugger", expanded=True):
            today = datetime.now()
            s_week = today + timedelta(days=1) if today.weekday() == 6 else today - timedelta(days=today.weekday())
            e_week = s_week + timedelta(days=6)
            st.write(f"**Scanning Window:** {s_week.strftime('%a %b %d')} — {e_week.strftime('%a %b %d')}")
            if bday_names:
                st.success(f"**Birthdays Found:** {', '.join(bday_names)}")
            else:
                st.info("No birthdays found this week (checked Master list).")
            st.caption(f"Checked {len(working_df)} total players in roster.")

        noted_players = available[available['Notes'].fillna('').astype(str).str.strip() != '']
        if not noted_players.empty:
            with st.expander(f"📝 Active Player Notes ({len(noted_players)})", expanded=False):
                for _, nrow in noted_players.iterrows():
                    st.write(f"- **{nrow['Full Name']}**: {nrow['Notes']}")

        # --- SHUFFLE TIE-BREAKS & SORT POOLS ---
        available = available.sample(frac=1, random_state=st.session_state.shuffle_seed).reset_index(drop=True)
        available['Status_Rank'] = available['Reg / Spare'].apply(lambda x: 0 if x == 'R' else 1)
        available = available.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False], kind='mergesort')

        pool_d = available[available['1st Choice'] == 'D'].copy()
        pool_f = available[available['1st Choice'] == 'F'].copy()
        pool_other = available[~available['1st Choice'].isin(['D', 'F'])].copy()

        # --- POSITIONAL BALANCER LOGIC (Target 4 D per team = 8 D total) ---
        if len(pool_d) < TARGET_D:
            needed = TARGET_D - len(pool_d)
            candidates = pool_f[pool_f['2nd Choice'] == 'D'].copy()
            # On Mondays, keep Gary Fera at Forward unless all other F->D options are exhausted
            if is_monday and not candidates.empty:
                candidates['Is_Gary'] = candidates['Full Name'].apply(clean_name_key) == clean_name_key("Gary Fera")
                candidates = candidates.sort_values(
                    by=['Is_Gary', 'Status_Rank', 'Score'], ascending=[True, True, False], kind='mergesort'
                ).drop(columns=['Is_Gary'])

            if not candidates.empty:
                conv = candidates.head(needed)
                pool_d = pd.concat([pool_d, conv])
                pool_f = pool_f.drop(conv.index)
                st.info(f"Balanced Defense (4 per team): Moved {len(conv)} F -> D: **{', '.join(conv['Full Name'])}**")

        if len(pool_d) > TARGET_D:
            excess = len(pool_d) - TARGET_D
            candidates_2nd = pool_d[pool_d['2nd Choice'] == 'F']
            if len(candidates_2nd) >= excess:
                conv = candidates_2nd.head(excess)
            else:
                rem = excess - len(candidates_2nd)
                others = pool_d.drop(candidates_2nd.index).tail(rem)
                conv = pd.concat([candidates_2nd, others])
            pool_f = pd.concat([pool_f, conv])
            pool_d = pool_d.drop(conv.index)
            st.info(f"Balanced Defense (capped at 4
