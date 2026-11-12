import streamlit as st
import pandas as pd
import urllib.parse
from datetime import datetime, timedelta

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="Hockey Team Balancer")

if "shuffle_seed" not in st.session_state:
    st.session_state.shuffle_seed = 0

# --- 2. CACHED DATA LOADER ---
@st.cache_data
def load_excel_data(uploaded_file):
    try:
        # sheet_name=None reads ALL sheets; header=1 means Row 2 is the header
        return pd.read_excel(uploaded_file, sheet_name=None, header=1), None
    except Exception as e:
        return None, str(e)

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
            team_a_list.append(player := players.iloc[i])
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

def get_birthday_message(players_df, bday_col):
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

    # Preserve tie-break shuffle order while sorting by Reg/Spare then Score
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
st.write("Upload your Excel player sheet.")

uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls', 'xlsm'])

if uploaded_file is not None:
    all_sheets, error_msg = load_excel_data(uploaded_file)

    if error_msg:
        st.error(f"❌ Error reading file: {error_msg}")
        st.warning("Troubleshooting:\n1. Ensure file is not open in Excel.\n2. Ensure file is not password protected.")
        st.stop()

    all_sheet_names = list(all_sheets.keys())
    master_sheet_name = next((s for s in all_sheet_names if "master" in s.lower()), None)

    # Strictly exclude Master, Instructions, and reference tabs
    block_list = ["master", "instructions", "instruction", "reference", "setup"]
    valid_sheets = [s for s in all_sheet_names if not any(b in s.strip().lower() for b in block_list)]

    if not valid_sheets:
        st.error("No valid game sheets found (Hidden: Master, Instructions).")
        st.stop()

    selected_sheet = st.selectbox("Select Roster:", valid_sheets)
    df = all_sheets[selected_sheet].copy()
    is_monday = "monday" in selected_sheet.strip().lower()
    is_wednesday = "wednesday" in selected_sheet.strip().lower()

    # --- NORMALIZE COLUMNS ---
    col_name = find_col_case_insensitive(df, ['name', 'full name', 'fullname'])
    col_first = find_col_case_insensitive(df, ['first_name', 'first name', 'firstname'])
    col_last = find_col_case_insensitive(df, ['last_name', 'last name', 'lastname'])
    col_avail = find_col_case_insensitive(df, ['availability', 'avail'])
    col_choice = find_col_case_insensitive(df, ['1st choice', '1stchoice', 'position'])
    col_score = find_col_case_insensitive(df, ['score', 'rating', 'skill'])

    if not (col_name or (col_first and col_last)):
        st.error(f"Missing Name columns in '{selected_sheet}'.")
        st.stop()
    if not col_avail or not col_score or not col_choice:
        st.error(f"Missing columns (Availability, Score, or 1st Choice) in '{selected_sheet}'.")
        st.stop()

    # Build Full Name and filter out empty template rows
    if col_name:
        df['Full Name'] = df[col_name].fillna('').astype(str).str.strip()
        if col_first and col_last:
            fallback_names = (
                df[col_first].fillna('').astype(str).str.strip() + ' ' +
                df[col_last].fillna('').astype(str).str.strip()
            ).str.strip()
            df['Full Name'] = df['Full Name'].where(df['Full Name'] != '', fallback_names)
    else:
        df['Full Name'] = (
            df[col_first].fillna('').astype(str).str.strip() + ' ' +
            df[col_last].fillna('').astype(str).str.strip()
        ).str.strip()

    df = df[df['Full Name'] != ''].copy()

    # Map existing and newly added columns
    col_2nd = find_col_case_insensitive(df, ['2nd choice', '2ndchoice'])
    col_email = find_col_case_insensitive(df, ['email', 'e-mail'])
    col_reg = find_col_case_insensitive(df, ['reg / spare', 'reg/spare', 'status'])
    col_notes = find_col_case_insensitive(df, ['notes', 'note'])
    col_phone = find_col_case_insensitive(df, ['phone', 'phone number'])
    col_days = find_col_case_insensitive(df, ['days', 'day'])

    df['Availability'] = df[col_avail].fillna('').astype(str).str.strip().str.upper()
    df['1st Choice'] = df[col_choice].fillna('').astype(str).str.strip().str.upper()
    df['2nd Choice'] = df[col_2nd].fillna('').astype(str).str.strip().str.upper() if col_2nd else ''
    df['Score'] = pd.to_numeric(df[col_score], errors='coerce').fillna(0)
    df['Email'] = df[col_email].fillna('').astype(str).str.strip() if col_email else ''
    df['Reg / Spare'] = df[col_reg].fillna('S').astype(str).str.strip().str.upper() if col_reg else 'R'
    df['Notes'] = df[col_notes].fillna('').astype(str).str.strip() if col_notes else ''
    df['Phone'] = df[col_phone].fillna('').astype(str).str.strip() if col_phone else ''
    df['Days'] = df[col_days].fillna('').astype(str).str.strip() if col_days else ''

    # --- DAY-SPECIFIC POSITION OVERRIDE ---
    # Gary Fera regularly plays Defense on Wednesdays
    if is_wednesday:
        gary_mask = df['Full Name'].apply(clean_name_key) == clean_name_key("Gary Fera")
        if gary_mask.any():
            df.loc[gary_mask, '1st Choice'] = 'D'
            df.loc[gary_mask, '2nd Choice'] = 'F'

    # --- MASTER LOOKUP (BIRTHDAYS) ---
    bday_col_name = 'B-day'
    birthday_source = df

    if master_sheet_name:
        try:
            df_master = all_sheets[master_sheet_name].copy()
            m_name = find_col_case_insensitive(df_master, ['name', 'full name'])
            m_first = find_col_case_insensitive(df_master, ['first_name', 'first name'])
            m_last = find_col_case_insensitive(df_master, ['last_name', 'last name'])
            m_bday = find_col_case_insensitive(df_master, ['b-day', 'bday', 'birthday', 'dob'])

            if m_name:
                df_master['Full Name'] = df_master[m_name].fillna('').astype(str).str.strip()
                if m_first and m_last:
                    m_fallback = (
                        df_master[m_first].fillna('').astype(str).str.strip() + ' ' +
                        df_master[m_last].fillna('').astype(str).str.strip()
                    ).str.strip()
                    df_master['Full Name'] = df_master['Full Name'].where(df_master['Full Name'] != '', m_fallback)
            elif m_first and m_last:
                df_master['Full Name'] = (
                    df_master[m_first].fillna('').astype(str).str.strip() + ' ' +
                    df_master[m_last].fillna('').astype(str).str.strip()
                ).str.strip()
            else:
                df_master['Full Name'] = ''

            df_master = df_master[df_master['Full Name'] != ''].copy()
            df_master['Key'] = df_master['Full Name'].apply(clean_name_key)

            if m_bday:
                df_master[bday_col_name] = pd.to_datetime(df_master[m_bday], errors='coerce')
                birthday_source = df_master

                df_master_unique = df_master.drop_duplicates(subset=['Key'])
                bday_map = df_master_unique.set_index('Key')[bday_col_name].to_dict()
                df['Key'] = df['Full Name'].apply(clean_name_key)
                df[bday_col_name] = df['Key'].map(bday_map)
            else:
                st.warning("Master sheet found but missing 'B-day' column.")
        except Exception as e:
            st.warning(f"Master sheet error: {e}")
    else:
        daily_bday = find_col_case_insensitive(df, ['b-day', 'bday', 'birthday'])
        if daily_bday:
            bday_col_name = daily_bday
            df[bday_col_name] = pd.to_datetime(df[bday_col_name], errors='coerce')

    # --- FILTER AVAILABLE PLAYERS ---
    available = df[df['Availability'].str.startswith('Y')].copy()
    if available.empty:
        st.error("No players marked 'Y' or 'Yes'.")
        st.stop()

    # --- TARGETS: ALWAYS BALANCE DEFENSE TO 4 PER TEAM (8 TOTAL), REMAINDER AS FORWARDS ---
    total = len(available)
    TARGET_D = 8        # 4 Defensemen per team

    missing_pos = available[~available['1st Choice'].isin(['F', 'D'])]
    if not missing_pos.empty:
        st.warning(
            f"⚠️ Missing 1st Choice position for: **{', '.join(missing_pos['Full Name'])}** "
            "(assigned as Forward; check Master sheet if needed)."
        )

    # --- BIRTHDAY CHECK ---
    bday_msg, bday_names = get_birthday_message(birthday_source, bday_col_name)
    with st.expander("🎂 Birthday Debugger", expanded=True):
        today = datetime.now()
        s_week = today + timedelta(days=1) if today.weekday() == 6 else today - timedelta(days=today.weekday())
        e_week = s_week + timedelta(days=6)
        st.write(f"**Scanning Window:** {s_week.strftime('%a %b %d')} — {e_week.strftime('%a %b %d')}")
        if bday_names:
            st.success(f"**Birthdays Found:** {', '.join(bday_names)}")
        else:
            st.info("No birthdays found this week (checked Master list).")
        st.caption(f"Checked {len(birthday_source)} players in Master list.")

    noted_players = available[available['Notes'] != '']
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
        st.info(f"Balanced Defense (capped at 4 per team): Moved {len(conv)} D -> F: **{', '.join(conv['Full Name'])}**")

    elif len(pool_d) % 2 == 1 and len(pool_d) > 4:
        candidates_2nd = pool_d[pool_d['2nd Choice'] == 'F']
        conv = candidates_2nd.head(1) if not candidates_2nd.empty else pool_d.tail(1)
        pool_f = pd.concat([pool_f, conv])
        pool_d = pool_d.drop(conv.index)
        st.info(f"Evened Defense across teams ({len(pool_d)//2} per team): Moved 1 D -> F: **{', '.join(conv['Full Name'])}**")

    if not pool_other.empty:
        pool_f = pd.concat([pool_f, pool_other])

    st.info(
        f"**Strategy ({selected_sheet}):** {total} players -> "
        f"**{len(pool_d)} Defense** ({len(pool_d)//2} per team) & **{len(pool_f)} Forwards**"
    )

    # --- BUILD ROSTERS ---
    ta, tb, cuts_d, cuts_f, rival_logs = build_rosters(pool_d, pool_f, is_wednesday, allow_wed_d_pair=False)

    if is_wednesday:
        common_cnt = min(len(ta), len(tb))
        score_gap = abs(get_top_n_score(ta, common_cnt) - get_top_n_score(tb, common_cnt))
        if score_gap > 1.5:
            ta_alt, tb_alt, cuts_d_alt, cuts_f_alt, rival_logs_alt = build_rosters(
                pool_d, pool_f, is_wednesday, allow_wed_d_pair=True
            )
            alt_gap = abs(get_top_n_score(ta_alt, common_cnt) - get_top_n_score(tb_alt, common_cnt))
            if alt_gap < score_gap:
                ta, tb, cuts_d, cuts_f, rival_logs = ta_alt, tb_alt, cuts_d_alt, cuts_f_alt, rival_logs_alt

    # --- UI ---
    if st.button("🔀 Shuffle Teams Again"):
        st.session_state.shuffle_seed += 1
        st.rerun()

    if rival_logs:
        st.divider()
        for r in rival_logs:
            st.success(f"⚖️ {r}")

    common_count = min(len(ta), len(tb))
    total_a = ta['Score'].sum()
    total_b = tb['Score'].sum()
    fair_a = get_top_n_score(ta, common_count)
    fair_b = get_top_n_score(tb, common_count)

    c1, c2 = st.columns(2)
    with c1:
        st.header("🔴 Red Team")
        st.write(f"**Total Score:** {total_a:.1f}")
        if len(ta) != len(tb):
            st.write(f"**Top {common_count} Score:** {fair_a:.1f}")
        st.write(f"Players: {len(ta)} ({len(ta[ta['Position'] == 'D'])} D / {len(ta[ta['Position'] == 'F'])} F)")
        st.dataframe(ta[['Full Name', 'Position']], hide_index=True)
    with c2:
        st.header("⚪ White Team")
        st.write(f"**Total Score:** {total_b:.1f}")
        if len(ta) != len(tb):
            st.write(f"**Top {common_count} Score:** {fair_b:.1f}")
        st.write(f"Players: {len(tb)} ({len(tb[tb['Position'] == 'D'])} D / {len(tb[tb['Position'] == 'F'])} F)")
        st.dataframe(tb[['Full Name', 'Position']], hide_index=True)

    if not cuts_d.empty or not cuts_f.empty:
        st.error(f"🚫 Cuts: {len(cuts_d)} D, {len(cuts_f)} F")
        if not cuts_d.empty:
            st.write(f"D: {', '.join(cuts_d['Full Name'])}")
        if not cuts_f.empty:
            st.write(f"F: {', '.join(cuts_f['Full Name'])}")

    st.divider()
    all_p = pd.concat([ta, tb])
    if not all_p.empty:
        recipients = [e for e in all_p['Email'].unique() if pd.notna(e) and str(e).strip()]
        bcc = ",".join(recipients)
        body = (
            f"{bday_msg}Hello everyone,\n\n"
            f"Here are the rosters for the upcoming game:\n\n"
            f"{format_team_list(ta, 'RED TEAM')}\n"
            f"{format_team_list(tb, 'WHITE TEAM')}\n"
            f"Keep your sticks on the ice!"
        )

        st.text_area("Email Draft", value=body, height=300)
        subj = urllib.parse.quote(f"Bendo Hockey Lineups - {selected_sheet}")
        safe_body = urllib.parse.quote(body)
        safe_bcc = urllib.parse.quote(bcc)

        b1, b2, b3 = st.columns(3)
        b1.link_button("📱 Default App", f"mailto:?bcc={safe_bcc}&subject={subj}&body={safe_body}")
        b2.link_button("💻 Gmail Web", f"https://mail.google.com/mail/?view=cm&fs=1&bcc={safe_bcc}&su={subj}&body={safe_body}")
        b3.link_button("🍎 iOS Gmail", f"googlegmail:///co?bcc={safe_bcc}&subject={subj}&body={safe_body}")
