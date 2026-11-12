import streamlit as st
import pandas as pd
import urllib.parse
from datetime import datetime, timedelta

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="Hockey Team Balancer")

# --- 2. CACHED DATA LOADER ---
@st.cache_data
def load_excel_data(uploaded_file):
    try:
        # sheet_name=None reads ALL sheets; header=1 means Row 2 is the header
        return pd.read_excel(uploaded_file, sheet_name=None, header=1), None
    except Exception as e:
        return None, str(e)

# --- 3. HELPER FUNCTIONS ---
def snake_draft(players):
    players = players.reset_index(drop=True)
    team_a_list, team_b_list = [], []
    for i in range(len(players)):
        player = players.iloc[i]
        if i % 4 == 0 or i % 4 == 3:
            team_a_list.append(player)
        else:
            team_b_list.append(player)

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
    if "wednesday" in selected_sheet.strip().lower():
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

    # --- TARGETS ---
    total = len(available)
    BASE_F, BASE_D, MIN_D_CRITICAL = 12, 8, 6

    if total <= 20:
        target_f, target_d = BASE_F, BASE_D
    else:
        ex = total - 20
        add_f = min(ex, 6)
        ex -= add_f
        target_f, target_d = BASE_F + add_f, BASE_D + min(ex, 4)

    st.info(f"**Strategy ({selected_sheet}):** {total} players -> {target_f} F / {target_d} D")

    missing_pos = available[~available['1st Choice'].isin(['F', 'D'])]
    if not missing_pos.empty:
        st.warning(
            f"⚠️ Missing 1st Choice position for: **{', '.join(missing_pos['Full Name'])}** "
            "(check if they are listed on the Master sheet)."
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

    # --- SORT & POOLS ---
    available = available.sample(frac=1).reset_index(drop=True)
    available['Status_Rank'] = available['Reg / Spare'].apply(lambda x: 0 if x == 'R' else 1)
    available = available.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])

    pool_d = available[available['1st Choice'] == 'D'].copy()
    pool_f = available[available['1st Choice'] == 'F'].copy()

    # --- BALANCER LOGIC ---
    if len(pool_d) < MIN_D_CRITICAL:
        needed = MIN_D_CRITICAL - len(pool_d)
        candidates = pool_f[pool_f['2nd Choice'] == 'D']
        if not candidates.empty:
            conv = candidates.head(needed)
            pool_d = pd.concat([pool_d, conv])
            pool_f = pool_f.drop(conv.index)
            st.warning(f"⚠️ Critical D Shortage (<6): Moved {len(conv)} F -> D: **{', '.join(conv['Full Name'])}**")

    if len(pool_f) < target_f:
        needed = target_f - len(pool_f)
        surplus_d = len(pool_d) - MIN_D_CRITICAL
        if surplus_d > 0:
            can_take = min(needed, surplus_d)
            candidates = pool_d[pool_d['2nd Choice'] == 'F']
            if not candidates.empty:
                conv = candidates.head(can_take)
                pool_f = pd.concat([pool_f, conv])
                pool_d = pool_d.drop(conv.index)
                st.info(f"Balanced Forwards: Moved {len(conv)} D -> F: **{', '.join(conv['Full Name'])}**")

    if len(pool_d) < target_d:
        d_shortage = target_d - len(pool_d)
        surplus_f = len(pool_f) - target_f
        if surplus_f > 0:
            amount_to_move = min(d_shortage, surplus_f)
            candidates = pool_f[pool_f['2nd Choice'] == 'D']
            if not candidates.empty:
                conv = candidates.head(amount_to_move)
                pool_d = pd.concat([pool_d, conv])
                pool_f = pool_f.drop(conv.index)
                st.info(f"Balanced Defense: Moved {len(conv)} F -> D: **{', '.join(conv['Full Name'])}**")

    # --- RIVALS ---
    pre_a, pre_b, rival_logs = [], [], []

    def extract(nm, pd_pool, pf_pool):
        key = clean_name_key(nm)
        pd_pool['K'] = pd_pool['Full Name'].apply(clean_name_key)
        pf_pool['K'] = pf_pool['Full Name'].apply(clean_name_key)
        md, mf = pd_pool[pd_pool['K'] == key], pf_pool[pf_pool['K'] == key]
        if not md.empty:
            row = md.iloc[0].copy()
            row['Assigned_Pos'] = 'D'
            return row, pd_pool.drop(md.index), pf_pool
        if not mf.empty:
            row = mf.iloc[0].copy()
            row['Assigned_Pos'] = 'F'
            return row, pd_pool, pf_pool.drop(mf.index)
        return None, pd_pool, pf_pool

    rivals = [("Mike Tonietto", "Jamie Devin"), ("Mark Hicks", "Gary Fera")]
    pidx = 0
    for p1, p2 in rivals:
        o1, pool_d, pool_f = extract(p1, pool_d, pool_f)
        o2, pool_d, pool_f = extract(p2, pool_d, pool_f)
        if o1 is not None and o2 is not None:
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
            if o1 is not None:
                if o1['Assigned_Pos'] == 'D':
                    pool_d = pd.concat([pool_d, o1.to_frame().T])
                else:
                    pool_f = pd.concat([pool_f, o1.to_frame().T])
            if o2 is not None:
                if o2['Assigned_Pos'] == 'D':
                    pool_d = pd.concat([pool_d, o2.to_frame().T])
                else:
                    pool_f = pd.concat([pool_f, o2.to_frame().T])

    # --- DRAFT ---
    for df_x in [pool_d, pool_f]:
        if 'K' in df_x.columns:
            del df_x['K']

    pool_d = pool_d.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])
    pool_f = pool_f.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])

    pd_cnt = len([p for p in pre_a + pre_b if p['Assigned_Pos'] == 'D'])
    pf_cnt = len([p for p in pre_a + pre_b if p['Assigned_Pos'] == 'F'])

    sel_d = pool_d.head(max(0, target_d - pd_cnt)).copy()
    sel_f = pool_f.head(max(0, target_f - pf_cnt)).copy()
    cuts_d = pool_d.iloc[len(sel_d):].copy()
    cuts_f = pool_f.iloc[len(sel_f):].copy()

    sel_d['Position'] = 'D'
    sel_f['Position'] = 'F'
    for p in pre_a + pre_b:
        p['Position'] = p['Assigned_Pos']

    da, db = snake_draft(sel_d)
    fa, fb = snake_draft(sel_f)

    # --- COMBINE ---
    cols = ['Full Name', 'Position', 'Score', 'Email', 'Phone', 'Notes']
    def to_df(lst):
        return pd.DataFrame(lst) if lst else pd.DataFrame(columns=cols)

    ta = pd.concat([to_df(pre_a), da, fa], ignore_index=True)
    tb = pd.concat([to_df(pre_b), db, fb], ignore_index=True)

    ta = ta.sort_values(by=['Position', 'Full Name']).reset_index(drop=True)
    tb = tb.sort_values(by=['Position', 'Full Name']).reset_index(drop=True)

    # --- UI ---
    if st.button("Shuffle"):
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
