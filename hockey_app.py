import streamlit as st
import pandas as pd
import urllib.parse
import urllib.request
import io
import re
from datetime import datetime, timedelta

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="Hockey Team Balancer", layout="wide")

# PASTE YOUR GOOGLE SHEET URL BETWEEN THE QUOTES BELOW:
DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/11bgbSoRAAL79hYniKkyI4w0wMOAgSSFe/edit?usp=sharing&ouid=114162045734133679348&rtpof=true&sd=true"

if "shuffle_seed" not in st.session_state:
    st.session_state.shuffle_seed = 0

# --- 2. CACHED GOOGLE SHEET LOADER ---
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

    if c_score:
        df["Score"] = pd.to_numeric(
            df[c_score], errors="coerce"
        ).fillna(0.0)
    else:
        df["Score"] = 0.0

    if c_email:
        df["Email"] = df[c_email].fillna("").astype(str).str.strip()
    else:
        df["Email"] = ""

    if c_reg:
        df["Reg / Spare"] = (
            df[c_reg].fillna("S").astype(str).str.strip().str.upper()
        )
    else:
        df["Reg / Spare"] = "S"

    if c_notes:
        df["Notes"] = df[c_notes].fillna("").astype(str).str.strip()
    else:
        df["Notes"] = ""

    if c_phone:
        df["Phone"] = df[c_phone].fillna("").astype(str).str.strip()
    else:
        df["Phone"] = ""

    if c_days:
        df["Days"] = df[c_days].fillna("").astype(str).str.strip()
    else:
        df["Days"] = ""

    if c_bday:
        df["B-day"] = pd.to_datetime(df[c_bday], errors="coerce")
    else:
        df["B-day"] = pd.NaT

    return df

def build_summary_table(all_sheets, valid_sheets, master_sheet_name):
    players_dict = {}

    if master_sheet_name and master_sheet_name in all_sheets:
        df_m = extract_clean_sheet_df(all_sheets[master_sheet_name])
        for _, row in df_m.iterrows():
            k = row["Key"]
            if k not in players_dict:
                reg_val = (
                    row["Reg / Spare"]
                    if row["Reg / Spare"] in ["R", "S"]
                    else "S"
                )
                entry = {
                    "Key": k,
                    "Full Name": row["Full Name"],
                    "Reg / Spare": reg_val,
                    "1st Choice": row["1st Choice"],
                    "2nd Choice": row["2nd Choice"],
                    "Score": float(row["Score"]),
                    "Email": row["Email"],
                    "Phone": row["Phone"],
                    "Notes": row["Notes"],
                    "B-day": row["B-day"],
                }
                for d in valid_sheets:
                    entry[d] = False
                players_dict[k] = entry

    for day in valid_sheets:
        df_d = extract_clean_sheet_df(all_sheets[day])
        for _, row in df_d.iterrows():
            k = row["Key"]
            is_avail = row["Availability"].startswith("Y")
            if k not in players_dict:
                reg_val = (
                    row["Reg / Spare"]
                    if row["Reg / Spare"] in ["R", "S"]
                    else "S"
                )
                entry = {
                    "Key": k,
                    "Full Name": row["Full Name"],
                    "Reg / Spare": reg_val,
                    "1st Choice": row["1st Choice"],
                    "2nd Choice": row["2nd Choice"],
                    "Score": float(row["Score"]),
                    "Email": row["Email"],
                    "Phone": row["Phone"],
                    "Notes": row["Notes"],
                    "B-day": row["B-day"],
                }
                for d in valid_sheets:
                    entry[d] = False
                players_dict[k] = entry
            else:
                fill_cols = [
                    "1st Choice", "2nd Choice", "Email", "Phone", "Notes"
                ]
                for col in fill_cols:
                    if not players_dict[k][col] and row[col]:
                        players_dict[k][col] = row[col]
                if players_dict[k]["Score"] == 0 and row["Score"] > 0:
                    players_dict[k]["Score"] = float(row["Score"])

            players_dict[k][day] = is_avail

    summary_df = pd.DataFrame(list(players_dict.values()))
    if not summary_df.empty:
        summary_df["Status_Sort"] = summary_df["Reg / Spare"].apply(
            lambda x: 0 if x == "R" else 1
        )
        summary_df = (
            summary_df.sort_values(by=["Status_Sort", "Full Name"])
            .drop(columns=["Status_Sort"])
            .reset_index(drop=True)
        )
    return summary_df

def get_birthday_message(players_df, bday_col="B-day"):
    if not bday_col or bday_col not in players_df.columns:
        return "", []

    today = datetime.now()
    if today.weekday() == 6:
        start_of_week = today + timedelta(days=1)
    else:
        start_of_week = today - timedelta(days=today.weekday())

    start_of_week = start_of_week.replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    end_of_week = start_of_week + timedelta(days=6)
    end_of_week = end_of_week.replace(
        hour=23, minute=59, second=59, microsecond=999999
    )

    celebrants = []
    work_df = players_df.copy()
    work_df[bday_col] = pd.to_datetime(
        work_df[bday_col], errors="coerce"
    )

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
                    candidates.append(
                        bday.replace(year=y, month=3, day=1)
                    )

            if any(start_of_week <= c <= end_of_week for c in candidates):
                raw_name = row.get("Full Name", row.get("Name", ""))
                name = str(raw_name).strip()
                valid = name and name.lower() != "nan"
                if valid and name not in celebrants:
                    celebrants.append(name)
        except Exception:
            continue

    if not celebrants:
        return "", []

    if len(celebrants) > 2:
        joined_head = ", ".join(celebrants[:-1])
        names_str = f"{joined_head} and {celebrants[-1]}"
    else:
        names_str = " and ".join(celebrants)

    if len(celebrants) == 1:
        verb = "is"
        noun = "birthday"
    else:
        verb = "are"
        noun = "birthdays"

    msg = (
        f"🎉 Congratulations to {names_str} who {verb} "
        f"celebrating their {noun} this week!\n\n"
    )
    return msg, celebrants

def build_rosters(pool_d, pool_f, is_wednesday, allow_wed_d_pair=False):
    pd_work = pool_d.copy()
    pf_work = pool_f.copy()
    pre_a, pre_b, rival_logs = [], [], []
    TARGET_D = 8
    MAX_PLAYERS = 30

    def extract(nm, pd_p, pf_p):
        key = clean_name_key(nm)
        pd_p["K"] = pd_p["Full Name"].apply(clean_name_key)
        pf_p["K"] = pf_p["Full Name"].apply(clean_name_key)
        md = pd_p[pd_p["K"] == key]
        mf = pf_p[pf_p["K"] == key]
        if not md.empty:
            row = md.iloc[0].copy()
            row["Assigned_Pos"] = "D"
            return row, pd_p.drop(md.index), pf_p
        if not mf.empty:
            row = mf.iloc[0].copy()
            row["Assigned_Pos"] = "F"
            return row, pd_p, pf_p.drop(mf.index)
        return None, pd_p, pf_p

    def return_to_pool(obj, pd_p, pf_p):
        if obj is not None:
            if obj["Assigned_Pos"] == "D":
                pd_p = pd.concat([pd_p, obj.to_frame().T])
            else:
                pf_p = pd.concat([pf_p, obj.to_frame().T])
        return pd_p, pf_p

    rivals = [
        ("Mike Tonietto", "Jamie Devin"),
        ("Mark Hicks", "Gary Fera"),
    ]
    pidx = 0
    for p1, p2 in rivals:
        o1, pd_work, pf_work = extract(p1, pd_work, pf_work)
        o2, pd_work, pf_work = extract(p2, pd_work, pf_work)

        if o1 is not None and o2 is not None:
            pair_set = {clean_name_key(p1), clean_name_key(p2)}
            hf_set = {
                clean_name_key("Mark Hicks"),
                clean_name_key("Gary Fera"),
            }
            is_hf = pair_set == hf_set
            gary_key = clean_name_key("Gary Fera")
            if clean_name_key(o1["Full Name"]) == gary_key:
                gary_obj = o1
            else:
                gary_obj = o2

            wed_exc = (
                is_wednesday
                and is_hf
                and gary_obj["Assigned_Pos"] == "D"
                and allow_wed_d_pair
            )
            if wed_exc:
                pd_work, pf_work = return_to_pool(o1, pd_work, pf_work)
                pd_work, pf_work = return_to_pool(o2, pd_work, pf_work)
                rival_logs.append(
                    "Wednesday Exception: Allowed Mark Hicks & Gary Fera "
                    "(on Defence) into open draft for team balance"
                )
            else:
                srt = sorted(
                    [o1, o2], key=lambda x: x["Score"], reverse=True
                )
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
        if "K" in df_x.columns:
            del df_x["K"]

    pd_work = pd_work.sort_values(
        by=["Status_Rank", "Score"],
        ascending=[True, False],
        kind="mergesort",
    )
    pf_work = pf_work.sort_values(
        by=["Status_Rank", "Score"],
        ascending=[True, False],
        kind="mergesort",
    )

    pre_d_a = len([p for p in pre_a if p["Assigned_Pos"] == "D"])
    pre_d_b = len([p for p in pre_b if p["Assigned_Pos"] == "D"])
    pre_f_a = len([p for p in pre_a if p["Assigned_Pos"] == "F"])
    pre_f_b = len([p for p in pre_b if p["Assigned_Pos"] == "F"])

    total_pre_d = pre_d_a + pre_d_b
    total_pre_f = pre_f_a + pre_f_b

    avail_d_total = len(pd_work) + total_pre_d
    rem_d = max(0, min(TARGET_D, avail_d_total) - total_pre_d)
    sel_d = pd_work.head(rem_d).copy()
    cuts_d = pd_work.iloc[len(sel_d):].copy()

    max_f = MAX_PLAYERS - (len(sel_d) + total_pre_d)
    rem_f = max(0, max_f - total_pre_f)
    sel_f = pf_work.head(rem_f).copy()
    cuts_f = pf_work.iloc[len(sel_f):].copy()

    sel_d["Position"] = "D"
    sel_f["Position"] = "F"
    for p in pre_a + pre_b:
        p["Position"] = p["Assigned_Pos"]

    da, db = true_snake_draft(sel_d, pre_d_a, pre_d_b)
    fa, fb = true_snake_draft(sel_f, pre_f_a, pre_f_b)

    cols = [
        "Full Name", "Position", "Score", "Email", "Phone", "Notes"
    ]

    def to_df(lst):
        if lst:
            return pd.DataFrame(lst)
        return pd.DataFrame(columns=cols)

    ta = pd.concat([to_df(pre_a), da, fa], ignore_index=True)
    tb = pd.concat([to_df(pre_b), db, fb], ignore_index=True)

    ta = ta.sort_values(by=["Position", "Full Name"]).reset_index(
        drop=True
    )
    tb = tb.sort_values(by=["Position", "Full Name"]).reset_index(
        drop=True
    )

    return ta, tb, cuts_d, cuts_f, rival_logs

# --- 4. MAIN APP ---
st.title("🏒 Hockey Team Generator")

col_url, col_ref = st.columns([4, 1])
with col_url:
    sheet_url = st.text_input(
        "Google Sheets Link:",
        value=DEFAULT_GOOGLE_SHEET_URL,
        placeholder="https://docs.google.com/spreadsheets/d/...",
    )
with col_ref:
    st.write("")
    st.write("")
    if st.button("🔄 Refresh Sheet"):
        load_google_sheet_data.clear()
        if "summary_df" in st.session_state:
            del st.session_state["summary_df"]
        st.rerun()

if "YOUR_SHEET_ID_HERE" in sheet_url or not sheet_url.strip():
    st.info(
        "Paste your Google Sheets URL above (or save it permanently in "
        "`DEFAULT_GOOGLE_SHEET_URL` inside `hockey_app.py`)."
    )
    st.stop()

all_sheets, error_msg = load_google_sheet_data(sheet_url.strip())
source_signature = f"gsheet_{sheet_url.strip()}"

if error_msg:
    st.error(f"❌ {error_msg}")
    st.stop()

if all_sheets is not None:
    all_sheet_names = list(all_sheets.keys())
    master_sheet_name = next(
        (s for s in all_sheet_names if "master" in s.lower()), None
    )

    block_list = [
        "master", "instructions", "instruction", "reference", "setup"
    ]
    valid_sheets = [
        s
        for s in all_sheet_names
        if not any(b in s.strip().lower() for b in block_list)
    ]

    if not valid_sheets:
        st.error(
            "No valid game sheets found (Hidden: Master, Instructions)."
        )
        st.stop()

    need_init = (
        st.session_state.get("source_signature") != source_signature
        or "summary_df" not in st.session_state
    )
    if need_init:
        st.session_state.summary_df = build_summary_table(
            all_sheets, valid_sheets, master_sheet_name
        )
        st.session_state.source_signature = source_signature

    tab_summary, tab_lineups = st.tabs(
        ["📋 Weekly Availability Summary", "🏒 Generate Day Lineups"]
    )

    # --- TAB 1: INTERACTIVE MULTI-DAY SUMMARY CHART ---
    with tab_summary:
        st.subheader("Master Roster & Game-Day Availability")
        st.write(
            "Check or uncheck who is playing on each day below. "
            "Changes update the **Generate Day Lineups** tab."
        )

        metric_cols = st.columns(len(valid_sheets))
        for idx, day in enumerate(valid_sheets):
            count_in = int(st.session_state.summary_df[day].sum())
            metric_cols[idx].metric(
                label=f"{day} Players In", value=count_in
            )

        display_cols = (
            ["Full Name", "Reg / Spare"]
            + valid_sheets
            + ["1st Choice", "2nd Choice", "Score", "Notes"]
        )

        col_config = {
            "Full Name": st.column_config.TextColumn(
                "Player Name", disabled=True
            ),
            "Reg / Spare": st.column_config.SelectboxColumn(
                "Reg / Spare", options=["R", "S"], width="small"
            ),
            "1st Choice": st.column_config.SelectboxColumn(
                "1st Choice", options=["F", "D", ""], width="small"
            ),
            "2nd Choice": st.column_config.SelectboxColumn(
                "2nd Choice", options=["F", "D", "G", ""], width="small"
            ),
            "Score": st.column_config.NumberColumn(
                "Score",
                min_value=0.0,
                max_value=10.0,
                step=0.5,
                width="small",
            ),
            "Notes": st.column_config.TextColumn("Notes"),
        }
        for day in valid_sheets:
            col_config[day] = st.column_config.CheckboxColumn(
                day, default=False, width="small"
            )

        edited_view = st.data_editor(
            st.session_state.summary_df[display_cols],
            column_config=col_config,
            hide_index=True,
            use_container_width=True,
            height=600,
            key="attendance_editor",
        )

        for col in display_cols:
            st.session_state.summary_df[col] = edited_view[col]

        if st.button("↩️ Reset Checkboxes to Sheet Defaults"):
            st.session_state.summary_df = build_summary_table(
                all_sheets, valid_sheets, master_sheet_name
            )
            st.rerun()

    # --- TAB 2: LINEUP GENERATOR FOR SELECTED DAY ---
    with tab_lineups:
        selected_sheet = st.selectbox(
            "Select Game Day to Generate Lineups:", valid_sheets
        )
        is_monday = "monday" in selected_sheet.strip().lower()
        is_wednesday = "wednesday" in selected_sheet.strip().lower()

        working_df = st.session_state.summary_df.copy()
        working_df["1st Choice"] = (
            working_df["1st Choice"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )
        working_df["2nd Choice"] = (
            working_df["2nd Choice"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )
        working_df["Reg / Spare"] = (
            working_df["Reg / Spare"]
            .fillna("S")
            .astype(str)
            .str.strip()
            .str.upper()
        )
        working_df["Score"] = pd.to_numeric(
            working_df["Score"], errors="coerce"
        ).fillna(0.0)

        # Gary Fera regularly plays Defense on Wednesdays
        if is_wednesday:
            gary_mask = (
                working_df["Full Name"].apply(clean_name_key)
                == clean_name_key("Gary Fera")
            )
            if gary_mask.any():
                working_df.loc[gary_mask, "1st Choice"] = "D"
                working_df.loc[gary_mask, "2nd Choice"] = "F"

        available = working_df[
            working_df[selected_sheet] == True
        ].copy()
        if available.empty:
            st.warning(
                f"No players checked in for **{selected_sheet}**. "
                "Use the **Weekly Availability Summary** tab first."
            )
            st.stop()

        total = len(available)
        TARGET_D = 8

        missing_pos = available[
            ~available["1st Choice"].isin(["F", "D"])
        ]
        if not missing_pos.empty:
            missing_names = ", ".join(missing_pos["Full Name"])
            st.warning(
                f"⚠️ Missing 1st Choice for: **{missing_names}** "
                "(assigned as Forward; edit in Summary tab if needed)."
            )

        # --- BIRTHDAY CHECK ---
        bday_msg, bday_names = get_birthday_message(
            working_df, "B-day"
        )
        with st.expander("🎂 Birthday Debugger", expanded=True):
            today = datetime.now()
            if today.weekday() == 6:
                s_week = today + timedelta(days=1)
            else:
                s_week = today - timedelta(days=today.weekday())
            e_week = s_week + timedelta(days=6)
            s_str = s_week.strftime("%a %b %d")
            e_str = e_week.strftime("%a %b %d")
            st.write(f"**Scanning Window:** {s_str} — {e_str}")
            if bday_names:
                st.success(
                    f"**Birthdays Found:** {', '.join(bday_names)}"
                )
            else:
                st.info(
                    "No birthdays found this week (checked Master)."
                )
            st.caption(
                f"Checked {len(working_df)} total players in roster."
            )

        notes_series = (
            available["Notes"].fillna("").astype(str).str.strip()
        )
        noted_players = available[notes_series != ""]
        if not noted_players.empty:
            label = f"📝 Active Player Notes ({len(noted_players)})"
            with st.expander(label, expanded=False):
                for _, nrow in noted_players.iterrows():
                    st.write(
                        f"- **{nrow['Full Name']}**: {nrow['Notes']}"
                    )

        # --- SHUFFLE TIE-BREAKS & SORT POOLS ---
        available = available.sample(
            frac=1, random_state=st.session_state.shuffle_seed
        ).reset_index(drop=True)
        available["Status_Rank"] = available["Reg / Spare"].apply(
            lambda x: 0 if x == "R" else 1
        )
        available = available.sort_values(
            by=["Status_Rank", "Score"],
            ascending=[True, False],
            kind="mergesort",
        )

        pool_d = available[available["1st Choice"] == "D"].copy()
        pool_f = available[available["1st Choice"] == "F"].copy()
        pool_other = available[
            ~available["1st Choice"].isin(["D", "F"])
        ].copy()

        # --- POSITIONAL BALANCER LOGIC (Target 4 D per team = 8 D) ---
        if len(pool_d) < TARGET_D:
            needed = TARGET_D - len(pool_d)
            candidates = pool_f[pool_f["2nd Choice"] == "D"].copy()
            if is_monday and not candidates.empty:
                candidates["Is_Gary"] = (
                    candidates["Full Name"].apply(clean_name_key)
                    == clean_name_key("Gary Fera")
                )
                candidates = candidates.sort_values(
                    by=["Is_Gary", "Status_Rank", "Score"],
                    ascending=[True, True, False],
                    kind="mergesort",
                ).drop(columns=["Is_Gary"])

            if not candidates.empty:
                conv = candidates.head(needed)
                pool_d = pd.concat([pool_d, conv])
                pool_f = pool_f.drop(conv.index)
                moved = ", ".join(conv["Full Name"])
                st.info(
                    f"Balanced Defense (4 per team): Moved "
                    f"{len(conv)} F -> D: **{moved}**"
                )

        if len(pool_d) > TARGET_D:
            excess = len(pool_d) - TARGET_D
            c_2nd = pool_d[pool_d["2nd Choice"] == "F"]
            if len(c_2nd) >= excess:
                conv = c_2nd.head(excess)
            else:
                rem = excess - len(c_2nd)
                others = pool_d.drop(c_2nd.index).tail(rem)
                conv = pd.concat([c_2nd, others])
            pool_f = pd.concat([pool_f, conv])
            pool_d = pool_d.drop(conv.index)
            moved = ", ".join(conv["Full Name"])
            st.info(
                f"Balanced Defense (capped at 4 per team): Moved "
                f"{len(conv)} D -> F: **{moved}**"
            )

        elif len(pool_d) % 2 == 1 and len(pool_d) > 4:
            c_2nd = pool_d[pool_d["2nd Choice"] == "F"]
            if not c_2nd.empty:
                conv = c_2nd.head(1)
            else:
                conv = pool_d.tail(1)
            pool_f = pd.concat([pool_f, conv])
            pool_d = pool_d.drop(conv.index)
            moved = ", ".join(conv["Full Name"])
            d_each = len(pool_d) // 2
            st.info(
                f"Evened Defense ({d_each} per team): Moved "
                f"1 D -> F: **{moved}**"
            )

        if not pool_other.empty:
            pool_f = pd.concat([pool_f, pool_other])

        d_each = len(pool_d) // 2
        st.info(
            f"**Strategy ({selected_sheet}):** {total} players -> "
            f"**{len(pool_d)} Defense** ({d_each} per team) & "
            f"**{len(pool_f)} Forwards**"
        )

        # --- BUILD ROSTERS ---
        ta, tb, cuts_d, cuts_f, rival_logs = build_rosters(
            pool_d, pool_f, is_wednesday, allow_wed_d_pair=False
        )

        if is_wednesday:
            common_cnt = min(len(ta), len(tb))
            score_a = get_top_n_score(ta, common_cnt)
            score_b = get_top_n_score(tb, common_cnt)
            score_gap = abs(score_a - score_b)
            if score_gap > 1.5:
                alt_res = build_rosters(
                    pool_d, pool_f, is_wednesday, allow_wed_d_pair=True
                )
                ta_alt, tb_alt, cd_alt, cf_alt, rl_alt = alt_res
                alt_a = get_top_n_score(ta_alt, common_cnt)
                alt_b = get_top_n_score(tb_alt, common_cnt)
                alt_gap = abs(alt_a - alt_b)
                if alt_gap < score_gap:
                    ta, tb = ta_alt, tb_alt
                    cuts_d, cuts_f = cd_alt, cf_alt
                    rival_logs = rl_alt

        # --- UI ---
        if st.button("🔀 Shuffle Teams Again"):
            st.session_state.shuffle_seed += 1
            st.rerun()

        if rival_logs:
            st.divider()
            for r in rival_logs:
                st.success(f"⚖️ {r}")

        common_count = min(len(ta), len(tb))
        total_a = ta["Score"].sum()
        total_b = tb["Score"].sum()
        fair_a = get_top_n_score(ta, common_count)
        fair_b = get_top_n_score(tb, common_count)

        c1, c2 = st.columns(2)
        with c1:
            st.header("🔴 Red Team")
            st.write(f"**Total Score:** {total_a:.1f}")
            if len(ta) != len(tb):
                st.write(f"**Top {common_count} Score:** {fair_a:.1f}")
            d_cnt_a = len(ta[ta["Position"] == "D"])
            f_cnt_a = len(ta[ta["Position"] == "F"])
            st.write(
                f"Players: {len(ta)} ({d_cnt_a} D / {f_cnt_a} F)"
            )
            st.dataframe(
                ta[["Full Name", "Position"]],
                hide_index=True,
                use_container_width=True,
            )
        with c2:
            st.header("⚪ White Team")
            st.write(f"**Total Score:** {total_b:.1f}")
            if len(ta) != len(tb):
                st.write(f"**Top {common_count} Score:** {fair_b:.1f}")
            d_cnt_b = len(tb[tb["Position"] == "D"])
            f_cnt_b = len(tb[tb["Position"] == "F"])
            st.write(
                f"Players: {len(tb)} ({d_cnt_b} D / {f_cnt_b} F)"
            )
            st.dataframe(
                tb[["Full Name", "Position"]],
                hide_index=True,
                use_container_width=True,
            )

        if not cuts_d.empty or not cuts_f.empty:
            st.error(f"🚫 Cuts: {len(cuts_d)} D, {len(cuts_f)} F")
            if not cuts_d.empty:
                st.write(f"D: {', '.join(cuts_d['Full Name'])}")
            if not cuts_f.empty:
                st.write(f"F: {', '.join(cuts_f['Full Name'])}")

        st.divider()
        all_p = pd.concat([ta, tb])
        if not all_p.empty:
            recipients = [
                e
                for e in all_p["Email"].unique()
                if pd.notna(e) and str(e).strip()
            ]
            bcc = ",".join(recipients)
            red_txt = format_team_list(ta, "RED TEAM")
            white_txt = format_team_list(tb, "WHITE TEAM")
            body = (
                f"{bday_msg}Hello everyone,\n\n"
                f"Here are the rosters for the upcoming game:\n\n"
                f"{red_txt}\n"
                f"{white_txt}\n"
                f"Keep your sticks on the ice!"
            )

            st.text_area("Email Draft", value=body, height=300)
            subj = urllib.parse.quote(
                f"Bendo Hockey Lineups - {selected_sheet}"
            )
            safe_body = urllib.parse.quote(body)
            safe_bcc = urllib.parse.quote(bcc)

            b1, b2, b3 = st.columns(3)
            b1.link_button(
                "📱 Default App",
                f"mailto:?bcc={safe_bcc}&subject={subj}&body={safe_body}",
            )
            b2.link_button(
                "💻 Gmail Web",
                f"https://mail.google.com/mail/?view=cm&fs=1"
                f"&bcc={safe_bcc}&su={subj}&body={safe_body}",
            )
            b3.link_button(
                "🍎 iOS Gmail",
                f"googlegmail:///co?bcc={safe_bcc}"
                f"&subject={subj}&body={safe_body}",
            )
