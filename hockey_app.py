import streamlit as st
import pandas as pd
import urllib.parse
import urllib.request
import io
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="Hockey Team Balancer", layout="wide")

# PASTE YOUR GOOGLE SHEET URL BETWEEN THE QUOTES BELOW:
DEFAULT_GOOGLE_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/11bgbSoRAAL79hYniKkyI4w0wMOAgSSFe/edit?usp=sharing&ouid=114162045734133679348&rtpof=true&sd=true"
)

if "shuffle_seed" not in st.session_state:
    st.session_state.shuffle_seed = 0


@st.cache_data(ttl=60)
def load_google_sheet_data(sheet_url):
    try:
        m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not m:
            return None, "Invalid Google Sheets URL."
        sid = m.group(1)
        url = (
            f"https://docs.google.com/spreadsheets/d/{sid}"
            "/export?format=xlsx"
        )
        req = urllib.request.Request(
            url, headers={"User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            raw = resp.read()
        data = pd.read_excel(
            io.BytesIO(raw), sheet_name=None, header=1
        )
        return data, None
    except Exception as e:
        msg = (
            f"Could not read Google Sheet ({e}). Ensure Share access "
            "is set to 'Anyone with the link can view'."
        )
        return None, msg


def true_snake_draft(players, pre_a=0, pre_b=0):
    players = players.reset_index(drop=True)
    mirror = pre_b > pre_a
    start = pre_a + pre_b
    list_a, list_b = [], []

    for i in range(len(players)):
        slot = start + i
        pick_a = (slot % 4 == 0) or (slot % 4 == 3)
        if mirror:
            pick_a = not pick_a
        if pick_a:
            list_a.append(players.iloc[i])
        else:
            list_b.append(players.iloc[i])

    cols = players.columns
    df_a = (
        pd.DataFrame(list_a, columns=cols)
        if list_a
        else pd.DataFrame(columns=cols)
    )
    df_b = (
        pd.DataFrame(list_b, columns=cols)
        if list_b
        else pd.DataFrame(columns=cols)
    )
    return df_a, df_b


def format_team_list(df, team_name):
    if df.empty:
        return f"{team_name} (0 players):\n"
    txt = f"{team_name} ({len(df)} players):\n"
    if "Position" in df.columns and "Full Name" in df.columns:
        s_df = df.sort_values(
            by=["Position", "Full Name"], ascending=[True, True]
        )
        for _, row in s_df.iterrows():
            txt += f"- {row['Full Name']} ({row['Position']})\n"
    return txt


def get_top_n_score(df, n):
    if df.empty or n <= 0:
        return 0
    return df.sort_values(by="Score", ascending=False).head(n)[
        "Score"
    ].sum()


def clean_name_key(name):
    if pd.isna(name):
        return ""
    return str(name).lower().replace(" ", "").strip()


def find_col(df, names):
    if isinstance(names, str):
        names = [names]
    targets = ["".join(t.lower().split()) for t in names]
    for col in df.columns:
        if "".join(str(col).lower().split()) in targets:
            return col
    return None


def extract_clean_sheet_df(raw_df, is_master=False):
    df = raw_df.copy()
    c_name = find_col(df, ["name", "full name", "fullname"])
    c_first = find_col(df, ["first_name", "first name", "firstname"])
    c_last = find_col(df, ["last_name", "last name", "lastname"])

    if c_name:
        df["Full Name"] = (
            df[c_name].fillna("").astype(str).str.strip()
        )
        if c_first and c_last:
            fp = df[c_first].fillna("").astype(str).str.strip()
            lp = df[c_last].fillna("").astype(str).str.strip()
            fb = (fp + " " + lp).str.strip()
            df["Full Name"] = df["Full Name"].where(
                df["Full Name"] != "", fb
            )
    elif c_first and c_last:
        fp = df[c_first].fillna("").astype(str).str.strip()
        lp = df[c_last].fillna("").astype(str).str.strip()
        df["Full Name"] = (fp + " " + lp).str.strip()
    else:
        df["Full Name"] = ""

    df = df[df["Full Name"] != ""].copy()
    df["Key"] = df["Full Name"].apply(clean_name_key)

    c_avail = find_col(df, ["availability", "avail"])
    c_1st = find_col(df, ["1st choice", "1stchoice", "position"])
    c_2nd = find_col(df, ["2nd choice", "2ndchoice"])
    c_score = find_col(df, ["score", "rating", "skill"])
    c_email = find_col(df, ["email", "e-mail"])
    c_reg = find_col(df, ["reg / spare", "reg/spare", "status"])
    c_notes = find_col(df, ["notes", "note"])
    c_phone = find_col(df, ["phone", "phone number"])
    c_days = find_col(df, ["days", "day"])
    c_bday = find_col(df, ["b-day", "bday", "birthday", "dob"])

    df["Availability"] = (
        df[c_avail].fillna("").astype(str).str.strip().str.upper()
        if c_avail
        else "N"
    )
    df["1st Choice"] = (
        df[c_1st].fillna("").astype(str).str.strip().str.upper()
        if c_1st
        else ""
    )
    df["2nd Choice"] = (
        df[c_2nd].fillna("").astype(str).str.strip().str.upper()
        if c_2nd
        else ""
    )
    df["Score"] = (
        pd.to_numeric(df[c_score], errors="coerce").fillna(0.0)
        if c_score
        else 0.0
    )
    df["Email"] = (
        df[c_email].fillna("").astype(str).str.strip()
        if c_email
        else ""
    )
    df["Reg / Spare"] = (
        df[c_reg].fillna("S").astype(str).str.strip().str.upper()
        if c_reg
        else "S"
    )
    df["Notes"] = (
        df[c_notes].fillna("").astype(str).str.strip()
        if c_notes
        else ""
    )
    df["Phone"] = (
        df[c_phone].fillna("").astype(str).str.strip()
        if c_phone
        else ""
    )

    if c_days:
        df["Days"] = (
            df[c_days].fillna("").astype(str).str.strip().str.upper()
        )
    elif is_master and len(df.columns) > 1:
        col_b = df.columns[1]
        df["Days"] = (
            df[col_b].fillna("").astype(str).str.strip().str.upper()
        )
    else:
        df["Days"] = ""

    df["B-day"] = (
        pd.to_datetime(df[c_bday], errors="coerce")
        if c_bday
        else pd.NaT
    )
    return df


def build_summary_table(all_sheets, valid_sheets, master_name):
    players = {}

    if master_name and master_name in all_sheets:
        df_m = extract_clean_sheet_df(
            all_sheets[master_name], is_master=True
        )
        for _, r in df_m.iterrows():
            k = r["Key"]
            if k not in players:
                reg = (
                    r["Reg / Spare"]
                    if r["Reg / Spare"] in ["R", "S"]
                    else "S"
                )
                row_obj = {
                    "Key": k,
                    "Full Name": r["Full Name"],
                    "Reg / Spare": reg,
                    "Days": r["Days"],
                    "1st Choice": r["1st Choice"],
                    "2nd Choice": r["2nd Choice"],
                    "Score": float(r["Score"]),
                    "Email": r["Email"],
                    "Phone": r["Phone"],
                    "Notes": r["Notes"],
                    "B-day": r["B-day"],
                }
                for d in valid_sheets:
                    row_obj[d] = False
                players[k] = row_obj

    for day in valid_sheets:
        df_d = extract_clean_sheet_df(
            all_sheets[day], is_master=False
        )
        for _, r in df_d.iterrows():
            k = r["Key"]
            avail = r["Availability"].startswith("Y")
            if k not in players:
                reg = (
                    r["Reg / Spare"]
                    if r["Reg / Spare"] in ["R", "S"]
                    else "S"
                )
                row_obj = {
                    "Key": k,
                    "Full Name": r["Full Name"],
                    "Reg / Spare": reg,
                    "Days": r["Days"],
                    "1st Choice": r["1st Choice"],
                    "2nd Choice": r["2nd Choice"],
                    "Score": float(r["Score"]),
                    "Email": r["Email"],
                    "Phone": r["Phone"],
                    "Notes": r["Notes"],
                    "B-day": r["B-day"],
                }
                for d in valid_sheets:
                    row_obj[d] = False
                players[k] = row_obj
            else:
                cols = [
                    "Days",
                    "1st Choice",
                    "2nd Choice",
                    "Email",
                    "Phone",
                    "Notes",
                ]
                for c in cols:
                    if not players[k][c] and r[c]:
                        players[k][c] = r[c]
                if players[k]["Score"] == 0 and r["Score"] > 0:
                    players[k]["Score"] = float(r["Score"])
            players[k][day] = avail

    out = pd.DataFrame(list(players.values()))
    if not out.empty:
        out["Sort_R"] = out["Reg / Spare"].apply(
            lambda x: 0 if x == "R" else 1
        )
        out = (
            out.sort_values(by=["Sort_R", "Full Name"])
            .drop(columns=["Sort_R"])
            .reset_index(drop=True)
        )
    return out


def get_birthday_message(players_df, bday_col="B-day"):
    if not bday_col or bday_col not in players_df.columns:
        return "", []

    today = datetime.now()
    if today.weekday() == 6:
        s_wk = today + timedelta(days=1)
    else:
        s_wk = today - timedelta(days=today.weekday())

    s_wk = s_wk.replace(hour=0, minute=0, second=0, microsecond=0)
    e_wk = (s_wk + timedelta(days=6)).replace(
        hour=23, minute=59, second=59, microsecond=999999
    )

    celebrants = []
    w_df = players_df.copy()
    w_df[bday_col] = pd.to_datetime(w_df[bday_col], errors="coerce")

    for _, row in w_df.iterrows():
        bday = row[bday_col]
        if pd.isna(bday):
            continue
        try:
            cands = []
            for y in [today.year, today.year - 1, today.year + 1]:
                try:
                    cands.append(bday.replace(year=y))
                except ValueError:
                    cands.append(
                        bday.replace(year=y, month=3, day=1)
                    )
            if any(s_wk <= c <= e_wk for c in cands):
                raw_n = row.get("Full Name", row.get("Name", ""))
                nm = str(raw_n).strip()
                if nm and nm.lower() != "nan" and nm not in celebrants:
                    celebrants.append(nm)
        except Exception:
            continue

    if not celebrants:
        return "", []

    if len(celebrants) > 2:
        head = ", ".join(celebrants[:-1])
        names_str = f"{head} and {celebrants[-1]}"
    else:
        names_str = " and ".join(celebrants)

    verb = "is" if len(celebrants) == 1 else "are"
    noun = "birthday" if len(celebrants) == 1 else "birthdays"
    msg = (
        f"🎉 Congratulations to {names_str} who {verb} "
        f"celebrating their {noun} this week!\n\n"
    )
    return msg, celebrants


def build_rosters(pool_d, pool_f, is_wed, allow_wed_pair=False):
    pd_w = pool_d.copy()
    pf_w = pool_f.copy()
    pre_a, pre_b, logs = [], [], []
    TARGET_D = 8
    MAX_P = 30

    def extract(nm, pd_p, pf_p):
        k = clean_name_key(nm)
        pd_p["K"] = pd_p["Full Name"].apply(clean_name_key)
        pf_p["K"] = pf_p["Full Name"].apply(clean_name_key)
        md = pd_p[pd_p["K"] == k]
        mf = pf_p[pf_p["K"] == k]
        if not md.empty:
            r = md.iloc[0].copy()
            r["Assigned_Pos"] = "D"
            return r, pd_p.drop(md.index), pf_p
        if not mf.empty:
            r = mf.iloc[0].copy()
            r["Assigned_Pos"] = "F"
            return r, pd_p, pf_p.drop(mf.index)
        return None, pd_p, pf_p

    def put_back(obj, pd_p, pf_p):
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
        o1, pd_w, pf_w = extract(p1, pd_w, pf_w)
        o2, pd_w, pf_w = extract(p2, pd_w, pf_w)

        if o1 is not None and o2 is not None:
            pset = {clean_name_key(p1), clean_name_key(p2)}
            hf = {
                clean_name_key("Mark Hicks"),
                clean_name_key("Gary Fera"),
            }
            gkey = clean_name_key("Gary Fera")
            g_obj = (
                o1 if clean_name_key(o1["Full Name"]) == gkey else o2
            )

            if (
                is_wed
                and pset == hf
                and g_obj["Assigned_Pos"] == "D"
                and allow_wed_pair
            ):
                pd_w, pf_w = put_back(o1, pd_w, pf_w)
                pd_w, pf_w = put_back(o2, pd_w, pf_w)
                logs.append(
                    "Wednesday Exception: Allowed Mark Hicks & Gary "
                    "Fera (on Defence) into open draft for balance"
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
                logs.append(f"Separated {p1} & {p2}")
                pidx += 1
        else:
            pd_w, pf_w = put_back(o1, pd_w, pf_w)
            pd_w, pf_w = put_back(o2, pd_w, pf_w)

    for df_x in [pd_w, pf_w]:
        if "K" in df_x.columns:
            del df_x["K"]

    pd_w = pd_w.sort_values(
        by=["Status_Rank", "Score"],
        ascending=[True, False],
        kind="mergesort",
    )
    pf_w = pf_w.sort_values(
        by=["Status_Rank", "Score"],
        ascending=[True, False],
        kind="mergesort",
    )

    pre_da = len([p for p in pre_a if p["Assigned_Pos"] == "D"])
    pre_db = len([p for p in pre_b if p["Assigned_Pos"] == "D"])
    pre_fa = len([p for p in pre_a if p["Assigned_Pos"] == "F"])
    pre_fb = len([p for p in pre_b if p["Assigned_Pos"] == "F"])

    tot_pre_d = pre_da + pre_db
    tot_pre_f = pre_fa + pre_fb

    rem_d = max(0, min(TARGET_D, len(pd_w) + tot_pre_d) - tot_pre_d)
    sel_d = pd_w.head(rem_d).copy()
    cuts_d = pd_w.iloc[len(sel_d):].copy()

    max_f = MAX_P - (len(sel_d) + tot_pre_d)
    rem_f = max(0, max_f - tot_pre_f)
    sel_f = pf_w.head(rem_f).copy()
    cuts_f = pf_w.iloc[len(sel_f):].copy()

    sel_d["Position"] = "D"
    sel_f["Position"] = "F"
    for p in pre_a + pre_b:
        p["Position"] = p["Assigned_Pos"]

    da, db = true_snake_draft(sel_d, pre_da, pre_db)
    fa, fb = true_snake_draft(sel_f, pre_fa, pre_fb)

    cols = [
        "Full Name", "Position", "Score", "Email", "Phone", "Notes"
    ]

    def to_df(lst):
        return (
            pd.DataFrame(lst)
            if lst
            else pd.DataFrame(columns=cols)
        )

    ta = pd.concat([to_df(pre_a), da, fa], ignore_index=True)
    tb = pd.concat([to_df(pre_b), db, fb], ignore_index=True)
    ta = ta.sort_values(by=["Position", "Full Name"]).reset_index(
        drop=True
    )
    tb = tb.sort_values(by=["Position", "Full Name"]).reset_index(
        drop=True
    )
    return ta, tb, cuts_d, cuts_f, logs


# --- MAIN APP ---
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
        "Paste your Google Sheets URL above (or save it permanently "
        "in `DEFAULT_GOOGLE_SHEET_URL` inside `hockey_app.py`)."
    )
    st.stop()

all_sheets, error_msg = load_google_sheet_data(sheet_url.strip())
sig = f"gsheet_{sheet_url.strip()}"

if error_msg:
