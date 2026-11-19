import streamlit as st
import pandas as pd
import urllib.parse
import urllib.request
import io
import re
import json
from datetime import datetime, timedelta

try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except ImportError:
    HAS_GSPREAD = False

st.set_page_config(page_title="Hockey Team Balancer", layout="wide")

# PASTE YOUR GOOGLE SHEET URL BETWEEN THE QUOTES BELOW:
DEFAULT_GOOGLE_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/1zUnLUXHHjZpkA-g_3a3gxIn8I0N0ujuvSe3N3BpdPb8/edit?usp=sharing"
)

if "shuffle_seed" not in st.session_state:
    st.session_state.shuffle_seed = 0


def get_gspread_client():
    if not HAS_GSPREAD:
        return None, "gspread library not installed."
    if "gcp_service_account" not in st.secrets:
        return None, "gcp_service_account not found in Streamlit Secrets."
    try:
        raw_secret = st.secrets["gcp_service_account"]
        if isinstance(raw_secret, str):
            creds_dict = json.loads(raw_secret)
        else:
            creds_dict = dict(raw_secret)
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(
            creds_dict, scopes=scopes
        )
        client = gspread.authorize(creds)
        return client, None
    except Exception as e:
        return None, f"Service account auth error: {e}"


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


def save_changes_to_google_sheet(sheet_url, summary_df, valid_sheets):
    client, err = get_gspread_client()
    if err:
        return False, err

    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not m:
        return False, "Invalid Google Sheets URL."
    sid = m.group(1)

    try:
        spreadsheet = client.open_by_key(sid)
        sheet_title = spreadsheet.title
        summary_map = {}
        for _, r in summary_df.iterrows():
            summary_map[r["Key"]] = r

        updated_tabs = []
        checked_info = []
        for day in valid_sheets:
            try:
                ws = spreadsheet.worksheet(day)
            except Exception:
                continue

            all_vals = ws.get_all_values()
            if len(all_vals) < 2:
                continue

            # Auto-detect header row (usually row 2 / index 1, or row 1 / index 0)
            hdr_idx = None
            norm_h = []
            for candidate_r in range(min(5, len(all_vals))):
                test_h = [
                    "".join(str(h).lower().split())
                    for h in all_vals[candidate_r]
                ]
                if "availability" in test_h or "avail" in test_h:
                    hdr_idx = candidate_r
                    norm_h = test_h
                    break

            if hdr_idx is None:
                continue

            def find_idx(targets):
                for t in targets:
                    if t in norm_h:
                        return norm_h.index(t)
                return None

            idx_avail = find_idx(["availability", "avail"])
            idx_name = find_idx(["name", "fullname"])
            idx_first = find_idx(["first_name", "firstname"])
            idx_last = find_idx(["last_name", "lastname"])
            idx_1st = find_idx(["1stchoice", "position"])
            idx_2nd = find_idx(["2ndchoice"])
            idx_score = find_idx(["score", "rating", "skill"])
            idx_reg = find_idx(["reg/spare", "status"])
            idx_notes = find_idx(["notes", "note"])

            if idx_avail is None:
                continue

            cells_to_update = []
            matched_rows = 0
            for r_idx in range(hdr_idx + 1, len(all_vals)):
                row_vals = all_vals[r_idx]
                keys_to_try = []

                if idx_name is not None and idx_name < len(row_vals):
                    nm_val = str(row_vals[idx_name]).strip()
                    if nm_val:
                        keys_to_try.append(clean_name_key(nm_val))

                if idx_first is not None and idx_last is not None:
                    fp = (
                        str(row_vals[idx_first]).strip()
                        if idx_first < len(row_vals)
                        else ""
                    )
                    lp = (
                        str(row_vals[idx_last]).strip()
                        if idx_last < len(row_vals)
                        else ""
                    )
                    combo = f"{fp} {lp}".strip()
                    if combo:
                        keys_to_try.append(clean_name_key(combo))

                matched_key = next(
                    (k for k in keys_to_try if k in summary_map), None
                )
                if not matched_key:
                    continue

                matched_rows += 1
                s_row = summary_map[matched_key]
                sheet_row_num = r_idx + 1

                new_avail = "Y" if bool(s_row[day]) else "N"
                cur_avail = (
                    str(row_vals[idx_avail]).strip().upper()
                    if idx_avail < len(row_vals)
                    else ""
                )
                cur_bool = cur_avail.startswith("Y")
                if bool(s_row[day]) != cur_bool:
                    cells_to_update.append(
                        gspread.Cell(
                            sheet_row_num, idx_avail + 1, new_avail
                        )
                    )

                field_pairs = [
                    (idx_1st, str(s_row["1st Choice"]).strip()),
                    (idx_2nd, str(s_row["2nd Choice"]).strip()),
                    (idx_reg, str(s_row["Reg / Spare"]).strip()),
                    (idx_notes, str(s_row["Notes"]).strip()),
                ]
                for c_idx, new_v in field_pairs:
                    if c_idx is not None:
                        old_v = (
                            str(row_vals[c_idx]).strip()
                            if c_idx < len(row_vals)
                            else ""
                        )
                        if new_v != old_v:
                            cells_to_update.append(
                                gspread.Cell(
                                    sheet_row_num, c_idx + 1, new_v
                                )
                            )

                if idx_score is not None:
                    new_sc = float(s_row["Score"])
                    old_raw = (
                        str(row_vals[idx_score]).strip()
                        if idx_score < len(row_vals)
                        else ""
                    )
                    try:
                        old_sc = float(old_raw) if old_raw else 0.0
                    except ValueError:
                        old_sc = -1.0
                    if abs(new_sc - old_sc) > 0.01:
                        cells_to_update.append(
                            gspread.Cell(
                                sheet_row_num, idx_score + 1, new_sc
                            )
                        )

            checked_info.append(f"{day}: {matched_rows} rows")
            if cells_to_update:
                ws.update_cells(
                    cells_to_update, value_input_option="USER_ENTERED"
                )
                updated_tabs.append(
                    f"{day} ({len(cells_to_update)} cells)"
                )

        if updated_tabs:
            return (
                True,
                f"Saved to '{sheet_title}': {', '.join(updated_tabs)}",
            )
        detail = ", ".join(checked_info) if checked_info else "no matching tabs"
        return (
            True,
            f"No differences found in '{sheet_title}' (checked {detail}).",
        )
    except Exception as e:
        return False, f"Error saving to Google Sheet: {e}"


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


def style_spare_rows(df, disp_cols):
    if df.empty or "Reg / Spare" not in df.columns:
        return df[disp_cols]
    is_spare = (
        df["Reg / Spare"].fillna("").astype(str).str.strip().str.upper()
        == "S"
    )

    def highlight(row):
        if is_spare.loc[row.name]:
            return ["background-color: #EFEFEF; color: #333333"] * len(
                row
            )
        return [""] * len(row)

    return df[disp_cols].style.apply(highlight, axis=1)


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
        "Full Name",
        "Position",
        "Reg / Spare",
        "Score",
        "Email",
        "Phone",
        "Notes",
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
    st.error(f"❌ {error_msg}")
    st.stop()

if all_sheets is not None:
    names = list(all_sheets.keys())
    master_name = next(
        (s for s in names if "master" in s.lower()), None
    )
    blocked = [
        "master", "instructions", "instruction", "reference", "setup"
    ]
    valid_sheets = [
        s
        for s in names
        if not any(b in s.strip().lower() for b in blocked)
    ]

    if not valid_sheets:
        st.error("No valid game sheets found.")
        st.stop()

    if (
        st.session_state.get("sig") != sig
        or "summary_df" not in st.session_state
    ):
        st.session_state.summary_df = build_summary_table(
            all_sheets, valid_sheets, master_name
        )
        st.session_state.sig = sig

    tab_sum, tab_gen = st.tabs(
        ["📋 Weekly Availability Summary", "🏒 Generate Day Lineups"]
    )

    with tab_sum:
        st.subheader("Master Roster & Game-Day Availability")
        st.write(
            "Check or uncheck who is playing on each day below. "
            "Changes update the **Generate Day Lineups** tab."
        )

        m_cols = st.columns(len(valid_sheets))
        for idx, day in enumerate(valid_sheets):
            cnt = int(st.session_state.summary_df[day].sum())
            m_cols[idx].metric(label=f"{day} Players In", value=cnt)

        disp_cols = (
            ["Full Name", "Reg / Spare", "Days"]
            + valid_sheets
            + ["1st Choice", "2nd Choice", "Score", "Notes"]
        )

        cfg = {
            "Full Name": st.column_config.TextColumn(
                "Player Name", disabled=True
            ),
            "Reg / Spare": st.column_config.SelectboxColumn(
                "Reg / Spare", options=["R", "S"], width="small"
            ),
            "Days": st.column_config.SelectboxColumn(
                "Days",
                help="M = Monday, MW = Mon & Wed, W = Wednesday",
                options=["M", "MW", "W", ""],
                width="small",
            ),
            "1st Choice": st.column_config.SelectboxColumn(
                "1st Choice", options=["F", "D", ""], width="small"
            ),
            "2nd Choice": st.column_config.SelectboxColumn(
                "2nd Choice",
                options=["F", "D", "G", ""],
                width="small",
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
            cfg[day] = st.column_config.CheckboxColumn(
                day, default=False, width="small"
            )

        styled_summary = style_spare_rows(
            st.session_state.summary_df, disp_cols
        )
        h_sum = (len(st.session_state.summary_df) + 1) * 35 + 3
        edited = st.data_editor(
            styled_summary,
            column_config=cfg,
            hide_index=True,
            use_container_width=True,
            height=h_sum,
            key="attendance_editor",
        )

        for col in disp_cols:
            st.session_state.summary_df[col] = edited[col]

        btn_c1, btn_c2 = st.columns([2, 2])
        with btn_c1:
            if st.button(
                "💾 Save Changes to Google Sheet",
                type="primary",
                use_container_width=True,
            ):
                with st.spinner("Saving to Google Sheets..."):
                    ok, msg = save_changes_to_google_sheet(
                        sheet_url.strip(),
                        st.session_state.summary_df,
                        valid_sheets,
                    )
                if ok:
                    load_google_sheet_data.clear()
                    st.success(f"✅ {msg}")
                else:
                    st.error(f"❌ {msg}")

        with btn_c2:
            if st.button(
                "↩️ Reset Checkboxes to Sheet Defaults",
                use_container_width=True,
            ):
                st.session_state.summary_df = build_summary_table(
                    all_sheets, valid_sheets, master_name
                )
                st.rerun()

    with tab_gen:
        sel_day = st.selectbox(
            "Select Game Day to Generate Lineups:", valid_sheets
        )
        is_mon = "monday" in sel_day.strip().lower()
        is_wed = "wednesday" in sel_day.strip().lower()

        w_df = st.session_state.summary_df.copy()
        w_df["1st Choice"] = (
            w_df["1st Choice"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )
        w_df["2nd Choice"] = (
            w_df["2nd Choice"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )
        w_df["Reg / Spare"] = (
            w_df["Reg / Spare"]
            .fillna("S")
            .astype(str)
            .str.strip()
            .str.upper()
        )
        w_df["Score"] = pd.to_numeric(
            w_df["Score"], errors="coerce"
        ).fillna(0.0)

        if is_wed:
            g_mask = (
                w_df["Full Name"].apply(clean_name_key)
                == clean_name_key("Gary Fera")
            )
            if g_mask.any():
                w_df.loc[g_mask, "1st Choice"] = "D"
                w_df.loc[g_mask, "2nd Choice"] = "F"

        avail = w_df[w_df[sel_day] == True].copy()
        if avail.empty:
            st.warning(
                f"No players checked in for **{sel_day}**. "
                "Use the **Weekly Availability Summary** tab first."
            )
            st.stop()

        total = len(avail)
        TARGET_D = 8

        miss = avail[~avail["1st Choice"].isin(["F", "D"])]
        if not miss.empty:
            m_names = ", ".join(miss["Full Name"])
            st.warning(
                f"⚠️ Missing 1st Choice for: **{m_names}** "
                "(assigned as Forward; edit in Summary tab)."
            )

        bday_msg, bday_names = get_birthday_message(w_df, "B-day")
        with st.expander("🎂 Birthday Debugger", expanded=True):
            today = datetime.now()
            s_wk = (
                today + timedelta(days=1)
                if today.weekday() == 6
                else today - timedelta(days=today.weekday())
            )
            e_wk = s_wk + timedelta(days=6)
            st.write(
                f"**Scanning Window:** {s_wk.strftime('%a %b %d')}"
                f" — {e_wk.strftime('%a %b %d')}"
            )
            if bday_names:
                st.success(
                    f"**Birthdays Found:** {', '.join(bday_names)}"
                )
            else:
                st.info("No birthdays found this week.")
            st.caption(f"Checked {len(w_df)} players in roster.")

        n_series = avail["Notes"].fillna("").astype(str).str.strip()
        noted = avail[n_series != ""]
        if not noted.empty:
            lbl = f"📝 Active Player Notes ({len(noted)})"
            with st.expander(lbl, expanded=False):
                for _, nr in noted.iterrows():
                    st.write(
                        f"- **{nr['Full Name']}**: {nr['Notes']}"
                    )

        avail = avail.sample(
            frac=1, random_state=st.session_state.shuffle_seed
        ).reset_index(drop=True)
        avail["Status_Rank"] = avail["Reg / Spare"].apply(
            lambda x: 0 if x == "R" else 1
        )
        avail = avail.sort_values(
            by=["Status_Rank", "Score"],
            ascending=[True, False],
            kind="mergesort",
        )

        pool_d = avail[avail["1st Choice"] == "D"].copy()
        pool_f = avail[avail["1st Choice"] == "F"].copy()
        pool_oth = avail[
            ~avail["1st Choice"].isin(["D", "F"])
        ].copy()

        if len(pool_d) < TARGET_D:
            needed = TARGET_D - len(pool_d)
            cands = pool_f[pool_f["2nd Choice"] == "D"].copy()
            if is_mon and not cands.empty:
                cands["Is_Gary"] = (
                    cands["Full Name"].apply(clean_name_key)
                    == clean_name_key("Gary Fera")
                )
                cands = cands.sort_values(
                    by=["Is_Gary", "Status_Rank", "Score"],
                    ascending=[True, True, False],
                    kind="mergesort",
                ).drop(columns=["Is_Gary"])

            if not cands.empty:
                conv = cands.head(needed)
                pool_d = pd.concat([pool_d, conv])
                pool_f = pool_f.drop(conv.index)
                mv = ", ".join(conv["Full Name"])
                st.info(
                    f"Balanced Defense (4 per team): Moved "
                    f"{len(conv)} F -> D: **{mv}**"
                )

        if len(pool_d) > TARGET_D:
            excess = len(pool_d) - TARGET_D
            c_2nd = pool_d[pool_d["2nd Choice"] == "F"].copy()
            if is_wed and not c_2nd.empty:
                c_2nd["Is_Gary"] = (
                    c_2nd["Full Name"].apply(clean_name_key)
                    == clean_name_key("Gary Fera")
                )
                c_2nd = c_2nd.sort_values(
                    by=["Is_Gary", "Status_Rank", "Score"],
                    ascending=[True, True, False],
                    kind="mergesort",
                ).drop(columns=["Is_Gary"])
            if len(c_2nd) >= excess:
                conv = c_2nd.head(excess)
            else:
                rem = excess - len(c_2nd)
                oth = pool_d.drop(c_2nd.index).copy()
                if is_wed and not oth.empty:
                    oth["Is_Gary"] = (
                        oth["Full Name"].apply(clean_name_key)
                        == clean_name_key("Gary Fera")
                    )
                    oth = oth.sort_values(
                        by=["Is_Gary", "Status_Rank", "Score"],
                        ascending=[True, True, False],
                        kind="mergesort",
                    ).drop(columns=["Is_Gary"])
                conv = pd.concat([c_2nd, oth.tail(rem)])
            pool_f = pd.concat([pool_f, conv])
            pool_d = pool_d.drop(conv.index)
            mv = ", ".join(conv["Full Name"])
            st.info(
                f"Balanced Defense (capped at 4 per team): Moved "
                f"{len(conv)} D -> F: **{mv}**"
            )

        elif len(pool_d) % 2 == 1 and len(pool_d) > 4:
            c_2nd = pool_d[pool_d["2nd Choice"] == "F"].copy()
            if is_wed and not c_2nd.empty:
                c_2nd["Is_Gary"] = (
                    c_2nd["Full Name"].apply(clean_name_key)
                    == clean_name_key("Gary Fera")
                )
                c_2nd = c_2nd.sort_values(
                    by=["Is_Gary", "Status_Rank", "Score"],
                    ascending=[True, True, False],
                    kind="mergesort",
                ).drop(columns=["Is_Gary"])
            if not c_2nd.empty:
                conv = c_2nd.head(1)
            else:
                oth = pool_d.copy()
                if is_wed:
                    oth["Is_Gary"] = (
                        oth["Full Name"].apply(clean_name_key)
                        == clean_name_key("Gary Fera")
                    )
                    oth = oth.sort_values(
                        by=["Is_Gary", "Status_Rank", "Score"],
                        ascending=[True, True, False],
                        kind="mergesort",
                    ).drop(columns=["Is_Gary"])
                conv = oth.tail(1)
            pool_f = pd.concat([pool_f, conv])
            pool_d = pool_d.drop(conv.index)
            mv = ", ".join(conv["Full Name"])
            d_each = len(pool_d) // 2
            st.info(
                f"Evened Defense ({d_each} per team): Moved "
                f"1 D -> F: **{mv}**"
            )

        if not pool_oth.empty:
            pool_f = pd.concat([pool_f, pool_oth])

        d_each = len(pool_d) // 2
        st.info(
            f"**Strategy ({sel_day}):** {total} players -> "
            f"**{len(pool_d)} Defense** ({d_each} per team) & "
            f"**{len(pool_f)} Forwards**"
        )

        ta, tb, cuts_d, cuts_f, rival_logs = build_rosters(
            pool_d, pool_f, is_wed, allow_wed_pair=False
        )

        if is_wed:
            c_cnt = min(len(ta), len(tb))
            gap = abs(
                get_top_n_score(ta, c_cnt)
                - get_top_n_score(tb, c_cnt)
            )
            if gap > 1.5:
                alt = build_rosters(
                    pool_d, pool_f, is_wed, allow_wed_pair=True
                )
                ta_a, tb_a, cd_a, cf_a, rl_a = alt
                a_gap = abs(
                    get_top_n_score(ta_a, c_cnt)
                    - get_top_n_score(tb_a, c_cnt)
                )
                if a_gap < gap:
                    ta, tb = ta_a, tb_a
                    cuts_d, cuts_f = cd_a, cf_a
                    rival_logs = rl_a

        if st.button("🔀 Shuffle Teams Again"):
            st.session_state.shuffle_seed += 1
            st.rerun()

        if rival_logs:
            st.divider()
            for r in rival_logs:
                st.success(f"⚖️ {r}")

        c_cnt = min(len(ta), len(tb))
        tot_a = ta["Score"].sum()
        tot_b = tb["Score"].sum()
        fair_a = get_top_n_score(ta, c_cnt)
        fair_b = get_top_n_score(tb, c_cnt)

        c1, c2 = st.columns(2)
        with c1:
            st.header("🔴 Red Team")
            st.write(f"**Total Score:** {tot_a:.1f}")
            if len(ta) != len(tb):
                st.write(f"**Top {c_cnt} Score:** {fair_a:.1f}")
            da_n = len(ta[ta["Position"] == "D"])
            fa_n = len(ta[ta["Position"] == "F"])
            st.write(f"Players: {len(ta)} ({da_n} D / {fa_n} F)")
            h_a = (len(ta) + 1) * 35 + 3
            st.dataframe(
                style_spare_rows(ta, ["Full Name", "Position"]),
                hide_index=True,
                use_container_width=True,
                height=h_a,
            )
        with c2:
            st.header("⚪ White Team")
            st.write(f"**Total Score:** {tot_b:.1f}")
            if len(ta) != len(tb):
                st.write(f"**Top {c_cnt} Score:** {fair_b:.1f}")
            db_n = len(tb[tb["Position"] == "D"])
            fb_n = len(tb[tb["Position"] == "F"])
            st.write(f"Players: {len(tb)} ({db_n} D / {fb_n} F)")
            h_b = (len(tb) + 1) * 35 + 3
            st.dataframe(
                style_spare_rows(tb, ["Full Name", "Position"]),
                hide_index=True,
                use_container_width=True,
                height=h_b,
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
            recips = [
                e
                for e in all_p["Email"].unique()
                if pd.notna(e) and str(e).strip()
            ]
            bcc = ",".join(recips)
            r_txt = format_team_list(ta, "RED TEAM")
            w_txt = format_team_list(tb, "WHITE TEAM")
            body = (
                f"{bday_msg}Hello everyone,\n\n"
                f"Here are the rosters for the upcoming game:\n\n"
                f"{r_txt}\n"
                f"{w_txt}\n"
                f"Keep your sticks on the ice!"
            )

            st.text_area("Email Draft", value=body, height=300)
            subj = urllib.parse.quote(
                f"Bendo Hockey Lineups - {sel_day}"
            )
            s_body = urllib.parse.quote(body)
            s_bcc = urllib.parse.quote(bcc)

            b1, b2, b3 = st.columns(3)
            b1.link_button(
                "📱 Default App",
                f"mailto:?bcc={s_bcc}&subject={subj}&body={s_body}",
            )
            b2.link_button(
                "💻 Gmail Web",
                f"https://mail.google.com/mail/?view=cm&fs=1"
                f"&bcc={s_bcc}&su={subj}&body={s_body}",
            )
            b3.link_button(
                "🍎 iOS Gmail",
                f"googlegmail:///co?bcc={s_bcc}"
                f"&subject={subj}&body={s_body}",
            )
