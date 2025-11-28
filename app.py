import streamlit as st
import pandas as pd
import io
import tempfile
import os
import re
from datetime import datetime

# -----------------------
# Configurations / defaults
# -----------------------
DAY_ROW = 4         # row index (0-based) where days live in Doodle export
TIME_ROW = 5        # row index (0-based) where times live
FIRST_COL = 2       # first column index where day/time cells start
MIN_ROWS = 6        # minimal rows expected in a Doodle sheet

WRONG_FORMAT_MSG = "this specific document is not on the right format, please input the correct version."
REQUIRED_MEM_COLS = {"Member Name", "Position", "Semesters at NJC"}

# -----------------------
# Utilities
# -----------------------
time_re = re.compile(r'(\d{1,2}:\d{2}\s*[APap][Mm])')

WEEKDAY_MAP = {
    "Mon": "Monday", "Tue": "Tuesday", "Wed": "Wednesday",
    "Thu": "Thursday", "Fri": "Friday", "Sat": "Saturday",
    "Sun": "Sunday",
    "Monday": "Monday", "Tuesday": "Tuesday", "Wednesday": "Wednesday",
    "Thursday": "Thursday", "Friday": "Friday", "Saturday": "Saturday",
    "Sunday": "Sunday"
}


def parse_slot_to_day_time(slot_text):
    """
    Returns (day_name, 'HH:MM') given a slot like "Mon 9:00 AM" or "Monday 09:00".
    If it can't parse day or time, returns (None, None).
    """
    if not isinstance(slot_text, str) or not slot_text.strip():
        return None, None

    s = " ".join(slot_text.split())
    tokens = s.split()
    day_token = tokens[0].rstrip(",")
    day = WEEKDAY_MAP.get(day_token, None)

    # Try AM/PM
    m = time_re.search(s)
    if m:
        time_str = m.group(1).upper().replace(" ", "")
        time_str = re.sub(r'([AP]M)$', r' \1', time_str)
        try:
            dt = datetime.strptime(time_str, "%I:%M %p")
            return day, dt.strftime("%H:%M")
        except Exception:
            return day, None

    # fallback: 24h hh:mm
    m2 = re.search(r'(\d{1,2}:\d{2})', s)
    if m2:
        parts = m2.group(1).split(":")
        return day, parts[0].zfill(2) + ":" + parts[1][:2]

    return day, None


def looks_like_member_sheet(df):
    """
    Heuristic: if any of the required member columns appear (case-insensitive)
    in the first few rows/columns, it's probably the member-info sheet.
    """
    try:
        text = df.head(8).astype(str).apply(lambda col: col.str.lower())
        for col in text.columns:
            colvals = text[col].tolist()
            for required in REQUIRED_MEM_COLS:
                if any(required.lower() in (str(v)) for v in colvals):
                    return True
        flat = " ".join(" ".join(row) for row in text.values.tolist())
        if any(req.lower() in flat for req in REQUIRED_MEM_COLS):
            return True
    except Exception:
        return False
    return False

# -----------------------
# Validation Helpers
# -----------------------


def _row_slice(df, row_idx):
    """Safely slice a row from FIRST_COL onwards; return None if it is out of bounds."""
    if row_idx >= df.shape[0] or df.shape[1] <= FIRST_COL:
        return None
    return df.iloc[row_idx, FIRST_COL:]


def _looks_like_weekday_row(ser):
    """True if any cell in the series starts with a weekday."""
    if ser is None:
        return False
    vals = ser.astype(str).str.strip()
    for v in vals:
        if not v or v.lower() == "nan":
            continue
        first = v.split()[0].rstrip(",")
        if first in WEEKDAY_MAP:
            return True
    return False


def _looks_like_time_row(ser):
    """True if any cell in the series contains a hh:mm (with or without AM/PM)."""
    if ser is None:
        return False
    vals = ser.astype(str).str.strip()
    time_pattern = re.compile(r"\b\d{1,2}:\d{2}\b")
    for v in vals:
        if not v or v.lower() == "nan":
            continue
        if time_re.search(v) or time_pattern.search(v):
            return True
    return False


def _has_availability_answers(df):
    """
    True if any non-empty cell exists in the availability region
    (rows 6+ and columns from FIRST_COL onward).
    """
    if df.shape[0] <= 6 or df.shape[1] <= FIRST_COL:
        return False
    try:
        sub = df.iloc[6:, FIRST_COL:]
        vals = sub.astype(str).values.flatten()
        for v in vals:
            if v is None:
                continue
            v_str = str(v).strip()
            if v_str and v_str.lower() not in ("nan", "none"):
                return True
    except Exception:
        return False
    return False


def validate_doodle_structure(df):
    """Analyse a Doodle export and report which key components are present."""
    expected = {
        "structure": "enough rows and columns (at least 6 rows and 3 columns)",
        "day": "a day row with weekday labels (e.g., Mon, Tue, Wed)",
        "time": "a time row with times (e.g., 9:00 AM or 09:00)",
        "answers": "availability cells (non-empty values in participant rows)",
    }

    has_structure = df.shape[0] >= MIN_ROWS and df.shape[1] > FIRST_COL
    day_ser = _row_slice(df, DAY_ROW)
    time_ser = _row_slice(df, TIME_ROW)
    has_day = _looks_like_weekday_row(day_ser)
    has_time = _looks_like_time_row(time_ser)
    has_answers = _has_availability_answers(df)

    present = []
    missing = []

    for key, desc in expected.items():
        cond = {
            "structure": has_structure,
            "day": has_day,
            "time": has_time,
            "answers": has_answers,
        }[key]
        (present if cond else missing).append(desc)

    ok = len(missing) == 0
    return ok, present, missing


def validate_member_info(mem_df):
    """Check if member-info sheet has the required columns."""
    cols = set(str(c) for c in mem_df.columns)
    required = set(REQUIRED_MEM_COLS)
    present = sorted([c for c in required if c in cols])
    missing = sorted([c for c in required if c not in cols])
    ok = len(missing) == 0
    return ok, present, missing

# -------------------------
# Doodle preprocess + parse
# -------------------------


def fill_days_in_doodle(
    doodle_path,
    doodle_cleaned_path,
    day_row=DAY_ROW,
    first_column=FIRST_COL,
):
    """
    Read the Doodle Excel (header=None), forward-fill the day row and write a
    cleaned workbook.

    Assumes the file has already passed `validate_doodle_structure`.
    """
    table = pd.read_excel(doodle_path, header=None)

    if table.shape[0] <= day_row or table.shape[1] <= first_column:
        raise ValueError("Doodle file does not have enough rows/columns for day headers.")

    day_series = table.loc[day_row, first_column:]

    day_series = (
        day_series.fillna("")
        .astype(str)
        .replace(["nan", "NaN", "None", "none"], "")
    )

    cleaned = day_series.replace(["", " "], pd.NA)

    if cleaned.isna().all():
        raise ValueError("Doodle day row appears to be empty after reading.")

    filled = cleaned.ffill()

    table.loc[day_row, first_column:] = filled
    table.to_excel(doodle_cleaned_path, index=False, header=False)


def parse_doodle(table, skip_names=None, day_row=DAY_ROW, first_column=FIRST_COL, time_row=TIME_ROW):
    """
    Parse the cleaned Doodle table (header=None) into availability dict and slots list.
    Returns (availability_dict, slots_list)
    """
    if skip_names is None:
        skip_names = set()

    if table.shape[0] <= time_row:
        raise ValueError("Doodle file missing expected rows for day/time headers.")

    try:
        days = table.loc[day_row, first_column:].fillna("").astype(str).str.strip()
        times = table.loc[time_row, first_column:].fillna("").astype(str).str.strip()
    except Exception as e:
        raise ValueError("Could not read day/time rows.") from e

    if days.eq("").all() or times.eq("").all():
        raise ValueError("Doodle day or time row is empty or malformed.")

    try:
        slots = (days + " " + times).tolist()
    except Exception:
        slots = []
        maxlen = max(len(days), len(times))
        for i in range(maxlen):
            d = days.iloc[i] if i < len(days) else ""
            t = times.iloc[i] if i < len(times) else ""
            slots.append(f"{d} {t}".strip())

    availability = {}
    skip_upper = {s.upper() for s in skip_names}

    for i in range(6, len(table)):
        name_cell = table.iloc[i, 0] if table.shape[1] > 0 else None
        if not isinstance(name_cell, str) or not name_cell.strip():
            continue
        name_clean = name_cell.strip()
        if name_clean.upper() in skip_upper:
            continue

        email = table.iloc[i, 1] if table.shape[1] > 1 else None
        yes_slots, ifnb_slots = [], []

        for j, slot in enumerate(slots):
            col_idx = first_column + j
            if col_idx >= table.shape[1]:
                break
            answer = table.iloc[i, col_idx]
            if pd.isna(answer):
                continue
            if isinstance(answer, str):
                a = answer.strip().upper()
            else:
                a = str(answer).strip().upper()

            if a == "YES":
                yes_slots.append(slot)
            elif a in ("IF NEED BE", "IF NEEDED", "IF NEED", "IFNEEDBE"):
                ifnb_slots.append(slot)

        availability[name_clean] = {"email": email, "yes": yes_slots, "ifnb": ifnb_slots}

    return availability, slots

# ----------------------------------------------------------------
# Classification & scheduling (with back-to-back interview option)
# ----------------------------------------------------------------


def classify_interviewers(interviewer_availability, member_info):
    seniors, juniors = [], []
    interviewer_slots, interviewer_yes, interviewer_ifnb = {}, {}, {}

    for name, data in interviewer_availability.items():
        row = member_info[member_info["Member Name"] == name]
        if row.empty:
            lower_mask = member_info["Member Name"].astype(str).str.lower() == name.lower()
            row = member_info[lower_mask]

        if row.empty:
            juniors.append(name)
        else:
            position = str(row.iloc[0].get("Position", "")).lower()
            try:
                semester = int(row.iloc[0].get("Semesters at NJC", 0))
            except Exception:
                semester = 0
            if "board" in position or "principal" in position or semester > 2:
                seniors.append(name)
            else:
                juniors.append(name)

        interviewer_slots[name] = set(data.get("yes", []) + data.get("ifnb", []))
        interviewer_yes[name] = set(data.get("yes", []))
        interviewer_ifnb[name] = set(data.get("ifnb", []))

    return seniors, juniors, interviewer_slots, interviewer_yes, interviewer_ifnb


def compute_slot_strength(seniors, juniors, interviewer_slots):
    if not interviewer_slots:
        return {}, set()
    all_slots = set().union(*interviewer_slots.values())
    strength = {}
    for slot in all_slots:
        s_count = sum(slot in interviewer_slots.get(s, set()) for s in seniors)
        j_count = sum(slot in interviewer_slots.get(j, set()) for j in juniors)
        sj_teams = min(s_count, j_count)
        ss_teams = s_count // 2
        strength[slot] = sj_teams + ss_teams
    return strength, all_slots


def schedule_interviews(candidate_availability, seniors, juniors,
                        interviewer_slots, interviewer_yes, interviewer_ifnb,
                        all_slots, slot_strength, allow_back_to_back=True):
    """
    Schedule interviews.

    If allow_back_to_back = False, the same interviewer will not be scheduled in
    two slots that are within 60 minutes of each other on the same day.

    Any candidate who cannot be allocated will appear with a
    "PROBLEM, CANDIDATE WAS NOT ALLOCATED DUE TO ..." message.
    """

    def sort_slots(slots):
        return sorted(
            [s for s in slots if isinstance(s, str) and s.strip()],
            key=lambda s: slot_strength.get(s, 0),
            reverse=True
        )

    def slot_to_day_minutes(slot):
        day, t = parse_slot_to_day_time(slot)
        if day is None or t is None:
            return None, None
        try:
            h, m = t.split(":")
            return day, int(h) * 60 + int(m)
        except Exception:
            return day, None

    interviewer_calendar = {name: [] for name in seniors + juniors}
    booked_s = {slot: set() for slot in all_slots}
    booked_j = {slot: set() for slot in all_slots}
    load = {name: 0 for name in seniors + juniors}
    results = []
    scheduled_candidates = set()

    def can_take_slot(name, slot):
        if allow_back_to_back:
            return True
        day, minutes = slot_to_day_minutes(slot)
        if day is None or minutes is None:
            return True
        for d_existing, m_existing in interviewer_calendar.get(name, []):
            if d_existing == day and abs(m_existing - minutes) <= 60:
                return False
        return True

    def register_slot(name, slot):
        day, minutes = slot_to_day_minutes(slot)
        if day is None or minutes is None:
            return
        interviewer_calendar.setdefault(name, []).append((day, minutes))

    # Main scheduling loop
    for candidate, cdata in candidate_availability.items():
        cand_slots_sorted = sort_slots(cdata.get("yes", []) + cdata.get("ifnb", []))
        allocated = False

        for slot in cand_slots_sorted:
            s_yes = [
                s for s in seniors
                if slot in interviewer_yes.get(s, set())
                and s not in booked_s.get(slot, set())
                and can_take_slot(s, slot)
            ]
            j_yes = [
                j for j in juniors
                if slot in interviewer_yes.get(j, set())
                and j not in booked_j.get(slot, set())
                and can_take_slot(j, slot)
            ]
            s_ifnb = [
                s for s in seniors
                if slot in interviewer_ifnb.get(s, set())
                and s not in booked_s.get(slot, set())
                and can_take_slot(s, slot)
            ]
            j_ifnb = [
                j for j in juniors
                if slot in interviewer_ifnb.get(j, set())
                and j not in booked_j.get(slot, set())
                and can_take_slot(j, slot)
            ]

            for lst in [s_yes, j_yes, s_ifnb, j_ifnb]:
                lst.sort(key=lambda x: load.get(x, 0))

            # Senior + Junior YES
            if s_yes and j_yes:
                S, J = s_yes[0], j_yes[0]
                booked_s.setdefault(slot, set()).add(S)
                booked_j.setdefault(slot, set()).add(J)
                load[S] += 1
                load[J] += 1
                register_slot(S, slot)
                register_slot(J, slot)
                results.append({
                    "Candidate": candidate,
                    "Email": cdata.get("email"),
                    "Slot": slot,
                    "Team Type": "Senior + Junior",
                    "Senior1": S, "Senior2": None,
                    "Junior1": J, "Junior2": None
                })
                allocated = True
                break

            # Senior + Senior YES
            if len(s_yes) >= 2:
                S1, S2 = s_yes[:2]
                booked_s.setdefault(slot, set()).update({S1, S2})
                load[S1] += 1
                load[S2] += 1
                register_slot(S1, slot)
                register_slot(S2, slot)
                results.append({
                    "Candidate": candidate,
                    "Email": cdata.get("email"),
                    "Slot": slot,
                    "Team Type": "Senior + Senior",
                    "Senior1": S1, "Senior2": S2,
                    "Junior1": None, "Junior2": None
                })
                allocated = True
                break

            # Senior + Junior fallback
            s_mix = s_yes + s_ifnb
            j_mix = j_yes + j_ifnb
            if s_mix and j_mix:
                S = sorted(s_mix, key=lambda x: load.get(x, 0))[0]
                J = sorted(j_mix, key=lambda x: load.get(x, 0))[0]
                booked_s.setdefault(slot, set()).add(S)
                booked_j.setdefault(slot, set()).add(J)
                load[S] += 1
                load[J] += 1
                register_slot(S, slot)
                register_slot(J, slot)
                results.append({
                    "Candidate": candidate,
                    "Email": cdata.get("email"),
                    "Slot": slot,
                    "Team Type": "Senior + Junior (Acceptable Fallback)",
                    "Senior1": S, "Senior2": None,
                    "Junior1": J, "Junior2": None
                })
                allocated = True
                break

            # Senior + Senior fallback
            if len(s_mix) >= 2:
                S1, S2 = sorted(s_mix, key=lambda x: load.get(x, 0))[:2]
                booked_s.setdefault(slot, set()).update({S1, S2})
                load[S1] += 1
                load[S2] += 1
                register_slot(S1, slot)
                register_slot(S2, slot)
                results.append({
                    "Candidate": candidate,
                    "Email": cdata.get("email"),
                    "Slot": slot,
                    "Team Type": "Senior + Senior (Acceptable Fallback)",
                    "Senior1": S1, "Senior2": S2,
                    "Junior1": None, "Junior2": None
                })
                allocated = True
                break

            # Junior + Junior fallback
            jj_mix = j_yes + j_ifnb
            if len(jj_mix) >= 2:
                J1, J2 = sorted(jj_mix, key=lambda x: load.get(x, 0))[:2]
                booked_j.setdefault(slot, set()).update({J1, J2})
                load[J1] += 1
                load[J2] += 1
                register_slot(J1, slot)
                register_slot(J2, slot)
                results.append({
                    "Candidate": candidate,
                    "Email": cdata.get("email"),
                    "Slot": slot,
                    "Team Type": "Junior + Junior (Problematic Fallback)",
                    "Senior1": None, "Senior2": None,
                    "Junior1": J1, "Junior2": J2
                })
                allocated = True
                break

        if allocated:
            scheduled_candidates.add(candidate)

    # Add PROBLEM rows for candidates that were not allocated
    for candidate, cdata in candidate_availability.items():
        if candidate in scheduled_candidates:
            continue

        cand_slots = set(cdata.get("yes", []) + cdata.get("ifnb", []))
        email = cdata.get("email")

        if not cand_slots:
            reason_detail = "NO RECORDED AVAILABILITY IN THE DOODLE FILE."
        else:
            all_interviewer_slots = set().union(*interviewer_slots.values()) if interviewer_slots else set()
            overlap = cand_slots & all_interviewer_slots
            if not overlap:
                reason_detail = "NO OVERLAPPING AVAILABILITY WITH ANY INTERVIEWER."
            else:
                reason_detail = (
                    "SCHEDULING CONSTRAINTS (load balancing, back-to-back limits, "
                    "or team composition rules)."
                )

        msg = f"PROBLEM, CANDIDATE WAS NOT ALLOCATED DUE TO {reason_detail}"

        results.append({
            "Candidate": candidate,
            "Email": email,
            "Slot": None,
            "Team Type": msg,
            "Senior1": None, "Senior2": None,
            "Junior1": None, "Junior2": None
        })

    return pd.DataFrame(results)

# -----------------------------------
# Calendar builders (display + excel)
# -----------------------------------
PALETTE = ["#e8f4ff", "#e4ffe8", "#fff4e5", "#f9e6ff", "#ffecec", "#f0f7ff", "#fff0f5"]
split_pattern = re.compile(r"\s*&\s*|\s*/\s*|\s*,\s*")


def build_weekly_calendar(assignments, weekdays=None):
    if weekdays is None:
        weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

    def normalize_time(t):
        if not isinstance(t, str) or not t:
            return None
        s = t.strip()
        m = re.match(r'^\s*(\d{1,2}):(\d{2})\s*([AaPp][Mm])\s*$', s)
        if m:
            try:
                dt = datetime.strptime(f"{m.group(1)}:{m.group(2)} {m.group(3).upper()}", "%I:%M %p")
                return dt.strftime("%H:%M")
            except Exception:
                return f"{int(m.group(1)):02d}:{m.group(2)}"
        m2 = re.match(r'^\s*(\d{1,2}):(\d{2})\s*$', s)
        if m2:
            return f"{int(m2.group(1)):02d}:{m2.group(2)}"
        return None

    raw_times = {normalize_time(a.get("time")) for a in assignments if a.get("time")}
    mins = sorted({int(t[:2]) * 60 + int(t[3:]) for t in raw_times if t})

    if not mins:
        times = [f"{h:02d}:00" for h in range(8, 20)]
    else:
        cleaned = []
        last = None
        for m in mins:
            if last is None:
                cleaned.append(m)
            elif m - last <= 45:
                cleaned.append(m)
            else:
                cleaned.append(None)
                cleaned.append(m)
            last = m
        times = [(f"{m // 60:02d}:{m % 60:02d}") if m is not None else "" for m in cleaned]

    table = pd.DataFrame("", index=times, columns=weekdays)
    buckets = {d: {t: [] for t in times} for d in weekdays}

    for a in assignments:
        d = a.get("day")
        t = normalize_time(a.get("time"))
        if d not in weekdays or not t:
            continue
        raw_panel = a.get("interviewer") or ""
        panel = [p.strip() for p in split_pattern.split(raw_panel) if p.strip()]
        if not panel:
            panel = ["TBD"]
        buckets[d][t].append({"candidate": a.get("candidate", "TBD"), "panel": panel})

    def format_cell_entries(entries, slot_index=0):
        if not entries:
            return ""
        html_blocks = []
        for i, e in enumerate(entries):
            bg = PALETTE[(slot_index + i) % len(PALETTE)]
            candidate_clean = str(e["candidate"]).replace("\n", " ").strip()
            panel_clean = [p.replace("\n", " ").strip() for p in e["panel"]]
            cand_html = f"<div style='font-weight:700;color:#183a6e;margin-bottom:4px'>{candidate_clean}</div>"
            panel_html = "<div style='color:#333;font-size:13px;margin-bottom:2px;'>" + " / ".join(panel_clean) + "</div>"
            block = (
                f"<div style='background:{bg};padding:10px;margin-bottom:8px;"
                f"border-radius:10px;border:1px solid rgba(0,0,0,0.06);"
                f"box-shadow:0 1px 2px rgba(0,0,0,0.06);'>"
                f"{cand_html}{panel_html}</div>"
            )
            html_blocks.append(block)
        return "".join(html_blocks)

    for idx, t in enumerate(times):
        for d in weekdays:
            entries = buckets[d][t]
            table.at[t, d] = format_cell_entries(entries, slot_index=idx)
    return table


def build_excel_calendar(assignments):
    """
    Build simple table Day | Time | Candidate | Interviewers
    """
    df = pd.DataFrame(assignments)
    if "day" in df.columns and "Day" not in df.columns:
        df = df.rename(columns={"day": "Day"})
    if "time" in df.columns and "Time" not in df.columns:
        df = df.rename(columns={"time": "Time"})

    for col in ["Day", "Time", "candidate", "interviewer"]:
        if col not in df.columns:
            df[col] = pd.NA

    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    df["DayOrder"] = df["Day"].apply(lambda x: day_order.index(x) if str(x) in day_order else 999)

    def time_to_minutes(t):
        if pd.isna(t):
            return 0
        try:
            parts = str(t).split(":")
            return int(parts[0]) * 60 + int(parts[1])
        except Exception:
            return 0

    df["TimeMinutes"] = df["Time"].apply(time_to_minutes)
    df = df.sort_values(["DayOrder", "TimeMinutes"]).reset_index(drop=True)
    df = df.drop(columns=["DayOrder", "TimeMinutes"])
    df = df[["Day", "Time", "candidate", "interviewer"]]
    return df

# -----------------------
# Streamlit UI / Flow
# -----------------------

st.set_page_config(page_title="NJC Interview Scheduler", layout="wide")

st.title("🎓 NJC Interview Scheduler")
st.markdown(
    "Welcome! This app helps you turn **Doodle availability exports** "
    "into a balanced interview schedule, a weekly calendar, and interviewer summaries.\n\n"
    "Any candidate we *cannot* place will be clearly flagged in the final table."
)

with st.expander("ℹ️ What files do I need? (quick overview)", expanded=True):
    st.markdown(
        """
        **You will need three Excel files:**

        1. **Candidates Doodle export (`.xlsx`)**  
           - Exported **directly from Doodle** (no edits to the structure).  
           - Each **row = one candidate**.  
           - Columns from the third column onwards contain **YES / IF NEED BE** answers.  
           - The organizer row like **“NJC”** is OK – the app will ignore it.

        2. **Interviewers Doodle export (`.xlsx`)**  
           - Also exported directly from Doodle.  
           - Each **row = one interviewer**.  
           - Same structure: days + times at the top, answers below.  
           - The organizer row **“NJC”** is ignored here as well.

        3. **Member Info sheet (`.xlsx`)**  
           - A regular Excel table with at least these columns **as headers**:  
             - `Member Name`  
             - `Position`  
             - `Semesters at NJC`  
           - Used to classify **seniors vs juniors**.
        """
    )

st.markdown("---")

col1, col2, col3 = st.columns(3)
with col1:
    cand_file = st.file_uploader("📄 Upload *Candidates* Doodle", type=["xlsx"])
with col2:
    int_file = st.file_uploader("👥 Upload *Interviewers* Doodle", type=["xlsx"])
with col3:
    mem_file = st.file_uploader("📘 Upload Member Info Sheet", type=["xlsx"])

back_to_back_choice = st.radio(
    "Can interviewers do back-to-back interviews (e.g., one interview at 13:00 then another at 14:00)?",
    ["Yes", "No"],
    index=0
)
allow_back_to_back = (back_to_back_choice == "Yes")

if st.button("🚀 Run Scheduling"):
    if not (cand_file and int_file and mem_file):
        st.error("Please upload all three files before running the scheduler.")
        st.stop()

    with tempfile.TemporaryDirectory() as tmp:
        cand_path = os.path.join(tmp, "cand.xlsx")
        int_path = os.path.join(tmp, "int.xlsx")
        mem_path = os.path.join(tmp, "mem.xlsx")

        open(cand_path, "wb").write(cand_file.read())
        open(int_path, "wb").write(int_file.read())
        open(mem_path, "wb").write(mem_file.read())

        cand_clean = os.path.join(tmp, "cand_clean.xlsx")
        int_clean = os.path.join(tmp, "int_clean.xlsx")

        # Structural validation – candidates
        try:
            raw_cand = pd.read_excel(cand_path, header=None)
        except Exception:
            st.error("Candidates Doodle: could not read the file. Please upload a valid Excel export from Doodle.")
            st.stop()

        ok_cand, present_cand, missing_cand = validate_doodle_structure(raw_cand)
        if not ok_cand:
            if not present_cand:
                st.error(
                    "Candidates Doodle: your file does not contain the required Doodle structure.\n\n"
                    "It **should contain**:\n"
                    "- enough rows and columns (at least 6 rows and 3 columns)\n"
                    "- a day row with weekday labels (e.g., Mon, Tue, Wed)\n"
                    "- a time row with times (e.g., 9:00 AM or 09:00)\n"
                    "- availability cells (non-empty values in participant rows)"
                )
            else:
                st.error(
                    "Candidates Doodle: this document is not in the expected Doodle format.\n\n"
                    "It **should contain** all of the following:\n"
                    "- enough rows and columns (at least 6 rows and 3 columns)\n"
                    "- a day row with weekday labels (e.g., Mon, Tue, Wed)\n"
                    "- a time row with times (e.g., 9:00 AM or 09:00)\n"
                    "- availability cells (non-empty values in participant rows)\n\n"
                    "Your file **seems to contain**:\n"
                    + "".join(f"- {item}\n" for item in present_cand) +
                    "\nBut it is **missing**:\n"
                    + "".join(f"- {item}\n" for item in missing_cand)
                )
            st.stop()

        # Structural validation – interviewers
        try:
            raw_int = pd.read_excel(int_path, header=None)
        except Exception:
            st.error("Interviewers Doodle: could not read the file. Please upload a valid Excel export from Doodle.")
            st.stop()

        ok_int, present_int, missing_int = validate_doodle_structure(raw_int)
        if not ok_int:
            if not present_int:
                st.error(
                    "Interviewers Doodle: your file does not contain the required Doodle structure.\n\n"
                    "It **should contain**:\n"
                    "- enough rows and columns (at least 6 rows and 3 columns)\n"
                    "- a day row with weekday labels (e.g., Mon, Tue, Wed)\n"
                    "- a time row with times (e.g., 9:00 AM or 09:00)\n"
                    "- availability cells (non-empty values in participant rows)"
                )
            else:
                st.error(
                    "Interviewers Doodle: this document is not in the expected Doodle format.\n\n"
                    "It **should contain** all of the following:\n"
                    "- enough rows and columns (at least 6 rows and 3 columns)\n"
                    "- a day row with weekday labels (e.g., Mon, Tue, Wed)\n"
                    "- a time row with times (e.g., 9:00 AM or 09:00)\n"
                    "- availability cells (non-empty values in participant rows)\n\n"
                    "Your file **seems to contain**:\n"
                    + "".join(f"- {item}\n" for item in present_int) +
                    "\nBut it is **missing**:\n"
                    + "".join(f"- {item}\n" for item in missing_int)
                )
            st.stop()

        # Preprocess doodles (after structural check)
        try:
            fill_days_in_doodle(cand_path, cand_clean)
        except ValueError as e:
            st.error(f"Candidates Doodle: {WRONG_FORMAT_MSG} ({str(e)})")
            st.stop()
        except Exception:
            st.error(f"Candidates Doodle: {WRONG_FORMAT_MSG}")
            st.stop()

        try:
            fill_days_in_doodle(int_path, int_clean)
        except ValueError as e:
            st.error(f"Interviewers Doodle: {WRONG_FORMAT_MSG} ({str(e)})")
            st.stop()
        except Exception:
            st.error(f"Interviewers Doodle: {WRONG_FORMAT_MSG}")
            st.stop()

        # Read cleaned files
        try:
            cand_table = pd.read_excel(cand_clean, header=None)
        except Exception:
            st.error(f"Candidates Doodle: {WRONG_FORMAT_MSG}")
            st.stop()

        try:
            int_table = pd.read_excel(int_clean, header=None)
        except Exception:
            st.error(f"Interviewers Doodle: {WRONG_FORMAT_MSG}")
            st.stop()

        # Member info
        try:
            mem_df = pd.read_excel(mem_path)
        except Exception:
            st.error(f"Member Info Sheet: {WRONG_FORMAT_MSG}")
            st.stop()

        ok_mem, present_mem, missing_mem = validate_member_info(mem_df)
        if not ok_mem:
            if not present_mem:
                st.error(
                    "Member Info Sheet: your file does not contain the required columns.\n\n"
                    "It **should contain** at least the following columns:\n"
                    "- Member Name\n- Position\n- Semesters at NJC"
                )
            else:
                st.error(
                    "Member Info Sheet: this document is missing some required columns.\n\n"
                    "It **should contain** the following columns:\n"
                    "- Member Name\n- Position\n- Semesters at NJC\n\n"
                    "Your file **currently has**:\n"
                    + "".join(f"- {c}\n" for c in present_mem) +
                    "\nBut it is **missing**:\n"
                    + "".join(f"- {c}\n" for c in missing_mem)
                )
            st.stop()

        # Detect obvious wrong-file uploads
        if looks_like_member_sheet(cand_table):
            st.error(f"Candidates Doodle: Looks like a Member Info sheet was uploaded. {WRONG_FORMAT_MSG}")
            st.stop()
        if looks_like_member_sheet(int_table):
            st.error(f"Interviewers Doodle: Looks like a Member Info sheet was uploaded. {WRONG_FORMAT_MSG}")
            st.stop()

        # Validate that day/time rows exist and have content
        try:
            cand_days = cand_table.loc[DAY_ROW, FIRST_COL:].fillna("").astype(str).str.strip()
            cand_times = cand_table.loc[TIME_ROW, FIRST_COL:].fillna("").astype(str).str.strip()
        except Exception:
            st.error(f"Candidates Doodle: {WRONG_FORMAT_MSG}")
            st.stop()
        if cand_days.eq("").all() or cand_times.eq("").all():
            st.error(f"Candidates Doodle: {WRONG_FORMAT_MSG}")
            st.stop()

        try:
            int_days = int_table.loc[DAY_ROW, FIRST_COL:].fillna("").astype(str).str.strip()
            int_times = int_table.loc[TIME_ROW, FIRST_COL:].fillna("").astype(str).str.strip()
        except Exception:
            st.error(f"Interviewers Doodle: {WRONG_FORMAT_MSG}")
            st.stop()
        if int_days.eq("").all() or int_times.eq("").all():
            st.error(f"Interviewers Doodle: {WRONG_FORMAT_MSG}")
            st.stop()

        # Parse doodles (ignore NJC / Count / Total where appropriate)
        try:
            cand_av, _ = parse_doodle(cand_table, skip_names={"NJC", "Count", "Total"})
        except ValueError as e:
            st.error(f"Candidates Doodle: {WRONG_FORMAT_MSG} ({str(e)})")
            st.stop()
        except Exception:
            st.error(f"Candidates Doodle: {WRONG_FORMAT_MSG}")
            st.stop()

        if not cand_av:
            st.error(f"Candidates Doodle: No valid candidate availability found. {WRONG_FORMAT_MSG}")
            st.stop()

        try:
            int_av, _ = parse_doodle(int_table, skip_names={"NJC"})
        except ValueError as e:
            st.error(f"Interviewers Doodle: {WRONG_FORMAT_MSG} ({str(e)})")
            st.stop()
        except Exception:
            st.error(f"Interviewers Doodle: {WRONG_FORMAT_MSG}")
            st.stop()

        if not int_av:
            st.error(f"Interviewers Doodle: No valid interviewer availability found. {WRONG_FORMAT_MSG}")
            st.stop()

        # Scheduling
        seniors, juniors, inter_slots, inter_yes, inter_ifnb = classify_interviewers(int_av, mem_df)
        slot_strength, all_slots = compute_slot_strength(seniors, juniors, inter_slots)

        if not all_slots:
            st.error("No interviewer slots available after parsing. Check your Interviewers Doodle file.")
            st.stop()

        schedule = schedule_interviews(
            cand_av, seniors, juniors,
            inter_slots, inter_yes, inter_ifnb,
            all_slots, slot_strength,
            allow_back_to_back=allow_back_to_back
        )

        # Merge teams into Senior / Junior columns
        def merge_team(r):
            tt = r.get("Team Type", "")
            if "Senior + Junior" in tt:
                return r.get("Senior1"), r.get("Junior1")
            if "Senior + Senior" in tt:
                s1 = r.get('Senior1') if pd.notna(r.get('Senior1')) else None
                s2 = r.get('Senior2') if pd.notna(r.get('Senior2')) else None
                if s1 and s2:
                    return f"{s1} & {s2}", None
                return s1, None
            if "Junior + Junior" in tt:
                j1 = r.get('Junior1') if pd.notna(r.get('Junior1')) else None
                j2 = r.get('Junior2') if pd.notna(r.get('Junior2')) else None
                if j1 and j2:
                    return None, f"{j1} & {j2}"
                return None, j1
            return None, None

        if schedule is None or schedule.empty:
            final_schedule = pd.DataFrame(columns=["Candidate", "Email", "Slot", "Senior", "Junior", "Team Type"])
        else:
            merged = schedule.apply(lambda r: pd.Series(merge_team(r), index=["Senior", "Junior"]), axis=1)
            schedule = pd.concat([schedule, merged], axis=1)
            for col in ["Candidate", "Email", "Slot", "Senior", "Junior", "Team Type"]:
                if col not in schedule.columns:
                    schedule[col] = pd.NA
            final_schedule = schedule[["Candidate", "Email", "Slot", "Senior", "Junior", "Team Type"]]

        # Workload + summary
        workload = {}
        summary_full = {}

        def inc(name):
            if not name or pd.isna(name):
                return
            workload[name] = workload.get(name, 0) + 1

        def add_summary(name, candidate, slot):
            if not name or pd.isna(name):
                return
            summary_full.setdefault(name, []).append((candidate, slot))

        if not schedule.empty:
            for _, row in schedule.iterrows():
                tt = row.get("Team Type", "")
                if "Senior + Junior" in tt:
                    inc(row.get("Senior1")); inc(row.get("Junior1"))
                    add_summary(row.get("Senior1"), row.get("Candidate"), row.get("Slot"))
                    add_summary(row.get("Junior1"), row.get("Candidate"), row.get("Slot"))
                elif "Senior + Senior" in tt:
                    inc(row.get("Senior1")); inc(row.get("Senior2"))
                    add_summary(row.get("Senior1"), row.get("Candidate"), row.get("Slot"))
                    add_summary(row.get("Senior2"), row.get("Candidate"), row.get("Slot"))
                elif "Junior + Junior" in tt:
                    inc(row.get("Junior1")); inc(row.get("Junior2"))
                    add_summary(row.get("Junior1"), row.get("Candidate"), row.get("Slot"))
                    add_summary(row.get("Junior2"), row.get("Candidate"), row.get("Slot"))

        if not workload:
            workload_df = pd.DataFrame(columns=["Interviewer", "Total Interviews"])
        else:
            workload_df = (
                pd.DataFrame([{"Interviewer": k, "Total Interviews": v} for k, v in workload.items()])
                .sort_values("Total Interviews", ascending=False)
                .reset_index(drop=True)
            )

        summary_txt = io.StringIO()
        for interviewer in sorted(summary_full.keys()):
            summary_txt.write(f"{interviewer}:\n")
            for candidate, slot in sorted(summary_full[interviewer], key=lambda x: (x[1] or "")):
                summary_txt.write(f"  - {candidate}: {slot}\n")
            summary_txt.write("\n")

        # Calendar assignments (only real slots)
        assignments = []
        for _, r in final_schedule.iterrows():
            if pd.isna(r["Slot"]) or not r["Slot"]:
                continue
            day, time = parse_slot_to_day_time(r["Slot"])
            if day is None or time is None:
                continue
            interviewer_texts = []
            if r.get("Senior") and pd.notna(r.get("Senior")):
                interviewer_texts.append(str(r.get("Senior")))
            if r.get("Junior") and pd.notna(r.get("Junior")):
                interviewer_texts.append(str(r.get("Junior")))
            interviewer_text = " & ".join(interviewer_texts) if interviewer_texts else "TBD"
            assignments.append({
                "day": day,
                "time": time,
                "candidate": r["Candidate"],
                "interviewer": interviewer_text
            })

        excel_cal = build_excel_calendar(assignments)

        # Determine weekday columns to show
        if assignments:
            present_days = sorted(
                {a["day"] for a in assignments if a["day"] is not None},
                key=lambda d: ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"].index(d)
                if d in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"] else 999
            )
            weekdays = [d for d in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
                        if d in present_days]
            if not weekdays:
                weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
            cal_df = build_weekly_calendar(assignments, weekdays=weekdays)
        else:
            cal_df = None

        # Tabs
        tab_schedule, tab_workload, tab_summary, tab_calendar, tab_downloads = st.tabs([
            "Final Schedule Table",
            "Interviewer Workload",
            "Interviewer Summary (detailed)",
            "Styled Weekly Calendar",
            "Downloads"
        ])

        with tab_schedule:
            st.subheader("📄 Final Schedule Table")
            st.dataframe(final_schedule, use_container_width=True)

        with tab_workload:
            st.subheader("🗂 Interviewer Workload")
            st.dataframe(workload_df, use_container_width=True)

        with tab_summary:
            st.subheader("📝 Interviewer Summary (detailed)")
            st.text_area("Interviewer assignments (text)", value=summary_txt.getvalue(), height=240)

        with tab_calendar:
            st.subheader("🗓 Styled Weekly Calendar (Candidate — Interviewer)")
            if cal_df is None or cal_df.empty:
                st.info("No calendar could be generated from the provided files.")
            else:
                custom_css = """
                <style>
                table { border-collapse: separate; border-spacing: 14px 10px; width: 100%; }
                th { text-align: center; background: #f7f9fc; border-radius: 8px; padding: 10px; font-size: 15px; font-weight: 700; color: #0d2a56; }
                td { background: #ffffff; border-radius: 10px; vertical-align: top; padding: 8px; min-width: 180px; border: 1px solid rgba(0,0,0,0.05); }
                tr:nth-child(even) td { background: #fafbfd; }
                @media (max-width: 800px) { th { font-size: 13px; padding: 8px; min-width: 120px; } td { padding: 6px; min-width: 120px; } }
                </style>
                """
                st.markdown(custom_css, unsafe_allow_html=True)
                st.markdown(cal_df.to_html(escape=False), unsafe_allow_html=True)

        with tab_downloads:
            st.subheader("⬇️ Downloads")
            st.download_button("Download schedule.csv", final_schedule.to_csv(index=False),
                               "schedule.csv", mime="text/csv")

            try:
                cal_buf = io.BytesIO()
                with pd.ExcelWriter(cal_buf, engine='openpyxl') as writer:
                    excel_cal.to_excel(writer, index=False, sheet_name='Weekly Calendar')
                cal_buf.seek(0)
                st.download_button(
                    "Download calendar.xlsx",
                    cal_buf.getvalue(),
                    file_name="calendar.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception:
                st.download_button("Download calendar.csv", excel_cal.to_csv(index=False),
                                   "calendar.csv", mime="text/csv")

            st.download_button("Download interviewer_summary_full.txt", summary_txt.getvalue(),
                               "interviewer_summary_full.txt", mime="text/plain")

        st.success("Done — schedule, calendar, summaries, and problem flags generated 🎉")




