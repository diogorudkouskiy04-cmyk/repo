import streamlit as st
import pandas as pd
import io
import tempfile
import os
import re
from datetime import datetime

# =========================================================
# ==================  MAIN FUNCTIONS  =====================
# =========================================================

def fill_days_in_doodle(doodle, doodle_cleaned, day_row=4, first_column=2, min_rows=6):
    """
    Read an exported Doodle excel, validate expected structure, forward-fill the day row,
    and write a cleaned file. Raises ValueError with a readable message on malformed input.
    """
    table = pd.read_excel(doodle, header=None)

    # Basic structural validation
    if table.shape[0] < min_rows:
        raise ValueError("Doodle file has too few rows.")
    if table.shape[1] <= first_column:
        raise ValueError("Doodle file has too few columns.")
    # Ensure day_row exists and has at least one non-empty cell after first_column
    try:
        day_series = table.loc[day_row, first_column:]
    except Exception:
        raise ValueError("Doodle file missing expected day row.")
    if day_series.isna().all() or day_series.replace(["", " ", "nan", "NaN", "None", "none"], pd.NA).isna().all():
        raise ValueError("Doodle file missing days row content.")

    # If time row exists, validate it has something too (best-effort)
    try:
        time_series = table.loc[day_row + 1, first_column:]
        if time_series.isna().all():
            # not fatal, but warn by raising to be caught upstream
            raise ValueError("Doodle file missing time row content.")
    except Exception:
        raise ValueError("Doodle file missing expected time row.")

    # Perform forward-fill on day labels
    DAYS = day_series
    cleaned = DAYS.replace(["", " ", "nan", "NaN", "None", "none"], pd.NA)
    filled = cleaned.ffill()
    table.loc[day_row, first_column:] = filled

    # Persist cleaned workbook (no header, no index) so downstream code can read it
    table.to_excel(doodle_cleaned, index=False, header=False)


def parse_doodle(table, skip_names=None, day_row=4, first_column=2, time_row=5):
    if skip_names is None:
        skip_names = set()

    # Defensive access: ensure expected rows exist
    if table.shape[0] <= time_row:
        # return empty availability gracefully
        return {}, []

    days = table.loc[day_row, first_column:]
    times = table.loc[time_row, first_column:]
    # coerce to string and normalize blanks to pd.NA for concatenation safety
    days = days.fillna("").astype(str)
    times = times.fillna("").astype(str)

    slots = (days + " " + times).tolist()

    availability = {}
    # rows with people typically start at index 6 in your workflow; be resilient by scanning further
    for i in range(6, len(table)):
        name = table.iloc[i, 0]
        if not isinstance(name, str):
            continue
        name_clean = name.strip()
        if not name_clean:
            continue
        if name_clean.upper() in {s.upper() for s in skip_names}:
            continue

        email = table.iloc[i, 1] if table.shape[1] > 1 else None
        yes_slots, ifnb_slots = [], []

        for j, slot in enumerate(slots):
            # ensure we don't index past table width
            col_idx = first_column + j
            if col_idx >= table.shape[1]:
                break
            answer = table.iloc[i, col_idx]
            if isinstance(answer, str):
                a = answer.strip().upper()
                if a == "YES":
                    yes_slots.append(slot.strip())
                elif a in ("IF NEED BE", "IF NEEDED", "IF NEED", "IFNEEDBE"):
                    ifnb_slots.append(slot.strip())

        availability[name_clean] = {
            "email": email,
            "yes": yes_slots,
            "ifnb": ifnb_slots
        }

    return availability, slots


def classify_interviewers(interviewer_availability, member_info):
    seniors, juniors = [], []
    interviewer_slots, interviewer_yes, interviewer_ifnb = {}, {}, {}

    for name, data in interviewer_availability.items():
        # match by exact "Member Name" value
        row = member_info[member_info["Member Name"] == name]
        if row.empty:
            # try case-insensitive match as fallback
            row = member_info[member_info["Member Name"].str.lower() == name.lower()]
        if row.empty:
            # no member info found for this interviewer; skip but keep them available
            # still populate sets so algorithm can consider them (no positions known -> junior)
            interviewer_slots[name] = set(data["yes"] + data["ifnb"])
            interviewer_yes[name] = set(data["yes"])
            interviewer_ifnb[name] = set(data["ifnb"])
            juniors.append(name)
            continue

        position = str(row.iloc[0].get("Position", "")).lower()
        try:
            semester = int(row.iloc[0].get("Semesters at NJC", 0))
        except Exception:
            semester = 0

        if "board" in position or "principal" in position or semester > 2:
            seniors.append(name)
        else:
            juniors.append(name)

        interviewer_slots[name] = set(data["yes"] + data["ifnb"])
        interviewer_yes[name] = set(data["yes"])
        interviewer_ifnb[name] = set(data["ifnb"])

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
                        all_slots, slot_strength):

    def sort_slots(slots):
        # remove possible empty strings and normalize
        return sorted([s for s in slots if isinstance(s, str) and s.strip()],
                      key=lambda s: slot_strength.get(s, 0),
                      reverse=True)

    booked_s = {slot: set() for slot in all_slots}
    booked_j = {slot: set() for slot in all_slots}
    load = {name: 0 for name in seniors + juniors}
    results = []

    for candidate, cdata in candidate_availability.items():
        cand_slots_sorted = sort_slots(cdata.get("yes", []) + cdata.get("ifnb", []))

        for slot in cand_slots_sorted:
            s_yes = [s for s in seniors if slot in interviewer_yes.get(s, set()) and s not in booked_s.get(slot, set())]
            j_yes = [j for j in juniors if slot in interviewer_yes.get(j, set()) and j not in booked_j.get(slot, set())]
            s_ifnb = [s for s in seniors if slot in interviewer_ifnb.get(s, set()) and s not in booked_s.get(slot, set())]
            j_ifnb = [j for j in juniors if slot in interviewer_ifnb.get(j, set()) and j not in booked_j.get(slot, set())]

            for lst in [s_yes, j_yes, s_ifnb, j_ifnb]:
                lst.sort(key=lambda x: load.get(x, 0))

            # Senior + Junior (both yes)
            if s_yes and j_yes:
                S, J = s_yes[0], j_yes[0]
                booked_s.setdefault(slot, set()).add(S); booked_j.setdefault(slot, set()).add(J)
                load[S] = load.get(S, 0) + 1; load[J] = load.get(J, 0) + 1
                results.append({"Candidate": candidate, "Email": cdata.get("email"),
                                "Slot": slot, "Team Type": "Senior + Junior (YES)",
                                "Senior1": S, "Senior2": None, "Junior1": J, "Junior2": None})
                break

            # Senior + Senior (both yes)
            if len(s_yes) >= 2:
                S1, S2 = s_yes[:2]
                booked_s.setdefault(slot, set()).update({S1, S2})
                load[S1] = load.get(S1, 0) + 1; load[S2] = load.get(S2, 0) + 1
                results.append({"Candidate": candidate, "Email": cdata.get("email"),
                                "Slot": slot, "Team Type": "Senior + Senior (YES)",
                                "Senior1": S1, "Senior2": S2, "Junior1": None, "Junior2": None})
                break

            s_mix = s_yes + s_ifnb
            j_mix = j_yes + j_ifnb
            # fallback Senior + Junior
            if s_mix and j_mix:
                S = sorted(s_mix, key=lambda x: load.get(x, 0))[0]
                J = sorted(j_mix, key=lambda x: load.get(x, 0))[0]
                booked_s.setdefault(slot, set()).add(S); booked_j.setdefault(slot, set()).add(J)
                load[S] = load.get(S, 0) + 1; load[J] = load.get(J, 0) + 1
                results.append({"Candidate": candidate, "Email": cdata.get("email"),
                                "Slot": slot, "Team Type": "Senior + Junior (fallback)",
                                "Senior1": S, "Senior2": None, "Junior1": J, "Junior2": None})
                break

            # fallback Senior + Senior
            if len(s_mix) >= 2:
                S1, S2 = sorted(s_mix, key=lambda x: load.get(x, 0))[:2]
                booked_s.setdefault(slot, set()).update({S1, S2})
                load[S1] = load.get(S1, 0) + 1; load[S2] = load.get(S2, 0) + 1
                results.append({"Candidate": candidate, "Email": cdata.get("email"),
                                "Slot": slot, "Team Type": "Senior + Senior (fallback)",
                                "Senior1": S1, "Senior2": S2, "Junior1": None, "Junior2": None})
                break

            # fallback Junior + Junior
            jj_mix = j_yes + j_ifnb
            if len(jj_mix) >= 2:
                J1, J2 = sorted(jj_mix, key=lambda x: load.get(x, 0))[:2]
                booked_j.setdefault(slot, set()).update({J1, J2})
                load[J1] = load.get(J1, 0) + 1; load[J2] = load.get(J2, 0) + 1
                results.append({"Candidate": candidate, "Email": cdata.get("email"),
                                "Slot": slot, "Team Type": "Junior + Junior (fallback)",
                                "Senior1": None, "Senior2": None, "Junior1": J1, "Junior2": J2})
                break

    return pd.DataFrame(results)

# =========================================================
# ===============  SLOT PARSER FOR CALENDAR  ==============
# =========================================================

WEEKDAY_MAP = {
    "Mon": "Monday", "Tue": "Tuesday", "Wed": "Wednesday",
    "Thu": "Thursday", "Fri": "Friday", "Sat": "Saturday",
    "Sun": "Sunday",
    "Monday":"Monday","Tuesday":"Tuesday","Wednesday":"Wednesday",
    "Thursday":"Thursday","Friday":"Friday","Saturday":"Saturday",
    "Sunday":"Sunday"
}

time_re = re.compile(r'(\d{1,2}:\d{2}\s*[APap][Mm])')

def parse_slot_to_day_time(slot_text):
    """
    Returns (day_name, 'HH:MM') given a slot like "Mon 9:00 AM" or "Monday 09:00".
    If it can't parse day or time, returns (None, None or None).
    """
    if not isinstance(slot_text, str):
        return None, None

    # normalize spaces
    s = " ".join(slot_text.split())
    tokens = s.split()
    if not tokens:
        return None, None

    # day might be first token or first token with comma (e.g., "Mon," or "Mon")
    day_token = tokens[0].rstrip(",")
    day = WEEKDAY_MAP.get(day_token, None)

    # find time with regex (AM/PM) or 24h hh:mm
    m = time_re.search(s)
    if m:
        start_time_str = m.group(1).upper().replace(" ", "")
        start_time_norm = re.sub(r'([AP]M)$', r' \1', start_time_str)
        try:
            dt = datetime.strptime(start_time_norm, "%I:%M %p")
            return day, dt.strftime("%H:%M")
        except:
            return day, None

    # fallback: look for hh:mm (24h)
    m2 = re.search(r'(\d{1,2}:\d{2})', s)
    if m2:
        hhmm = m2.group(1)
        parts = hhmm.split(':')
        return day, parts[0].zfill(2) + ":" + parts[1][:2]

    return day, None

# =========================================================
# ===============  WEEKLY CALENDAR BUILDER  ===============
# =========================================================

# pastel palette for event backgrounds
PALETTE = ["#e8f4ff", "#e4ffe8", "#fff4e5", "#f9e6ff", "#ffecec", "#f0f7ff", "#fff0f5"]

def build_weekly_calendar(assignments, weekdays=None):
    """
    assignments: list of dicts {day:'Monday', time:'09:00', candidate:, interviewer: 'A & B' }
    Returns DataFrame with HTML strings in cells (properly formatted, no raw \n)
    """

    if weekdays is None:
        weekdays = ["Monday","Tuesday","Wednesday","Thursday","Friday"]

    # helper: normalize time inputs to HH:MM
    def normalize_time(t):
        if not isinstance(t, str) or not t:
            return None
        s = t.strip()
        # AM/PM form
        m = re.match(r'^\s*(\d{1,2}):(\d{2})\s*([AaPp][Mm])\s*$', s)
        if m:
            hhmm = f"{int(m.group(1)):02d}:{m.group(2)}"
            try:
                dt = datetime.strptime(f"{m.group(1)}:{m.group(2)} {m.group(3).upper()}", "%I:%M %p")
                return dt.strftime("%H:%M")
            except:
                return hhmm
        # 24-hour hh:mm
        m2 = re.match(r'^\s*(\d{1,2}):(\d{2})\s*$', s)
        if m2:
            return f"{int(m2.group(1)):02d}:{m2.group(2)}"
        return None

    # collect times and convert to minutes-from-midnight for sorting/clustering
    raw_times = {normalize_time(a["time"]) for a in assignments if a.get("time")}
    mins = sorted({int(t[:2]) * 60 + int(t[3:]) for t in raw_times if t})

    # If no times found, fill default 08:00-19:00
    if not mins:
        times = [f"{h:02d}:00" for h in range(8, 20)]
    else:
        # cluster times and insert visual break rows for large gaps
        cleaned = []
        last = None
        for m in mins:
            if last is None:
                cleaned.append(m)
            elif m - last <= 45:
                cleaned.append(m)
            else:
                # large gap -> mark a break (None) then continue
                cleaned.append(None)
                cleaned.append(m)
            last = m
        # convert back to HH:MM strings, keeping "" for break rows so empty cells render subtly
        times = [(f"{m//60:02d}:{m%60:02d}") if m is not None else "" for m in cleaned]

    # Build empty table: times x weekdays
    table = pd.DataFrame("", index=times, columns=weekdays)

    # bucket events per cell as structured entries
    buckets = {d: {t: [] for t in times} for d in weekdays}

    # improved panel splitting: split on & / , only (not the word 'and')
    split_pattern = re.compile(r"\s*&\s*|\s*/\s*|\s*,\s*")

    for a in assignments:
        d = a.get("day")
        t = normalize_time(a.get("time"))
        if d not in weekdays or not t:
            continue

        raw_panel = a.get("interviewer") or ""
        # DON'T split on the word 'and' (multi-word names preserved)
        panel = [p.strip() for p in split_pattern.split(raw_panel) if p.strip()]
        if not panel:
            panel = ["TBD"]

        buckets[d][t].append({"candidate": a.get("candidate", "TBD"), "panel": panel})

    # format entries into pretty HTML blocks
    def format_cell_entries(entries, slot_index=0):
        if not entries:
            return ""
        html_blocks = []
        for i, e in enumerate(entries):
            bg = PALETTE[(slot_index + i) % len(PALETTE)]
            candidate_clean = e["candidate"].replace("\n", "").strip()
            panel_clean = [p.replace("\n", "").strip() for p in e["panel"]]

            cand_html = (f"<div style='font-weight:700;color:#183a6e;margin-bottom:4px'>{candidate_clean}</div>")
            panel_html = ("<div style='color:#333;font-size:13px;margin-bottom:2px;'>" + " / ".join(panel_clean) + "</div>")

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

# =========================================================
# ===============  EXCEL CALENDAR (CLEAN)  ===============
# =========================================================

def build_excel_calendar(assignments):
    """
    Build a simple, tabular DataFrame suitable for Excel export with columns:
    Day | Time | Candidate | Interviewers
    """
    rows = []
    for a in assignments:
        # normalize text fields to avoid HTML or newlines
        day = a.get("day") or ""
        time = a.get("time") or ""
        candidate = str(a.get("candidate") or "").replace("\n", " ").strip()
        interviewer = str(a.get("interviewer") or "").replace("\n", " ").strip()
        rows.append({"Day": day, "Time": time, "Candidate": candidate, "Interviewers": interviewer})
    df = pd.DataFrame(rows)
    # try to sort by day order then time (if times are HH:MM they sort lexicographically)
    day_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    df["DayOrder"] = df["Day"].apply(lambda x: day_order.index(x) if x in day_order else 999)
    df = df.sort_values(["DayOrder", "Time"]).drop(columns=["DayOrder"]).reset_index(drop=True)
    return df

# =========================================================
# ===================  STREAMLIT APP  =====================
# =========================================================

# Use the screenshot path in the session history (developer provided file)
SAMPLE_IMAGE_PATH = "/mnt/data/Screenshot 2025-11-23 at 16.38.00.png"

st.set_page_config(page_title="Interview Scheduler", layout="wide")
st.title("📅 Interview Scheduling + Weekly Calendar")
st.markdown("Upload the two Excel files exported from Doodle and the Member Info sheet. After running you'll get the schedule table, a weekly calendar, and an interviewer summary.")

# Upload UI
col1, col2, col3 = st.columns(3)
with col1:
    cand_file = st.file_uploader("Upload Candidates Doodle", type=["xlsx"])
with col2:
    int_file = st.file_uploader("Upload Interviewers Doodle", type=["xlsx"])
with col3:
    mem_file = st.file_uploader("Upload Member Info Sheet", type=["xlsx"])

# show user screenshot example if exists
if os.path.exists(SAMPLE_IMAGE_PATH):
    st.markdown("**Example calendar screenshot you provided:**")
    st.image(SAMPLE_IMAGE_PATH, use_column_width=True)

# A small helper to show the exact phrasing requested when a file has wrong format
WRONG_FORMAT_MSG = "this specific document is not on the right format, please input the correct version."

if st.button("Run Scheduling"):
    if not (cand_file and int_file and mem_file):
        st.error("Please upload all files.")
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

        # Attempt to preprocess doodles; if they raise or appear malformed, show a friendly message
        try:
            fill_days_in_doodle(cand_path, cand_clean)
        except Exception as e:
            # surface the friendly message (do not expose raw exception)
            st.error(f"Candidates Doodle: {WRONG_FORMAT_MSG}")
            st.stop()

        try:
            fill_days_in_doodle(int_path, int_clean)
        except Exception as e:
            st.error(f"Interviewers Doodle: {WRONG_FORMAT_MSG}")
            st.stop()

        # Read the cleaned files and member sheet
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

        try:
            mem_df = pd.read_excel(mem_path)
        except Exception:
            st.error(f"Member Info Sheet: {WRONG_FORMAT_MSG}")
            st.stop()

        # Validate member info columns (common expected columns)
        required_mem_cols = {"Member Name", "Position", "Semesters at NJC"}
        if not required_mem_cols.issubset(set(mem_df.columns)):
            st.error(f"Member Info Sheet: {WRONG_FORMAT_MSG}")
            st.stop()

        # Quick structural checks for doodle tables: make sure day and time rows have some content
        try:
            if cand_table.shape[0] <= 5 or cand_table.shape[1] <= 3 or cand_table.loc[4, 2:].replace(["", " ", "nan", "NaN", "None", "none"], pd.NA).isna().all():
                st.error(f"Candidates Doodle: {WRONG_FORMAT_MSG}")
                st.stop()
        except Exception:
            st.error(f"Candidates Doodle: {WRONG_FORMAT_MSG}")
            st.stop()

        try:
            if int_table.shape[0] <= 5 or int_table.shape[1] <= 3 or int_table.loc[4, 2:].replace(["", " ", "nan", "NaN", "None", "none"], pd.NA).isna().all():
                st.error(f"Interviewers Doodle: {WRONG_FORMAT_MSG}")
                st.stop()
        except Exception:
            st.error(f"Interviewers Doodle: {WRONG_FORMAT_MSG}")
            st.stop()

        # If we reach here, files look OK-ish — proceed with parsing/scheduling
        cand_av, _ = parse_doodle(cand_table, skip_names={"NJC"})
        int_av, all_slots_list = parse_doodle(int_table)

        seniors, juniors, inter_slots, inter_yes, inter_ifnb = classify_interviewers(int_av, mem_df)
        slot_strength, all_slots = compute_slot_strength(seniors, juniors, inter_slots)

        schedule = schedule_interviews(
            cand_av, seniors, juniors,
            inter_slots, inter_yes, inter_ifnb,
            all_slots, slot_strength
        )

        # Merge teams (Senior / Junior columns)
        def merge_team(r):
            if "Senior + Junior" in r.get("Team Type", ""):
                return r.get("Senior1"), r.get("Junior1")
            if "Senior + Senior" in r.get("Team Type", ""):
                s1 = r.get('Senior1') if pd.notna(r.get('Senior1')) else None
                s2 = r.get('Senior2') if pd.notna(r.get('Senior2')) else None
                if s1 and s2:
                    return f"{s1} & {s2}", None
                return s1, None
            if "Junior + Junior" in r.get("Team Type", ""):
                j1 = r.get('Junior1') if pd.notna(r.get('Junior1')) else None
                j2 = r.get('Junior2') if pd.notna(r.get('Junior2')) else None
                if j1 and j2:
                    return None, f"{j1} & {j2}"
                return None, j1
            return None, None

        if schedule is None or schedule.empty:
            # ensure final_schedule exists even if empty
            final_schedule = pd.DataFrame(columns=["Candidate", "Email", "Slot", "Senior", "Junior", "Team Type"])
        else:
            merged = schedule.apply(lambda r: pd.Series(merge_team(r), index=["Senior", "Junior"]), axis=1)
            schedule = pd.concat([schedule, merged], axis=1)
            final_schedule = schedule[["Candidate", "Email", "Slot", "Senior", "Junior", "Team Type"]]

        # --------------------------------
        # Interviewer workload + summary
        # --------------------------------
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
                st.warning("No valid interview assignments found. The calendar may be empty due to an incorrect Doodle file.")
                workload_df = pd.DataFrame(columns=["Interviewer", "Total Interviews"])
            else:
                workload_df = (
                    pd.DataFrame(
                        [{"Interviewer": k, "Total Interviews": v} for k, v in workload.items()]
                    )
                    .sort_values("Total Interviews", ascending=False)
                    .reset_index(drop=True)
                )

        # create interviewer summary text
        summary_txt = io.StringIO()
        for interviewer in sorted(summary_full.keys()):
            summary_txt.write(f"{interviewer}:\n")
            for candidate, slot in sorted(summary_full[interviewer], key=lambda x: (x[1] or "")):
                summary_txt.write(f"  - {candidate}: {slot}\n")
            summary_txt.write("\n")

        # Build assignments for calendar
        assignments = []
        for _, r in final_schedule.iterrows():
            day, time = parse_slot_to_day_time(r["Slot"])
            if day is None or time is None:
                # skip slots we cannot parse into day/time
                continue

            # combine interviewer names (prefer merged Senior/Junior if available)
            interviewer_texts = []
            if r.get("Senior") and pd.notna(r.get("Senior")):
                interviewer_texts.append(str(r.get("Senior")))
            else:
                if r.get("Senior1") and pd.notna(r.get("Senior1")):
                    interviewer_texts.append(str(r.get("Senior1")))
                if r.get("Senior2") and pd.notna(r.get("Senior2")):
                    interviewer_texts.append(str(r.get("Senior2")))

            if r.get("Junior") and pd.notna(r.get("Junior")):
                interviewer_texts.append(str(r.get("Junior")))
            else:
                if r.get("Junior1") and pd.notna(r.get("Junior1")):
                    interviewer_texts.append(str(r.get("Junior1")))
                if r.get("Junior2") and pd.notna(r.get("Junior2")):
                    interviewer_texts.append(str(r.get("Junior2")))

            interviewer_text = " & ".join(interviewer_texts) if interviewer_texts else "TBD"

            assignments.append({
                "day": day,
                "time": time,
                "candidate": r["Candidate"],
                "interviewer": interviewer_text
            })

        # Determine weekdays to show (keep Mon-Sun order)
        present_days = sorted({a["day"] for a in assignments if a["day"] is not None},
                              key=lambda d: ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"].index(d) if d in ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"] else 999)
        weekdays = [d for d in ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"] if d in present_days]
        if not weekdays:
            weekdays = ["Monday","Tuesday","Wednesday","Thursday","Friday"]

        # Build calendar DataFrame (styled HTML for display) and clean Excel calendar
        cal_df = build_weekly_calendar(assignments, weekdays=weekdays)
        excel_cal = build_excel_calendar(assignments)

        # Display outputs
        st.subheader("📄 Final Schedule Table")
        st.dataframe(final_schedule, use_container_width=True)

        st.subheader("🗂 Interviewer Workload")
        st.dataframe(workload_df, use_container_width=True)

        st.subheader("📝 Interviewer Summary (detailed)")
        st.text_area("Interviewer assignments (text)", value=summary_txt.getvalue(), height=240)

        st.subheader("🗓 Styled Weekly Calendar (Candidate — Interviewer)")
        # Add a bit of CSS to make the HTML table look nicer
        custom_css = """
        <style>
        table {
            border-collapse: separate;
            border-spacing: 14px 10px;
            width: 100%;
        }

        th {
            text-align: center;
            background: #f7f9fc;
            border-radius: 8px;
            padding: 10px;
            font-size: 15px;
            font-weight: 700;
            color: #0d2a56;
        }

        td {
            background: #ffffff;
            border-radius: 10px;
            vertical-align: top;
            padding: 8px;
            min-width: 180px;
            border: 1px solid rgba(0,0,0,0.05);
        }

        tr:nth-child(even) td {
            background: #fafbfd;
        }

        /* small screens: reduce padding */
        @media (max-width: 800px) {
            th { font-size: 13px; padding: 8px; }
            td { padding: 6px; min-width: 120px; }
        }
        </style>
        """
        st.markdown(custom_css, unsafe_allow_html=True)
        st.markdown(cal_df.to_html(escape=False), unsafe_allow_html=True)

        # Downloads
        # schedule CSV
        try:
            st.download_button("Download schedule.csv", final_schedule.to_csv(index=False), "schedule.csv", mime="text/csv")
        except Exception:
            # fallback: show as plain text download
            st.download_button("Download schedule.csv", final_schedule.to_csv(index=False), "schedule.csv")

        # Create an Excel file for the calendar (clean, no HTML)
        try:
            cal_buf = io.BytesIO()
            with pd.ExcelWriter(cal_buf, engine='openpyxl') as writer:
                # write the clean calendar DataFrame; index=False ensures Time is a column
                excel_cal.to_excel(writer, index=False, sheet_name='Weekly Calendar')
            cal_buf.seek(0)
            st.download_button(
                "Download calendar.xlsx",
                cal_buf.getvalue(),
                file_name="calendar.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception:
            # fallback to CSV if Excel writer not available
            st.download_button("Download calendar.csv", excel_cal.to_csv(index=False), "calendar.csv", mime="text/csv")

        # interviewer summary text
        st.download_button("Download interviewer_summary_full.txt", summary_txt.getvalue(), "interviewer_summary_full.txt", mime="text/plain")

        st.success("Done — schedule, calendar and summaries generated 🎉")


