"""
DG Session & Attendance Sheet Generator  v3.0
==============================================
• User configures Units, TLO range, Timing, Faculty ID directly in UI
• Attendance dates auto-loaded and shown day-wise in editable table
• Session sheet filled from UI table values
• Day-wise attendance files generated per date
"""

import io
import re
import zipfile
import warnings
from collections import Counter
from copy import copy
from datetime import datetime, date, timedelta

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="DG Sheet Generator",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ══════════════════════════════════════════════════════════════════════
# CSS
# ══════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
.main{padding:0.5rem 1.5rem}
.app-title{
    background:linear-gradient(135deg,#1a2f5a 0%,#2471a3 100%);
    color:white;padding:1.4rem 2rem;border-radius:12px;
    margin-bottom:1.2rem;text-align:center;
}
.app-title h1{margin:0;font-size:1.7rem;font-weight:800}
.app-title p{margin:.3rem 0 0;opacity:.88;font-size:.92rem}
.section{background:#f8fafd;border:1px solid #d0dff0;border-radius:10px;
    padding:1rem 1.2rem;margin-bottom:1rem}
.sec-title{font-size:1rem;font-weight:700;color:#1a2f5a;
    border-bottom:2px solid #2471a3;padding-bottom:.4rem;margin-bottom:.8rem}
.info-box{background:#eaf4fb;border-left:4px solid #2471a3;
    padding:.6rem 1rem;border-radius:0 6px 6px 0;margin:.4rem 0;font-size:.85rem}
.success-box{background:#d4edda;border-left:4px solid #28a745;
    padding:.6rem 1rem;border-radius:0 6px 6px 0;margin:.4rem 0;font-size:.85rem}
.warn-box{background:#fff8e1;border-left:4px solid #f39c12;
    padding:.6rem 1rem;border-radius:0 6px 6px 0;margin:.4rem 0;font-size:.85rem}
.error-box{background:#fde8e8;border-left:4px solid #e74c3c;
    padding:.6rem 1rem;border-radius:0 6px 6px 0;margin:.4rem 0;font-size:.85rem}
.badge-ok{background:#d4edda;color:#155724;padding:.15rem .55rem;
    border-radius:10px;font-size:.75rem;font-weight:700}
.stDownloadButton>button{
    background:#1a2f5a !important;color:white !important;
    border-radius:8px !important;font-weight:600 !important;width:100% !important}
.stDownloadButton>button:hover{background:#2471a3 !important}
div[data-testid="stDataEditor"]{border:1px solid #d0dff0;border-radius:8px}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
# ─── DATE PARSER ─────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════

def parse_date(raw) -> date | None:
    if raw is None:
        return None
    try:
        if pd.isna(raw):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    s = str(raw).strip()
    if s.lower() in ("nat", "nan", "none", "", "pd.nat"):
        return None
    s = re.sub(r"\([^)]*\)", "", s)
    s = re.sub(r"\b(sub|Sub|SUB)\b", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    if not s:
        return None
    for fmt in ("%d %b %Y", "%d %B %Y", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d %b %y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    try:
        r = pd.to_datetime(s, dayfirst=True)
        return None if pd.isna(r) else r.date()
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════════════
# ─── BEST SHEET PICKER ───────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════

def best_sheet_df(file_bytes: bytes, hints: list = None) -> pd.DataFrame:
    xf = pd.ExcelFile(io.BytesIO(file_bytes))
    names = xf.sheet_names
    if len(names) == 1:
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=names[0], header=None)
    if hints:
        for kw in hints:
            for s in names:
                if kw.lower() in s.lower():
                    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, header=None)
    best, most = names[0], -1
    for s in names:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, header=None)
        n = int(df.notna().sum().sum())
        if n > most:
            most, best = n, s
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=best, header=None)


# ══════════════════════════════════════════════════════════════════════
# ─── PARSE ATTENDANCE SHEET → DATES + STUDENTS ───────────────────────
# ══════════════════════════════════════════════════════════════════════

def parse_attendance(file_bytes: bytes) -> dict:
    df_raw = best_sheet_df(file_bytes, hints=["attendance", "student"])

    # Find date row
    date_row_idx = date_col_start = None
    for i, row in df_raw.iterrows():
        hits, first = [], None
        for j, v in enumerate(row):
            d = parse_date(v)
            if d:
                hits.append(d)
                first = first if first is not None else j
        if len(hits) >= 2:
            date_row_idx = i
            date_col_start = first
            break

    if date_row_idx is None:
        raise ValueError("❌ Attendance sheet mein dates nahi mili. Date row check karo.")

    date_row = df_raw.iloc[date_row_idx]
    col_to_date: dict = {}
    dates_ordered: list = []
    seen: set = set()
    for j in range(date_col_start, df_raw.shape[1]):
        d = parse_date(date_row.iloc[j])
        if d:
            col_to_date[j] = d
            if d not in seen:
                dates_ordered.append(d)
                seen.add(d)

    # Header row for student columns
    name_col = enroll_col = None
    header_row_idx = None
    for i in range(date_row_idx):
        row = df_raw.iloc[i]
        row_str = " ".join(str(v).lower() for v in row if pd.notna(v))
        if any(kw in row_str for kw in ["name", "enrol", "roll", "sr."]):
            header_row_idx = i
            for j, v in enumerate(row):
                s = str(v).lower().strip()
                if "name" in s and name_col is None:
                    name_col = j
                elif "enrol" in s and enroll_col is None:
                    enroll_col = j
            break

    # Meta
    meta: dict = {}
    for i in range(header_row_idx or 0):
        row = df_raw.iloc[i]
        for j in range(len(row) - 1):
            k, v = str(row.iloc[j]).strip(), str(row.iloc[j + 1]).strip()
            if k not in ("nan", "") and v not in ("nan", ""):
                meta[k] = v

    # Students
    students: list = []
    for i in range(date_row_idx + 1, df_raw.shape[0]):
        row = df_raw.iloc[i]
        if row.notna().sum() < 3:
            continue
        name = str(row.iloc[name_col]).strip() if name_col is not None and pd.notna(row.iloc[name_col]) else ""
        if not name or name.lower() in ("nan", "none", ""):
            continue
        enroll = str(row.iloc[enroll_col]).strip() if enroll_col is not None and pd.notna(row.iloc[enroll_col]) else ""
        att = {}
        for col_j, d in col_to_date.items():
            val = row.iloc[col_j]
            try:
                att[d] = int(float(val)) if not pd.isna(val) else None
            except Exception:
                att[d] = None
        students.append({"name": name, "enrollment": enroll, "attendance": att})

    return {"dates": dates_ordered, "students": students, "meta": meta}


# ══════════════════════════════════════════════════════════════════════
# ─── BUILD SESSION TABLE FROM UI INPUTS ──────────────────────────────
# ══════════════════════════════════════════════════════════════════════

DAY_MAP = {
    "Monday": 0, "Tuesday": 1, "Wednesday": 2,
    "Thursday": 3, "Friday": 4, "Saturday": 5, "Sunday": 6
}

def build_session_table(
    dates: list,
    units: list,           # [{"name": str, "sessions": int}, ...]
    day_timings: dict,     # {weekday_name: "HH:MM - HH:MM"}
    tlo_start: int,
    tlo_end: int,
) -> pd.DataFrame:
    """
    Auto-build the session table that user will see and can edit.
    Maps each date → unit + session number + timing + TLO
    """
    # Expand units into per-session list
    session_list = []
    for u in units:
        uname = u["name"].strip() or "Unnamed Unit"
        count = max(1, u["sessions"])
        for s in range(1, count + 1):
            session_list.append({
                "module": uname,
                "session_label": f"Session {s} – {uname}",
            })

    tlo_range = list(range(tlo_start, tlo_end + 1)) if tlo_end >= tlo_start else [tlo_start]

    rows = []
    for idx, d in enumerate(dates):
        if d is None:
            continue
        day_name = d.strftime("%A")
        raw_timing = day_timings.get(day_name, "10:00 - 11:00")
        times = re.findall(r"\d{1,2}:\d{2}", raw_timing)
        start_t = times[0] if len(times) >= 1 else "10:00"
        end_t   = times[1] if len(times) >= 2 else "11:00"

        sess = session_list[idx] if idx < len(session_list) else {
            "module": "Extra Session",
            "session_label": "Extra Session",
        }
        tlo_val = tlo_range[idx % len(tlo_range)]

        rows.append({
            "Sr": idx + 1,
            "Date": d.strftime("%d %b %Y"),
            "Day": day_name,
            "Module Name": sess["module"],
            "Session / Topic": sess["session_label"],
            "Start Time": f"{d} {start_t}:00",
            "End Time":   f"{d} {end_t}:00",
            "TLO": f"TLO{tlo_val}",
            "Mandatory": "TRUE",
        })

    return pd.DataFrame(rows)


# ══════════════════════════════════════════════════════════════════════
# ─── CELL STYLE COPY ─────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════

def copy_style(src, dst):
    if src.has_style:
        dst.font       = copy(src.font)
        dst.border     = copy(src.border)
        dst.fill       = copy(src.fill)
        dst.alignment  = copy(src.alignment)
        dst.protection = copy(src.protection)
        dst.number_format = src.number_format


# ══════════════════════════════════════════════════════════════════════
# ─── GENERATE SESSION SHEET FROM EDITED TABLE ────────────────────────
# ══════════════════════════════════════════════════════════════════════

def generate_session_sheet(
    template_bytes: bytes,
    session_df: pd.DataFrame,
    faculty_reg_id: str,
) -> bytes:
    """
    Fill DG Session Sheet template row by row from the edited DataFrame.
    Columns used: Module Name, Start Time, End Time, Session/Topic, TLO, Mandatory
    """
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # Detect column positions from header row
    COL_HINTS = {
        "module":      ["module"],
        "start":       ["start"],
        "end":         ["end"],
        "title":       ["title"],
        "description": ["description", "desc"],
        "mandatory":   ["mandatory"],
        "tlo":         ["tlo"],
        "faculty":     ["faculty", "teaching", "registration"],
    }
    col_map: dict = {}
    header_row = 1
    for r in range(1, min(6, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or "").lower().strip()
            if val:
                for key, hints in COL_HINTS.items():
                    if key not in col_map and any(h in val for h in hints):
                        col_map[key] = c
                        header_row = r

    data_start = header_row + 1

    # Capture style from first data row
    style: dict = {}
    if ws.max_row >= data_start:
        for c in range(1, ws.max_column + 1):
            style[c] = ws.cell(data_start, c)

    # Clear old rows
    for r in range(ws.max_row, header_row, -1):
        ws.delete_rows(r)

    fac_id = faculty_reg_id.strip() or "IILMGG006412025"

    for i, row in session_df.iterrows():
        rn = data_start + i
        row_vals = {
            col_map.get("module"):      str(row.get("Module Name", "")),
            col_map.get("start"):       str(row.get("Start Time", "")),
            col_map.get("end"):         str(row.get("End Time", "")),
            col_map.get("title"):       str(row.get("Session / Topic", "")),
            col_map.get("description"): str(row.get("Session / Topic", "")),
            col_map.get("mandatory"):   str(row.get("Mandatory", "TRUE")),
            col_map.get("tlo"):         str(row.get("TLO", "TLO1")),
            col_map.get("faculty"):     fac_id,
        }
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(rn, c)
            if c in style:
                copy_style(style[c], cell)
            if c in row_vals and row_vals[c] not in (None, "None", "nan"):
                cell.value = row_vals[c]

    # Auto-fit
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        max_w = max((len(str(cell.value or "")) for cell in col), default=10)
        ws.column_dimensions[letter].width = min(max_w + 4, 55)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════
# ─── GENERATE DAY-WISE ATTENDANCE ────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════

def generate_daywise_attendance(
    template_bytes: bytes,
    att_data: dict,
    session_date: date,
) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    header_row = 1
    email_col = regid_col = attend_col = None
    for r in range(1, min(5, ws.max_row + 1)):
        vals = [str(ws.cell(r, c).value or "").lower() for c in range(1, ws.max_column + 1)]
        if any(kw in " ".join(vals) for kw in ["email", "registration", "attendance"]):
            header_row = r
            for ci, v in enumerate(vals, 1):
                if "email" in v:                        email_col  = ci
                elif "registration" in v or "reg" in v: regid_col  = ci
                elif "attendance" in v:                 attend_col = ci
            break

    # Email lookup
    reg_to_email: dict = {}
    try:
        tdf = pd.read_excel(io.BytesIO(template_bytes), header=header_row - 1)
        tdf.columns = [str(c).strip() for c in tdf.columns]
        ec = next((c for c in tdf.columns if "email" in c.lower()), None)
        rc = next((c for c in tdf.columns if "reg" in c.lower()), None)
        if ec and rc:
            for _, row in tdf.iterrows():
                em = str(row[ec]).strip()
                rg = str(row[rc]).strip()
                if em and em.lower() not in ("nan", "none"):
                    reg_to_email[rg] = em
    except Exception:
        pass

    # Style
    style: dict = {}
    if ws.max_row >= header_row + 1:
        for c in range(1, ws.max_column + 1):
            style[c] = ws.cell(header_row + 1, c)

    for r in range(ws.max_row, header_row, -1):
        ws.delete_rows(r)

    GREEN_F = PatternFill("solid", start_color="C6EFCE")
    RED_F   = PatternFill("solid", start_color="FFC7CE")
    GREEN_T = Font(color="006100", bold=True)
    RED_T   = Font(color="9C0006", bold=True)

    for idx, student in enumerate(att_data["students"]):
        rn     = header_row + 1 + idx
        enroll = student.get("enrollment", "")
        att_v  = student["attendance"].get(session_date)
        status = "PRESENT" if att_v == 1 else "ABSENT"
        email  = reg_to_email.get(enroll, "")
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(rn, c)
            if c in style:
                copy_style(style[c], cell)
            if c == email_col:
                cell.value = email
            elif c == regid_col:
                cell.value = enroll
            elif c == attend_col:
                cell.value     = status
                cell.fill      = GREEN_F if status == "PRESENT" else RED_F
                cell.font      = GREEN_T if status == "PRESENT" else RED_T
                cell.alignment = Alignment(horizontal="center", vertical="center")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════
# ─── ZIP BUILDER ─────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════

def build_zip(session_bytes: bytes, daywise: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("session_sheet.xlsx", session_bytes)
        for fname, fb in daywise.items():
            zf.writestr(f"attendance/{fname}", fb)
    buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════
# ─── MAIN UI ─────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════

def main():
    st.markdown("""
    <div class="app-title">
        <h1>📋 DG Session & Attendance Sheet Generator</h1>
        <p>Dates attendance se · Units aap define karo · Table mein edit karo · Generate!</p>
    </div>
    """, unsafe_allow_html=True)

    if "att_data"      not in st.session_state: st.session_state.att_data = None
    if "session_df"    not in st.session_state: st.session_state.session_df = None
    if "results"       not in st.session_state: st.session_state.results = None

    # ══════════════════════════════════════════════════════════════════
    # ROW 1: FILE UPLOADS
    # ══════════════════════════════════════════════════════════════════
    st.markdown('<div class="sec-title">📂 Step 1 — Files Upload Karo</div>', unsafe_allow_html=True)
    fc1, fc2, fc3 = st.columns(3)

    with fc1:
        att_file = st.file_uploader("📅 Master Attendance Sheet (.xlsx)", type=["xlsx"], key="att")
        if att_file:
            st.markdown(f'<span class="badge-ok">✅ {att_file.name}</span>', unsafe_allow_html=True)

    with fc2:
        ses_tpl = st.file_uploader("📋 DG Session Format Template (.xlsx)", type=["xlsx"], key="stpl")
        if ses_tpl:
            st.markdown(f'<span class="badge-ok">✅ {ses_tpl.name}</span>', unsafe_allow_html=True)

    with fc3:
        att_tpl = st.file_uploader("🗂️ DG Attendance Format Template (.xlsx)", type=["xlsx"], key="atpl")
        if att_tpl:
            st.markdown(f'<span class="badge-ok">✅ {att_tpl.name}</span>', unsafe_allow_html=True)

    # ── Auto-load attendance when uploaded ────────────────────────────
    if att_file:
        try:
            if (st.session_state.att_data is None or
                    st.session_state.get("_last_att") != att_file.name):
                st.session_state.att_data = parse_attendance(att_file.read())
                st.session_state._last_att = att_file.name
                st.session_state.session_df = None  # reset table
        except Exception as ex:
            st.markdown(f'<div class="error-box">❌ Attendance parse error: {ex}</div>',
                        unsafe_allow_html=True)

    att_data = st.session_state.att_data
    dates = att_data["dates"] if att_data else []

    if dates:
        st.markdown(
            f'<div class="success-box">✅ Attendance loaded: '
            f'<b>{len(dates)} session dates</b> · '
            f'<b>{len(att_data["students"])} students</b></div>',
            unsafe_allow_html=True)

    st.markdown("---")

    # ══════════════════════════════════════════════════════════════════
    # ROW 2: CONFIGURATION (two columns)
    # ══════════════════════════════════════════════════════════════════
    st.markdown('<div class="sec-title">⚙️ Step 2 — Session Details Configure Karo</div>',
                unsafe_allow_html=True)

    left_cfg, right_cfg = st.columns([1, 1])

    # ── LEFT: Faculty, TLO, Units ─────────────────────────────────────
    with left_cfg:
        st.markdown("#### 🧑‍🏫 Faculty & TLO")
        faculty_id = st.text_input(
            "Faculty Registration ID *",
            placeholder="e.g. IILMGG006412025",
            help="Yeh ID session sheet ke har row mein jayegi",
        )

        tlo_col1, tlo_col2 = st.columns(2)
        with tlo_col1:
            tlo_start = st.number_input("TLO Range — Start", min_value=1, max_value=100, value=1)
        with tlo_col2:
            tlo_end = st.number_input("TLO Range — End", min_value=1, max_value=100, value=5)
        if tlo_end < tlo_start:
            st.markdown('<div class="warn-box">⚠️ TLO End, Start se kam nahi hona chahiye</div>',
                        unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("#### 📚 Unit / Module Configuration")
        num_units = st.selectbox(
            "Kitne Units hain?",
            options=list(range(1, 11)),
            index=4,
            format_func=lambda x: f"{x} Unit{'s' if x > 1 else ''}",
        )

        units_config = []
        total_sessions_configured = 0
        for i in range(num_units):
            uc1, uc2 = st.columns([2, 1])
            with uc1:
                uname = st.text_input(
                    f"Unit {i+1} Name",
                    key=f"uname_{i}",
                    placeholder=f"e.g. UNIT {i+1}: Topic Name",
                )
            with uc2:
                usess = st.number_input(
                    "Sessions",
                    min_value=1, max_value=50,
                    value=max(1, len(dates) // num_units) if dates else 5,
                    key=f"usess_{i}",
                )
            units_config.append({"name": uname, "sessions": usess})
            total_sessions_configured += usess

        if dates:
            diff = total_sessions_configured - len(dates)
            if diff > 0:
                st.markdown(
                    f'<div class="warn-box">⚠️ Total configured sessions ({total_sessions_configured}) '
                    f'> Attendance dates ({len(dates)}). Last {diff} sessions won\'t have dates.</div>',
                    unsafe_allow_html=True)
            elif diff < 0:
                st.markdown(
                    f'<div class="warn-box">⚠️ Attendance mein {len(dates)} dates hain, '
                    f'aapne sirf {total_sessions_configured} sessions configure kiye. '
                    f'Baaki {abs(diff)} dates "Extra Session" honge.</div>',
                    unsafe_allow_html=True)
            else:
                st.markdown(
                    f'<div class="success-box">✅ Sessions ({total_sessions_configured}) = '
                    f'Dates ({len(dates)}) — Perfect match!</div>',
                    unsafe_allow_html=True)

    # ── RIGHT: Day-wise Timing + Date Preview ─────────────────────────
    with right_cfg:
        st.markdown("#### ⏰ Lecture Timing (Din ke Hisab se)")
        st.markdown(
            '<div class="info-box">Har din ka timing alag ho sakta hai — '
            'e.g. Monday 10:00-11:00, Friday 11:10-12:10</div>',
            unsafe_allow_html=True)

        # Detect which days exist in attendance
        present_days = sorted(set(d.strftime("%A") for d in dates if d),
                              key=lambda x: DAY_MAP.get(x, 7))

        day_timings: dict = {}
        all_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        days_to_show = present_days if present_days else all_days

        for day in days_to_show:
            timing_val = st.text_input(
                f"⏱️ {day}",
                value="10:00 - 11:00",
                key=f"timing_{day}",
                placeholder="HH:MM - HH:MM",
            )
            day_timings[day] = timing_val

        # ── Date preview table ────────────────────────────────────────
        if dates:
            st.markdown("#### 📅 Dates from Attendance Sheet")
            preview_rows = []
            for d in dates[:15]:
                if d is None:
                    continue
                day_name = d.strftime("%A")
                raw_t = day_timings.get(day_name, "10:00 - 11:00")
                times = re.findall(r"\d{1,2}:\d{2}", raw_t)
                preview_rows.append({
                    "Date": d.strftime("%d %b %Y"),
                    "Day": day_name,
                    "Timing": raw_t if times else "⚠️ Invalid",
                })
            preview_df = pd.DataFrame(preview_rows)

            def style_day(row):
                colors = {
                    "Monday": "#fff8e1", "Tuesday": "#e8f4fd",
                    "Wednesday": "#f0fff0", "Thursday": "#fdf0f8",
                    "Friday": "#f0f0ff", "Saturday": "#fdf5e6",
                }
                bg = colors.get(row["Day"], "white")
                return [f"background-color:{bg}"] * len(row)

            st.dataframe(
                preview_df.style.apply(style_day, axis=1),
                use_container_width=True,
                hide_index=True,
                height=350,
            )
            if len(dates) > 15:
                st.caption(f"Showing first 15 of {len(dates)} dates")

    st.markdown("---")

    # ══════════════════════════════════════════════════════════════════
    # ROW 3: BUILD SESSION TABLE BUTTON
    # ══════════════════════════════════════════════════════════════════
    st.markdown('<div class="sec-title">📝 Step 3 — Session Table Dekho & Edit Karo</div>',
                unsafe_allow_html=True)

    build_btn = st.button(
        "🔄 Build / Refresh Session Table",
        type="secondary",
        disabled=not bool(dates),
        use_container_width=True,
    )

    if build_btn and dates:
        df = build_session_table(
            dates=dates,
            units=units_config,
            day_timings=day_timings,
            tlo_start=tlo_start,
            tlo_end=tlo_end,
        )
        st.session_state.session_df = df
        st.session_state.results = None

    # ── Editable Table ────────────────────────────────────────────────
    if st.session_state.session_df is not None:
        st.markdown(
            '<div class="info-box">✏️ Neeche table mein directly edit kar sakte ho — '
            'Module Name, Session/Topic, Timing, TLO kuch bhi badlo.</div>',
            unsafe_allow_html=True)

        edited_df = st.data_editor(
            st.session_state.session_df,
            use_container_width=True,
            num_rows="fixed",
            hide_index=True,
            column_config={
                "Sr":             st.column_config.NumberColumn("Sr", width="small", disabled=True),
                "Date":           st.column_config.TextColumn("Date", width="medium", disabled=True),
                "Day":            st.column_config.TextColumn("Day", width="small", disabled=True),
                "Module Name":    st.column_config.TextColumn("Module Name ✏️", width="large"),
                "Session / Topic":st.column_config.TextColumn("Session / Topic ✏️", width="large"),
                "Start Time":     st.column_config.TextColumn("Start Time ✏️", width="medium"),
                "End Time":       st.column_config.TextColumn("End Time ✏️", width="medium"),
                "TLO":            st.column_config.TextColumn("TLO ✏️", width="small"),
                "Mandatory":      st.column_config.SelectboxColumn(
                    "Mandatory ✏️", options=["TRUE", "FALSE"], width="small"),
            },
        )
        st.session_state.session_df = edited_df

    st.markdown("---")

    # ══════════════════════════════════════════════════════════════════
    # ROW 4: GENERATE
    # ══════════════════════════════════════════════════════════════════
    st.markdown('<div class="sec-title">🚀 Step 4 — Files Generate Karo</div>',
                unsafe_allow_html=True)

    can_generate = bool(
        st.session_state.session_df is not None
        and ses_tpl is not None
        and att_tpl is not None
        and att_data is not None
        and faculty_id.strip()
    )
    if not can_generate:
        missing = []
        if not dates:                              missing.append("Attendance Sheet")
        if not ses_tpl:                            missing.append("Session Template")
        if not att_tpl:                            missing.append("Attendance Template")
        if not faculty_id.strip():                 missing.append("Faculty Registration ID")
        if st.session_state.session_df is None:    missing.append("Session Table (Step 3 mein button dabao)")
        st.markdown(
            f'<div class="info-box">ℹ️ Baaki cheezein fill karo: <b>{", ".join(missing)}</b></div>',
            unsafe_allow_html=True)

    gen_btn = st.button("⚡ Generate Session Sheet + All Attendance Files",
                        type="primary", disabled=not can_generate,
                        use_container_width=True)

    if gen_btn and can_generate:
        errors: list = []
        progress = st.progress(0)
        log = st.empty()

        try:
            # Session Sheet
            log.markdown('<div class="info-box">📋 Session sheet bana raha hoon…</div>',
                         unsafe_allow_html=True)
            session_xlsx = generate_session_sheet(
                template_bytes=ses_tpl.read(),
                session_df=st.session_state.session_df,
                faculty_reg_id=faculty_id,
            )
            st.markdown(f'<div class="success-box">✅ Session sheet: '
                        f'<b>{len(st.session_state.session_df)} rows</b> written.</div>',
                        unsafe_allow_html=True)
            progress.progress(40)

            # Day-wise Attendance
            log.markdown('<div class="info-box">🗂️ Day-wise attendance files bana raha hoon…</div>',
                         unsafe_allow_html=True)
            att_tpl_bytes = att_tpl.read()
            daywise: dict = {}
            bar = st.progress(0)

            for i, d in enumerate(dates):
                if d is None:
                    bar.progress((i + 1) / len(dates))
                    continue
                try:
                    fname = f"attendance_{d.strftime('%Y-%m-%d')}.xlsx"
                    daywise[fname] = generate_daywise_attendance(att_tpl_bytes, att_data, d)
                except Exception as ex:
                    errors.append(f"{d}: {ex}")
                bar.progress((i + 1) / len(dates))

            st.markdown(f'<div class="success-box">✅ Day-wise: <b>{len(daywise)} files</b> ready.</div>',
                        unsafe_allow_html=True)
            progress.progress(80)

            # ZIP
            log.markdown('<div class="info-box">📦 ZIP pack kar raha hoon…</div>',
                         unsafe_allow_html=True)
            zip_bytes = build_zip(session_xlsx, daywise)
            progress.progress(100)
            log.markdown('<div class="success-box">🎉 Sab ho gaya! Neeche se download karo.</div>',
                         unsafe_allow_html=True)

            st.session_state.results = {
                "session": session_xlsx,
                "daywise": daywise,
                "zip": zip_bytes,
                "errors": errors,
            }

        except Exception as ex:
            st.markdown(f'<div class="error-box">❌ Error: {ex}</div>', unsafe_allow_html=True)
            import traceback
            st.code(traceback.format_exc())

    # ══════════════════════════════════════════════════════════════════
    # ROW 5: DOWNLOADS
    # ══════════════════════════════════════════════════════════════════
    if st.session_state.results:
        res = st.session_state.results
        st.markdown("---")
        st.markdown('<div class="sec-title">📥 Step 5 — Download Karo</div>',
                    unsafe_allow_html=True)

        for e in res.get("errors", []):
            st.markdown(f'<div class="warn-box">⚠️ Skipped — {e}</div>', unsafe_allow_html=True)

        dc1, dc2 = st.columns(2)
        with dc1:
            st.download_button(
                "📦 ⬇ Download ALL as ZIP",
                data=res["zip"],
                file_name=f"DG_Output_{datetime.now():%Y%m%d_%H%M}.zip",
                mime="application/zip",
                use_container_width=True,
            )
        with dc2:
            st.download_button(
                "📋 ⬇ Download session_sheet.xlsx",
                data=res["session"],
                file_name="session_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.markdown("#### 📅 Day-wise Attendance Files")
        files = list(res["daywise"].items())
        N = 4
        for row_i in range(0, len(files), N):
            chunk = files[row_i: row_i + N]
            cols = st.columns(N)
            for ci, (fname, fb) in enumerate(chunk):
                with cols[ci]:
                    dp = fname.replace("attendance_", "").replace(".xlsx", "")
                    try:
                        label = datetime.strptime(dp, "%Y-%m-%d").strftime("%d %b %Y")
                    except Exception:
                        label = dp
                    st.download_button(
                        f"📅 {label}",
                        data=fb, file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"dl_{fname}",
                    )

        st.markdown("---")
        if st.button("🔄 Reset"):
            st.session_state.results = None
            st.session_state.session_df = None
            st.session_state.att_data = None
            st.rerun()

    st.markdown(
        "<br><center style='color:#aaa;font-size:.75rem;'>"
        "🔒 Sab kuch in-memory process hota hai — koi data server pe store nahi hota</center>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
