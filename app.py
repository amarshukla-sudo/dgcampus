"""
DG Sheet Generator v4.0
========================
• Exact DL.xlsx format: Module Name | Start DateTime | End DateTime |
  Title | Description | Attendance Mandatory | TLO | Faculty Reg ID
• User configures modules, TLO range, titles from syllabus
• Date-specific timing (not day-wise)
• Day-wise attendance fixed
"""

import io, re, zipfile, warnings
from copy import copy
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="DG Sheet Generator", page_icon="📋",
                   layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
.main{padding:.5rem 1.2rem}
.ttl{background:linear-gradient(135deg,#1a2f5a,#2471a3);color:#fff;
     padding:1.2rem 2rem;border-radius:12px;text-align:center;margin-bottom:1rem}
.ttl h1{margin:0;font-size:1.65rem;font-weight:800}
.ttl p{margin:.2rem 0 0;opacity:.88;font-size:.88rem}
.sh{font-size:.95rem;font-weight:700;color:#1a2f5a;border-bottom:2px solid #2471a3;
    padding-bottom:.3rem;margin:1rem 0 .6rem}
.ib{background:#eaf4fb;border-left:4px solid #2471a3;padding:.5rem .9rem;
    border-radius:0 6px 6px 0;margin:.3rem 0;font-size:.83rem}
.sb{background:#d4edda;border-left:4px solid #28a745;padding:.5rem .9rem;
    border-radius:0 6px 6px 0;margin:.3rem 0;font-size:.83rem}
.wb{background:#fff8e1;border-left:4px solid #f39c12;padding:.5rem .9rem;
    border-radius:0 6px 6px 0;margin:.3rem 0;font-size:.83rem}
.eb{background:#fde8e8;border-left:4px solid #e74c3c;padding:.5rem .9rem;
    border-radius:0 6px 6px 0;margin:.3rem 0;font-size:.83rem}
.bk{background:#d4edda;color:#155724;padding:.1rem .5rem;border-radius:9px;
    font-size:.73rem;font-weight:700}
.stDownloadButton>button{background:#1a2f5a !important;color:#fff !important;
    border-radius:7px !important;font-weight:600 !important;width:100% !important}
.stDownloadButton>button:hover{background:#2471a3 !important}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────────────────────────────

def ibox(msg): st.markdown(f'<div class="ib">ℹ️ {msg}</div>', unsafe_allow_html=True)
def sbox(msg): st.markdown(f'<div class="sb">✅ {msg}</div>', unsafe_allow_html=True)
def wbox(msg): st.markdown(f'<div class="wb">⚠️ {msg}</div>', unsafe_allow_html=True)
def ebox(msg): st.markdown(f'<div class="eb">❌ {msg}</div>', unsafe_allow_html=True)
def sec(title): st.markdown(f'<div class="sh">{title}</div>', unsafe_allow_html=True)


def parse_date(raw) -> date | None:
    if raw is None: return None
    try:
        if pd.isna(raw): return None
    except: pass
    if isinstance(raw, datetime): return raw.date()
    if isinstance(raw, date):    return raw
    s = str(raw).strip()
    if s.lower() in ("nat","nan","none",""): return None
    s = re.sub(r"\([^)]*\)","",s)
    s = re.sub(r"\b(sub|Sub|SUB)\b","",s)
    s = re.sub(r"\s+"," ",s).strip()
    if not s: return None
    for fmt in ("%d %b %Y","%d %B %Y","%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%d %b %y"):
        try: return datetime.strptime(s,fmt).date()
        except: pass
    try:
        r = pd.to_datetime(s, dayfirst=True)
        return None if pd.isna(r) else r.date()
    except: return None


def best_sheet(fb: bytes, hints=None) -> pd.DataFrame:
    xf = pd.ExcelFile(io.BytesIO(fb))
    ns = xf.sheet_names
    if len(ns)==1:
        return pd.read_excel(io.BytesIO(fb), sheet_name=ns[0], header=None)
    if hints:
        for kw in hints:
            for s in ns:
                if kw.lower() in s.lower():
                    return pd.read_excel(io.BytesIO(fb), sheet_name=s, header=None)
    best,most = ns[0],-1
    for s in ns:
        df = pd.read_excel(io.BytesIO(fb), sheet_name=s, header=None)
        n  = int(df.notna().sum().sum())
        if n>most: most,best = n,s
    return pd.read_excel(io.BytesIO(fb), sheet_name=best, header=None)


def copy_style(src, dst):
    if src.has_style:
        dst.font       = copy(src.font)
        dst.border     = copy(src.border)
        dst.fill       = copy(src.fill)
        dst.alignment  = copy(src.alignment)
        dst.protection = copy(src.protection)
        dst.number_format = src.number_format


# ─────────────────────────────────────────────────────────────────────
# PARSE ATTENDANCE SHEET
# ─────────────────────────────────────────────────────────────────────

def parse_attendance(fb: bytes) -> dict:
    """Returns dict: dates, students, meta"""
    df_raw = best_sheet(fb, hints=["attendance","student"])

    # Find date row
    date_row_idx = date_col_start = None
    for i, row in df_raw.iterrows():
        hits, first = [], None
        for j,v in enumerate(row):
            d = parse_date(v)
            if d:
                hits.append(d)
                first = first if first is not None else j
        if len(hits) >= 2:
            date_row_idx = i; date_col_start = first; break

    if date_row_idx is None:
        raise ValueError("Attendance sheet mein dates nahi mili.")

    date_row = df_raw.iloc[date_row_idx]
    col_to_date, dates_ordered, seen = {}, [], set()
    for j in range(date_col_start, df_raw.shape[1]):
        d = parse_date(date_row.iloc[j])
        if d:
            col_to_date[j] = d
            if d not in seen: dates_ordered.append(d); seen.add(d)

    # Header row
    name_col = enroll_col = None
    for i in range(date_row_idx):
        row = df_raw.iloc[i]
        rs = " ".join(str(v).lower() for v in row if pd.notna(v))
        if any(k in rs for k in ["name","enrol","roll","sr."]):
            for j,v in enumerate(row):
                s = str(v).lower().strip()
                if "name" in s and name_col is None:      name_col   = j
                elif "enrol" in s and enroll_col is None: enroll_col = j
            break

    # Meta
    meta = {}
    for i in range(0, date_row_idx):
        row = df_raw.iloc[i]
        for j in range(len(row)-1):
            k,v = str(row.iloc[j]).strip(), str(row.iloc[j+1]).strip()
            if k not in ("nan","") and v not in ("nan",""): meta[k] = v

    # Students
    students = []
    for i in range(date_row_idx+1, df_raw.shape[0]):
        row = df_raw.iloc[i]
        if row.notna().sum() < 3: continue
        name = str(row.iloc[name_col]).strip() if name_col is not None and pd.notna(row.iloc[name_col]) else ""
        if not name or name.lower() in ("nan","none",""): continue
        enroll = str(row.iloc[enroll_col]).strip() if enroll_col is not None and pd.notna(row.iloc[enroll_col]) else ""
        att = {}
        for col_j,d in col_to_date.items():
            val = row.iloc[col_j]
            try: att[d] = int(float(val)) if not pd.isna(val) else None
            except: att[d] = None
        students.append({"name":name,"enrollment":enroll,"att":att})

    return {"dates":dates_ordered,"students":students,"meta":meta}


# ─────────────────────────────────────────────────────────────────────
# PARSE SYLLABUS (docx / xlsx / txt)
# ─────────────────────────────────────────────────────────────────────

def parse_syllabus_structured(fb: bytes, filename: str) -> dict:
    """
    Returns {unit_name: [(title, description), ...], ...}
    """
    ext = filename.lower().rsplit(".",1)[-1]
    result = {}

    if ext == "docx":
        import docx as _docx
        doc = _docx.Document(io.BytesIO(fb))
        cur_unit = None
        for p in doc.paragraphs:
            t = p.text.strip()
            if not t: continue
            if re.match(r"^UNIT\s+\d+", t, re.IGNORECASE):
                cur_unit = t; result[cur_unit] = []; continue
            if cur_unit:
                if re.match(r"^Case.?law", t, re.IGNORECASE): continue
                if re.search(r"\bv\.\b|\bAIR\b|\bILR\b|\bSCC\b|https?://|\(\d{4}\)", t): continue
                if len(t) > 5:
                    result[cur_unit].append((t, ""))   # title, desc (empty — user fills)

    elif ext in ("xlsx","xls"):
        df_raw = best_sheet(fb, hints=["syllabus","topic"])
        cur_unit = "Module 1"
        result[cur_unit] = []
        for i, row in df_raw.iterrows():
            for j, val in enumerate(row):
                s = str(val).strip()
                if not s or s.lower() in ("nan","none",""): continue
                if re.match(r"^UNIT\s+\d+|^MODULE\s+\d+|^CHAPTER\s+\d+", s, re.IGNORECASE):
                    cur_unit = s; result[cur_unit] = []; break
                elif len(s) > 5:
                    result[cur_unit].append((s,""))
                    break

    elif ext == "txt":
        cur_unit = "Module 1"; result[cur_unit] = []
        for line in fb.decode("utf-8","ignore").splitlines():
            t = line.strip()
            if not t: continue
            if re.match(r"^UNIT\s+\d+|^MODULE\s+\d+", t, re.IGNORECASE):
                cur_unit = t; result[cur_unit] = []; continue
            if len(t) > 5: result[cur_unit].append((t,""))

    # Remove empty units
    result = {k:v for k,v in result.items() if v}
    return result


# ─────────────────────────────────────────────────────────────────────
# GENERATE SESSION SHEET (exact DL.xlsx format)
# ─────────────────────────────────────────────────────────────────────

HEADER_FILL = PatternFill("solid", start_color="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
DATA_FILL_A = PatternFill("solid", start_color="D6EAF8")
DATA_FILL_B = PatternFill("solid", start_color="EBF5FB")
THIN = Side(style="thin", color="AAAAAA")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

HEADERS = [
    "Module Name*", "Start Date Time*", "End Date Time*",
    "Title*", "Description*", "Attendance Mandatory*",
    "TLO", "Teaching Faculty Registration ID*"
]
COL_WIDTHS = [28, 22, 22, 40, 55, 22, 18, 30]


def generate_session_sheet(rows: list[dict]) -> bytes:
    """
    rows: list of dicts with keys:
      module, start_dt, end_dt, title, description, mandatory, tlo, faculty_id
    """
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Session Sheet"

    # Header row
    for c, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS), 1):
        cell = ws.cell(1, c, h)
        cell.fill   = HEADER_FILL
        cell.font   = HEADER_FONT
        cell.border = BORDER
        cell.alignment = CENTER
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 22

    # Data rows
    for i, row in enumerate(rows, 2):
        fill = DATA_FILL_A if i % 2 == 0 else DATA_FILL_B
        vals = [
            row.get("module", ""),
            row.get("start_dt", ""),
            row.get("end_dt", ""),
            row.get("title", ""),
            row.get("description", ""),
            row.get("mandatory", "TRUE"),
            row.get("tlo", "TLO1"),
            row.get("faculty_id", ""),
        ]
        aligns = [WRAP, CENTER, CENTER, WRAP, WRAP, CENTER, CENTER, CENTER]
        for c, (v, al) in enumerate(zip(vals, aligns), 1):
            cell = ws.cell(i, c, v)
            cell.fill      = fill
            cell.border    = BORDER
            cell.alignment = al
            cell.font      = Font(size=10)
        ws.row_dimensions[i].height = 18

    # Freeze header
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────
# GENERATE DAY-WISE ATTENDANCE  (fixed)
# ─────────────────────────────────────────────────────────────────────

def generate_daywise_attendance(att_tpl_bytes: bytes,
                                 att_data: dict,
                                 session_date: date) -> bytes:
    """
    Uses DG Attendance Template (Email | RegID | Attendance*)
    Matches template RegID with master enrollment → marks PRESENT/ABSENT
    """
    # Load template with all students
    wb = load_workbook(io.BytesIO(att_tpl_bytes))
    ws = wb.active

    # Find column positions from header row 1
    email_col = regid_col = att_col = None
    for c in range(1, ws.max_column + 1):
        hval = str(ws.cell(1, c).value or "").lower().strip()
        if "email" in hval:           email_col = c
        elif "registration" in hval or ("reg" in hval and "id" in hval): regid_col = c
        elif "attendance" in hval:    att_col   = c

    if not all([email_col, regid_col, att_col]):
        # fallback: cols 1,2,3
        email_col, regid_col, att_col = 1, 2, 3

    # Build enrollment → 0/1 map from master attendance
    enroll_att: dict[str, int] = {}
    for st in att_data["students"]:
        v = st["att"].get(session_date)
        enroll_att[str(st["enrollment"]).strip()] = (1 if v == 1 else 0)

    # Update each student row in template
    GREEN_F = PatternFill("solid", start_color="C6EFCE")
    RED_F   = PatternFill("solid", start_color="FFC7CE")
    GREEN_T = Font(color="006100", bold=True, size=10)
    RED_T   = Font(color="9C0006", bold=True, size=10)
    CENTER2 = Alignment(horizontal="center", vertical="center")

    for r in range(2, ws.max_row + 1):
        reg_val = str(ws.cell(r, regid_col).value or "").strip()
        if not reg_val: continue

        # Match: template reg_id ↔ master enrollment
        present = enroll_att.get(reg_val, 0)
        status  = "PRESENT" if present == 1 else "ABSENT"

        cell = ws.cell(r, att_col)
        cell.value     = status
        cell.fill      = GREEN_F if status == "PRESENT" else RED_F
        cell.font      = GREEN_T if status == "PRESENT" else RED_T
        cell.alignment = CENTER2

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────
# BUILD ZIP
# ─────────────────────────────────────────────────────────────────────

def build_zip(session_bytes: bytes, daywise: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("session_sheet.xlsx", session_bytes)
        for fname, fb in daywise.items():
            zf.writestr(f"attendance/{fname}", fb)
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────
# ════════════════════  MAIN UI  ══════════════════════════════════════
# ─────────────────────────────────────────────────────────────────────

def main():
    st.markdown("""
    <div class="ttl">
      <h1>📋 DG Session & Attendance Sheet Generator</h1>
      <p>Exact DL.xlsx format · Date-wise timing · Syllabus se titles · Day-wise attendance ✅</p>
    </div>""", unsafe_allow_html=True)

    # Session state init
    for k in ["att_data","syl_data","date_timing_df","session_rows","results","_att_name"]:
        if k not in st.session_state: st.session_state[k] = None

    # ═════════════════════════════════════════════════════════════════
    # STEP 1 — FILE UPLOADS
    # ═════════════════════════════════════════════════════════════════
    sec("📂 Step 1 — Files Upload Karo")
    u1,u2,u3,u4 = st.columns(4)

    with u1:
        att_file = st.file_uploader("📅 Master Attendance Sheet",  type=["xlsx"], key="att")
        if att_file: st.markdown(f'<span class="bk">✅ {att_file.name}</span>', unsafe_allow_html=True)

    with u2:
        syl_file = st.file_uploader("📚 Syllabus File (optional)", type=["xlsx","docx","txt"], key="syl")
        if syl_file: st.markdown(f'<span class="bk">✅ {syl_file.name}</span>', unsafe_allow_html=True)

    with u3:
        att_tpl = st.file_uploader("🗂️ DG Attendance Template",   type=["xlsx"], key="atpl")
        if att_tpl: st.markdown(f'<span class="bk">✅ {att_tpl.name}</span>', unsafe_allow_html=True)

    with u4:
        st.markdown("**ℹ️ Session sheet output**")
        st.caption("Session sheet is generated fresh in DL.xlsx exact format — no template needed.")

    # Auto-parse attendance
    if att_file and att_file.name != st.session_state._att_name:
        try:
            st.session_state.att_data    = parse_attendance(att_file.read())
            st.session_state._att_name   = att_file.name
            st.session_state.date_timing_df = None
            st.session_state.session_rows   = None
        except Exception as ex:
            ebox(f"Attendance parse error: {ex}")

    # Auto-parse syllabus
    if syl_file:
        try:
            st.session_state.syl_data = parse_syllabus_structured(syl_file.read(), syl_file.name)
        except Exception as ex:
            wbox(f"Syllabus parse warning: {ex}")

    att_data = st.session_state.att_data
    syl_data = st.session_state.syl_data or {}
    dates    = att_data["dates"] if att_data else []

    if dates:
        sbox(f"Attendance loaded: <b>{len(dates)} dates</b> · <b>{len(att_data['students'])} students</b>")

    st.divider()

    # ═════════════════════════════════════════════════════════════════
    # STEP 2 — BASIC SETTINGS
    # ═════════════════════════════════════════════════════════════════
    sec("⚙️ Step 2 — Basic Settings")
    b1,b2,b3 = st.columns(3)
    with b1:
        faculty_id = st.text_input("🧑‍🏫 Faculty Registration ID *",
                                   placeholder="e.g. IILMGG006412025")
    with b2:
        tlo_max = st.number_input("📊 TLO Max Number (e.g. 5 means TLO1–TLO5)",
                                  min_value=1, max_value=100, value=5)
    with b3:
        mandatory = st.selectbox("✅ Attendance Mandatory?", ["TRUE","FALSE"], index=0)

    all_tlos = [f"TLO{i}" for i in range(1, tlo_max+1)]

    st.divider()

    # ═════════════════════════════════════════════════════════════════
    # STEP 3 — MODULE CONFIGURATION
    # ═════════════════════════════════════════════════════════════════
    sec("📚 Step 3 — Modules Configure Karo")

    syl_unit_names = list(syl_data.keys()) if syl_data else []
    if syl_unit_names:
        ibox(f"Syllabus se <b>{len(syl_unit_names)} units</b> mili: {', '.join(syl_unit_names[:3])}{'...' if len(syl_unit_names)>3 else ''}")

    num_modules = st.selectbox("Kitne Modules/Units hain?",
                               list(range(1,11)), index=4,
                               format_func=lambda x: f"{x} Module{'s' if x>1 else ''}")

    modules_cfg = []   # [{name, tlos, num_titles, titles_list}]

    for i in range(num_modules):
        with st.expander(f"📦 Module {i+1} Configuration", expanded=(i==0)):
            mc1,mc2 = st.columns([2,1])
            with mc1:
                # Suggest from syllabus if available
                default_name = syl_unit_names[i] if i < len(syl_unit_names) else f"Module {i+1}"
                mname = st.text_input(f"Module Name *", value=default_name, key=f"mn_{i}")
            with mc2:
                mtlos = st.multiselect(f"TLOs for this module",
                                       options=all_tlos,
                                       default=[all_tlos[i % len(all_tlos)]] if all_tlos else [],
                                       key=f"mt_{i}")
            tlo_str = " | ".join(mtlos) if mtlos else "TLO1"

            # Titles & Descriptions
            # Pull from syllabus if unit name matches
            syl_titles = []
            for syl_unit, syl_pairs in syl_data.items():
                if (mname.strip().lower() in syl_unit.lower() or
                        syl_unit.lower() in mname.strip().lower()):
                    syl_titles = syl_pairs
                    break
            if not syl_titles and i < len(syl_unit_names):
                syl_titles = syl_data.get(syl_unit_names[i], [])

            n_titles = st.number_input(
                f"Kitne Titles/Sessions is module mein?",
                min_value=1, max_value=60,
                value=len(syl_titles) if syl_titles else max(1, len(dates)//num_modules),
                key=f"nt_{i}")

            st.markdown(f"**✏️ Titles & Descriptions (Module {i+1})**")
            st.caption("Title = short topic name | Description = detail — dono edit kar sakte ho")

            # Build default rows
            title_rows = []
            for j in range(n_titles):
                if j < len(syl_titles):
                    t_val, d_val = syl_titles[j]
                else:
                    t_val, d_val = f"Session {j+1} – {mname[:30]}", ""
                title_rows.append({"Title": t_val, "Description": d_val})

            title_df_default = pd.DataFrame(title_rows)
            edited_titles = st.data_editor(
                title_df_default,
                use_container_width=True,
                num_rows="dynamic",
                hide_index=True,
                key=f"te_{i}",
                column_config={
                    "Title":       st.column_config.TextColumn("Title* ✏️",       width="large"),
                    "Description": st.column_config.TextColumn("Description ✏️",  width="large"),
                },
            )

            modules_cfg.append({
                "name":    mname,
                "tlos":    tlo_str,
                "titles":  edited_titles.to_dict("records"),
            })

    total_sessions = sum(len(m["titles"]) for m in modules_cfg)
    if dates:
        diff = total_sessions - len(dates)
        if diff == 0:
            sbox(f"Total sessions ({total_sessions}) = Dates ({len(dates)}) — Perfect match!")
        elif diff > 0:
            wbox(f"Sessions ({total_sessions}) > Dates ({len(dates)}). Last {diff} sessions won't have dates.")
        else:
            wbox(f"Dates ({len(dates)}) > Sessions ({total_sessions}). Last {abs(diff)} dates → 'Extra Session'.")

    st.divider()

    # ═════════════════════════════════════════════════════════════════
    # STEP 4 — DATE-WISE TIMING
    # ═════════════════════════════════════════════════════════════════
    sec("⏰ Step 4 — Date-wise Lecture Timing")

    if not dates:
        ibox("Pehle attendance sheet upload karo — dates load hongi yahan.")
    else:
        ibox(f"<b>{len(dates)} dates</b> mili hain attendance se. Har date ke liye start aur end time bharo.")

        # Build/restore timing dataframe
        if st.session_state.date_timing_df is None or \
           len(st.session_state.date_timing_df) != len(dates):
            timing_rows = []
            for d in dates:
                if d is None: continue
                timing_rows.append({
                    "Date":       d.strftime("%Y-%m-%d"),
                    "Day":        d.strftime("%A"),
                    "Start Time": "10:00",
                    "End Time":   "11:00",
                })
            st.session_state.date_timing_df = pd.DataFrame(timing_rows)

        tc1, tc2 = st.columns([3,1])
        with tc2:
            st.markdown("**⚡ Bulk Set Timing**")
            bulk_day   = st.selectbox("Din select karo",
                                      ["All","Monday","Tuesday","Wednesday",
                                       "Thursday","Friday","Saturday","Sunday"],
                                      key="bday")
            bulk_start = st.text_input("Start Time", value="10:00", key="bst")
            bulk_end   = st.text_input("End Time",   value="11:00", key="bet")
            if st.button("Apply Bulk Timing", use_container_width=True):
                df_t = st.session_state.date_timing_df.copy()
                mask = (df_t["Day"] == bulk_day) if bulk_day != "All" else pd.Series([True]*len(df_t))
                df_t.loc[mask, "Start Time"] = bulk_start
                df_t.loc[mask, "End Time"]   = bulk_end
                st.session_state.date_timing_df = df_t
                st.rerun()

        with tc1:
            edited_timing = st.data_editor(
                st.session_state.date_timing_df,
                use_container_width=True,
                num_rows="fixed",
                hide_index=True,
                key="timing_editor",
                column_config={
                    "Date":       st.column_config.TextColumn("Date",       width="medium", disabled=True),
                    "Day":        st.column_config.TextColumn("Day",        width="small",  disabled=True),
                    "Start Time": st.column_config.TextColumn("Start Time ✏️ (HH:MM)", width="medium"),
                    "End Time":   st.column_config.TextColumn("End Time ✏️ (HH:MM)",   width="medium"),
                },
            )
            st.session_state.date_timing_df = edited_timing

    st.divider()

    # ═════════════════════════════════════════════════════════════════
    # STEP 5 — GENERATE
    # ═════════════════════════════════════════════════════════════════
    sec("🚀 Step 5 — Generate Karo")

    can_gen = bool(
        att_data and dates and
        faculty_id.strip() and
        modules_cfg and
        st.session_state.date_timing_df is not None
    )
    if not can_gen:
        missing = []
        if not att_data:           missing.append("Attendance Sheet")
        if not faculty_id.strip(): missing.append("Faculty Registration ID")
        if missing: ibox(f"Fill karo pehle: <b>{', '.join(missing)}</b>")

    gen_btn = st.button("⚡ Generate Session Sheet + Attendance Files",
                        type="primary", disabled=not can_gen,
                        use_container_width=True)

    if gen_btn and can_gen:
        errors = []
        prog   = st.progress(0)
        log    = st.empty()

        try:
            # ── Build session rows ────────────────────────────────────
            log.markdown('<div class="ib">📋 Session rows bana raha hoon…</div>',
                         unsafe_allow_html=True)

            # Timing lookup: date_str → (start, end)
            timing_df = st.session_state.date_timing_df
            timing_map: dict[str, tuple] = {}
            for _, tr in timing_df.iterrows():
                timing_map[tr["Date"]] = (str(tr["Start Time"]).strip(),
                                          str(tr["End Time"]).strip())

            # Flatten all sessions from modules
            all_sessions = []
            for m in modules_cfg:
                for t in m["titles"]:
                    all_sessions.append({
                        "module":  m["name"],
                        "tlo":     m["tlos"],
                        "title":   t.get("Title",""),
                        "desc":    t.get("Description",""),
                    })

            session_rows = []
            valid_dates  = [d for d in dates if d is not None]

            for idx, d in enumerate(valid_dates):
                date_str = d.strftime("%Y-%m-%d")
                start_t, end_t = timing_map.get(date_str, ("10:00","11:00"))

                # Build datetime strings (DL.xlsx format)
                start_dt = f"{date_str} {start_t}:00"
                end_dt   = f"{date_str} {end_t}:00"

                if idx < len(all_sessions):
                    sess = all_sessions[idx]
                else:
                    sess = {"module":"Extra Session","tlo":"TLO1",
                            "title":"Extra Session","desc":""}

                session_rows.append({
                    "module":     sess["module"],
                    "start_dt":   start_dt,
                    "end_dt":     end_dt,
                    "title":      sess["title"],
                    "description": sess["desc"],
                    "mandatory":  mandatory,
                    "tlo":        sess["tlo"],
                    "faculty_id": faculty_id.strip(),
                })

            prog.progress(25)

            # ── Generate session sheet ────────────────────────────────
            log.markdown('<div class="ib">📊 Session sheet generate ho rahi hai…</div>',
                         unsafe_allow_html=True)
            session_xlsx = generate_session_sheet(session_rows)
            sbox(f"Session sheet: <b>{len(session_rows)} rows</b> written.")
            prog.progress(50)

            # ── Day-wise attendance ───────────────────────────────────
            daywise = {}
            if att_tpl:
                log.markdown('<div class="ib">🗂️ Day-wise attendance files bana raha hoon…</div>',
                             unsafe_allow_html=True)
                att_tpl_bytes = att_tpl.read()
                bar = st.progress(0)
                for i, d in enumerate(valid_dates):
                    try:
                        fname = f"attendance_{d.strftime('%Y-%m-%d')}.xlsx"
                        daywise[fname] = generate_daywise_attendance(att_tpl_bytes, att_data, d)
                    except Exception as ex:
                        errors.append(f"{d}: {ex}")
                    bar.progress((i+1)/len(valid_dates))
                sbox(f"Day-wise attendance: <b>{len(daywise)} files</b> ready.")
            else:
                wbox("DG Attendance Template upload nahi hua — sirf session sheet download hogi.")

            prog.progress(80)

            # ── ZIP ───────────────────────────────────────────────────
            zip_bytes = build_zip(session_xlsx, daywise)
            prog.progress(100)
            log.markdown('<div class="sb">🎉 Sab ho gaya! Neeche se download karo.</div>',
                         unsafe_allow_html=True)

            st.session_state.results = {
                "session":  session_xlsx,
                "daywise":  daywise,
                "zip":      zip_bytes,
                "errors":   errors,
                "n_rows":   len(session_rows),
            }

        except Exception as ex:
            ebox(f"Error: {ex}")
            import traceback; st.code(traceback.format_exc())

    # ═════════════════════════════════════════════════════════════════
    # STEP 6 — DOWNLOADS
    # ═════════════════════════════════════════════════════════════════
    if st.session_state.results:
        res = st.session_state.results
        st.divider()
        sec("📥 Step 6 — Download Karo")

        for e in res.get("errors",[]): wbox(f"Skipped: {e}")

        # Summary
        st.markdown(f"""
| Output | Count | Status |
|--------|-------|--------|
| 📋 session_sheet.xlsx | {res['n_rows']} rows | ✅ Ready |
| 📅 Day-wise attendance | {len(res['daywise'])} files | ✅ Ready |
| 📦 ZIP archive | All in one | ✅ Ready |
""")

        d1,d2 = st.columns(2)
        with d1:
            st.download_button("📦 ⬇ Download ALL as ZIP",
                data=res["zip"],
                file_name=f"DG_Output_{datetime.now():%Y%m%d_%H%M}.zip",
                mime="application/zip",
                use_container_width=True)
        with d2:
            st.download_button("📋 ⬇ Download session_sheet.xlsx",
                data=res["session"],
                file_name="session_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        if res["daywise"]:
            st.markdown("#### 📅 Individual Day-wise Attendance Files")
            files = list(res["daywise"].items())
            for row_i in range(0, len(files), 4):
                chunk = files[row_i:row_i+4]
                cols  = st.columns(4)
                for ci,(fname,fb) in enumerate(chunk):
                    with cols[ci]:
                        dp = fname.replace("attendance_","").replace(".xlsx","")
                        try:    label = datetime.strptime(dp,"%Y-%m-%d").strftime("%d %b %Y")
                        except: label = dp
                        st.download_button(f"📅 {label}",
                            data=fb, file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key=f"dl_{fname}")

        st.divider()
        if st.button("🔄 Reset & Start Over"):
            for k in ["att_data","syl_data","date_timing_df","session_rows","results","_att_name"]:
                st.session_state[k] = None
            st.rerun()

    st.markdown(
        "<center style='color:#aaa;font-size:.72rem;margin-top:1rem'>"
        "🔒 In-memory processing — no data stored on server</center>",
        unsafe_allow_html=True)


if __name__ == "__main__":
    main()
