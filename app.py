"""
DG Sheet Generator v5.0
========================
Flow:
  1. Upload files
  2. Ask: Total sessions to create (e.g. 50) → N rows in output
  3. Modules config: name + TLO assignment
  4. Syllabus auto-extracted + balanced across modules = N titles
  5. Descriptions auto-generated from titles
  6. Days/week config → timing per day (Mon 10:00-11:00, Fri 11:10-12:10)
  7. Dates from attendance → first N dates used
  8. Preview editable table → Generate
"""

import io, re, zipfile, warnings
from copy import copy
from datetime import datetime, date
from math import ceil

import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="DG Sheet Generator", page_icon="📋",
                   layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
.main{padding:.4rem 1.2rem}
.ttl{background:linear-gradient(135deg,#1a2f5a,#2471a3);color:#fff;
     padding:1.1rem 2rem;border-radius:12px;text-align:center;margin-bottom:.8rem}
.ttl h1{margin:0;font-size:1.6rem;font-weight:800}
.ttl p{margin:.2rem 0 0;opacity:.88;font-size:.85rem}
.sh{font-size:.95rem;font-weight:700;color:#1a2f5a;
    border-bottom:2px solid #2471a3;padding-bottom:.3rem;margin:.9rem 0 .6rem}
.ib{background:#eaf4fb;border-left:4px solid #2471a3;padding:.45rem .85rem;
    border-radius:0 6px 6px 0;margin:.3rem 0;font-size:.82rem}
.sb{background:#d4edda;border-left:4px solid #28a745;padding:.45rem .85rem;
    border-radius:0 6px 6px 0;margin:.3rem 0;font-size:.82rem}
.wb{background:#fff8e1;border-left:4px solid #f39c12;padding:.45rem .85rem;
    border-radius:0 6px 6px 0;margin:.3rem 0;font-size:.82rem}
.eb{background:#fde8e8;border-left:4px solid #e74c3c;padding:.45rem .85rem;
    border-radius:0 6px 6px 0;margin:.3rem 0;font-size:.82rem}
.bk{background:#d4edda;color:#155724;padding:.1rem .5rem;
    border-radius:9px;font-size:.73rem;font-weight:700}
.stDownloadButton>button{background:#1a2f5a !important;color:#fff !important;
    border-radius:7px !important;font-weight:600 !important;width:100% !important}
.stDownloadButton>button:hover{background:#2471a3 !important}
</style>
""", unsafe_allow_html=True)

def ib(m): st.markdown(f'<div class="ib">ℹ️ {m}</div>', unsafe_allow_html=True)
def sb(m): st.markdown(f'<div class="sb">✅ {m}</div>', unsafe_allow_html=True)
def wb(m): st.markdown(f'<div class="wb">⚠️ {m}</div>', unsafe_allow_html=True)
def eb(m): st.markdown(f'<div class="eb">❌ {m}</div>', unsafe_allow_html=True)
def sh(t): st.markdown(f'<div class="sh">{t}</div>', unsafe_allow_html=True)

WEEKDAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
DAY_IDX  = {d:i for i,d in enumerate(WEEKDAYS)}

# ─────────────────────────────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────────────────────────────

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
    if len(ns)==1: return pd.read_excel(io.BytesIO(fb), sheet_name=ns[0], header=None)
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


# ─────────────────────────────────────────────────────────────────────
# PARSE ATTENDANCE → DATES + STUDENTS
# ─────────────────────────────────────────────────────────────────────

def parse_attendance(fb: bytes) -> dict:
    df_raw = best_sheet(fb, hints=["attendance","student"])
    date_row_idx = date_col_start = None
    for i, row in df_raw.iterrows():
        hits, first = [], None
        for j,v in enumerate(row):
            d = parse_date(v)
            if d: hits.append(d); first = first if first is not None else j
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

    meta = {}
    for i in range(date_row_idx):
        row = df_raw.iloc[i]
        for j in range(len(row)-1):
            k,v = str(row.iloc[j]).strip(), str(row.iloc[j+1]).strip()
            if k not in ("nan","") and v not in ("nan",""): meta[k] = v

    students = []
    for i in range(date_row_idx+1, df_raw.shape[0]):
        row = df_raw.iloc[i]
        if row.notna().sum() < 3: continue
        name   = str(row.iloc[name_col]).strip()   if name_col   is not None and pd.notna(row.iloc[name_col])   else ""
        enroll = str(row.iloc[enroll_col]).strip() if enroll_col is not None and pd.notna(row.iloc[enroll_col]) else ""
        if not name or name.lower() in ("nan","none",""): continue
        att = {}
        for col_j,d in col_to_date.items():
            val = row.iloc[col_j]
            try:    att[d] = int(float(val)) if not pd.isna(val) else None
            except: att[d] = None
        students.append({"name":name,"enrollment":enroll,"att":att})

    return {"dates":dates_ordered,"students":students,"meta":meta}


# ─────────────────────────────────────────────────────────────────────
# PARSE SYLLABUS → structured {unit: [(title,desc),...]}
# ─────────────────────────────────────────────────────────────────────

def parse_syllabus(fb: bytes, filename: str) -> dict:
    """Returns OrderedDict: {unit_name: [title, ...]}"""
    ext = filename.lower().rsplit(".",1)[-1]
    result = {}

    if ext == "docx":
        import docx as _d
        doc = _d.Document(io.BytesIO(fb))
        cur = None
        for p in doc.paragraphs:
            t = p.text.strip()
            if not t: continue
            if re.match(r"^UNIT\s+\d+", t, re.IGNORECASE):
                cur = t; result[cur] = []; continue
            if cur:
                if re.match(r"^Case.?law", t, re.IGNORECASE): continue
                # Skip legal citations
                if re.search(r"\sv\.\s|\bAIR\b|\bILR\b|\bSCC\b|https?://|\(\d{4}\)\s+\d|\bLL\s+\(", t): continue
                if re.match(r"^[A-Z][a-zA-Z\s]+\s+v\.\s+[A-Z]", t): continue
                if len(t) > 5: result[cur].append(t)

    elif ext in ("xlsx","xls"):
        df_raw = best_sheet(fb)
        cur = "Module 1"; result[cur] = []
        for i, row in df_raw.iterrows():
            for _, val in enumerate(row):
                s = str(val).strip()
                if not s or s.lower() in ("nan","none",""): continue
                if re.match(r"^(UNIT|MODULE|CHAPTER)\s+\d+", s, re.IGNORECASE):
                    cur = s; result[cur] = []; break
                elif len(s) > 5:
                    result[cur].append(s); break

    elif ext == "txt":
        cur = "Module 1"; result[cur] = []
        for line in fb.decode("utf-8","ignore").splitlines():
            t = line.strip()
            if not t: continue
            if re.match(r"^(UNIT|MODULE)\s+\d+", t, re.IGNORECASE):
                cur = t; result[cur] = []; continue
            if len(t) > 5: result[cur].append(t)

    result = {k:v for k,v in result.items() if v}
    return result


# ─────────────────────────────────────────────────────────────────────
# AUTO-BALANCE TITLES across modules
# ─────────────────────────────────────────────────────────────────────

def balance_titles(syl_data: dict, modules: list, total_sessions: int) -> list:
    """
    modules: [{"name":str, "tlo":str, "unit_key":str}, ...]
    Returns: [{"module":str,"tlo":str,"title":str}, ...] length = total_sessions
    """
    # Map each module to its syllabus unit
    mod_titles = []
    for m in modules:
        unit_key = m.get("unit_key","")
        raw_titles = syl_data.get(unit_key, [])
        # Filter out case-law entries (already done in parse but double-check)
        clean = [t for t in raw_titles if len(t) > 5]
        mod_titles.append({"module": m["name"], "tlo": m["tlo"], "titles": clean})

    n_mods = len(mod_titles)
    if n_mods == 0: return []

    # Distribute total_sessions proportionally based on available titles
    total_available = sum(len(m["titles"]) for m in mod_titles)
    alloc = []
    for m in mod_titles:
        if total_available > 0:
            share = round(total_sessions * len(m["titles"]) / total_available)
        else:
            share = total_sessions // n_mods
        alloc.append(max(1, share))

    # Adjust to exactly total_sessions
    while sum(alloc) < total_sessions:
        alloc[alloc.index(min(alloc))] += 1
    while sum(alloc) > total_sessions:
        alloc[alloc.index(max(alloc))] -= 1

    result = []
    for m, n in zip(mod_titles, alloc):
        titles = m["titles"]
        if not titles:
            titles = [f"Session – {m['module']}"]
        # Distribute n titles from this module's list
        for i in range(n):
            if i < len(titles):
                t = titles[i]
            else:
                # Repeat last title with sequence suffix
                t = f"{titles[-1]} (Part {i - len(titles) + 2})"
            result.append({"module": m["module"], "tlo": m["tlo"], "title": t})

    return result[:total_sessions]


# ─────────────────────────────────────────────────────────────────────
# AUTO-GENERATE DESCRIPTION from title
# ─────────────────────────────────────────────────────────────────────

def auto_description(title: str, module_name: str) -> str:
    t = title.strip()
    mod = re.sub(r"^UNIT\s+\d+:\s*","",module_name,flags=re.IGNORECASE).strip()
    tl  = t.lower()

    if any(w in tl for w in ["definition","define","meaning of"]):
        core = re.sub(r"^(definition\s+(of\s+)?|define\s+)", "", t, flags=re.IGNORECASE).strip()
        return (f"This session covers the statutory definition and essential elements of {core}. "
                f"Students will examine the legal provisions, judicial interpretations, and practical "
                f"significance within the framework of {mod}.")

    if any(w in tl for w in ["rights","duties","liability","liabilities","obligation"]):
        return (f"This session examines the rights, duties, and liabilities arising under {t}. "
                f"Students will analyse relevant statutory provisions, landmark judgments, and the "
                f"legal consequences for parties involved in {mod}.")

    if any(w in tl for w in ["distinction","difference","compare","vs","versus","and"]):
        return (f"This session provides a comparative analysis of {t}. "
                f"Students will identify key distinctions through statutory provisions and case-law, "
                f"enabling them to apply differential reasoning in {mod} disputes.")

    if any(w in tl for w in ["type","kind","classif","categor","form"]):
        return (f"This session classifies the different types and categories covered under {t}. "
                f"Students will study the legal significance of each category and their practical "
                f"application within the domain of {mod}.")

    if any(w in tl for w in ["termination","discharge","revocation","dissolution"]):
        return (f"This session covers the modes of {t} and their legal consequences. "
                f"Students will study the statutory provisions, conditions, and judicial precedents "
                f"governing this aspect of {mod}.")

    if any(w in tl for w in ["creation","formation","essential","element","requisite"]):
        return (f"This session discusses the process of {t} and the requisite legal elements. "
                f"Students will examine statutory requirements, judicial interpretations, "
                f"and practical illustrations relevant to {mod}.")

    if any(w in tl for w in ["position","english law","indian law","common law"]):
        return (f"This session presents a comparative study of {t}, contrasting the approach "
                f"under Indian and English legal systems. Students will critically evaluate "
                f"doctrinal developments and their influence on {mod}.")

    if any(w in tl for w in ["remedy","remedies","compensation","damages"]):
        return (f"This session discusses the available remedies and relief in cases involving {t}. "
                f"Students will examine statutory and equitable remedies, judicial approaches, "
                f"and practical computation of relief in {mod} disputes.")

    if any(w in tl for w in ["nature","scope","extent","concept"]):
        return (f"This session explores the nature, scope, and conceptual framework of {t}. "
                f"Students will critically analyse the theoretical underpinnings and legislative "
                f"intent governing this concept within {mod}.")

    if any(w in tl for w in ["case","judgment","ruling","court","v."]):
        return (f"This session analyses {t} as a landmark judicial decision. "
                f"Students will examine the facts, legal issues, reasoning of the court, "
                f"and the precedential value of this ruling in {mod}.")

    # Generic fallback
    return (f"This session provides a comprehensive study of {t} within the domain of {mod}. "
            f"Students will analyse the relevant statutory provisions, judicial precedents, "
            f"and practical applications through structured discussion and case-based learning.")


# ─────────────────────────────────────────────────────────────────────
# GENERATE SESSION SHEET (exact DL.xlsx format)
# ─────────────────────────────────────────────────────────────────────

H_FILL  = PatternFill("solid", start_color="1F4E79")
H_FONT  = Font(bold=True, color="FFFFFF", size=11)
DA_FILL = PatternFill("solid", start_color="D6EAF8")
DB_FILL = PatternFill("solid", start_color="EBF5FB")
THIN    = Side(style="thin", color="BBBBBB")
BRD     = Border(left=THIN,right=THIN,top=THIN,bottom=THIN)
CTR     = Alignment(horizontal="center", vertical="center", wrap_text=True)
LFT     = Alignment(horizontal="left",   vertical="center", wrap_text=True)

HEADERS   = ["Module Name*","Start Date Time*","End Date Time*","Title*",
             "Description*","Attendance Mandatory*","TLO","Teaching Faculty Registration ID*"]
COL_WIDTHS= [30, 22, 22, 42, 58, 22, 18, 32]


def generate_session_sheet(rows: list) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Session Sheet"

    for c,(h,w) in enumerate(zip(HEADERS,COL_WIDTHS),1):
        cell = ws.cell(1,c,h)
        cell.fill = H_FILL; cell.font = H_FONT
        cell.border = BRD;  cell.alignment = CTR
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 24

    aligns = [LFT,CTR,CTR,LFT,LFT,CTR,CTR,CTR]
    for i,row in enumerate(rows,2):
        fill = DA_FILL if i%2==0 else DB_FILL
        vals = [row.get("module",""), row.get("start_dt",""), row.get("end_dt",""),
                row.get("title",""),  row.get("description",""), row.get("mandatory","TRUE"),
                row.get("tlo","TLO1"), row.get("faculty_id","")]
        for c,(v,al) in enumerate(zip(vals,aligns),1):
            cell = ws.cell(i,c,v)
            cell.fill=fill; cell.border=BRD
            cell.alignment=al; cell.font=Font(size=10)
        ws.row_dimensions[i].height = 30

    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────
# GENERATE DAY-WISE ATTENDANCE (fixed)
# ─────────────────────────────────────────────────────────────────────

def generate_daywise_attendance(att_tpl_bytes: bytes, att_data: dict,
                                 session_date: date) -> bytes:
    wb  = load_workbook(io.BytesIO(att_tpl_bytes))
    ws  = wb.active

    email_col=regid_col=att_col=None
    for c in range(1, ws.max_column+1):
        hv = str(ws.cell(1,c).value or "").lower().strip()
        if "email" in hv:                                 email_col = c
        elif "registration" in hv or ("reg" in hv and "id" in hv): regid_col = c
        elif "attendance" in hv:                          att_col   = c
    if not all([email_col,regid_col,att_col]):
        email_col,regid_col,att_col = 1,2,3

    # Build reg_id → attendance map from master sheet
    enroll_att: dict = {}
    for st in att_data["students"]:
        v = st["att"].get(session_date)
        enroll_att[str(st["enrollment"]).strip()] = (1 if v==1 else 0)

    GF = PatternFill("solid", start_color="C6EFCE")
    RF = PatternFill("solid", start_color="FFC7CE")
    GT = Font(color="006100", bold=True, size=10)
    RT = Font(color="9C0006", bold=True, size=10)
    CA = Alignment(horizontal="center", vertical="center")

    for r in range(2, ws.max_row+1):
        reg = str(ws.cell(r,regid_col).value or "").strip()
        if not reg: continue
        present = enroll_att.get(reg, 0)
        status  = "PRESENT" if present==1 else "ABSENT"
        cell = ws.cell(r, att_col)
        cell.value     = status
        cell.fill      = GF if status=="PRESENT" else RF
        cell.font      = GT if status=="PRESENT" else RT
        cell.alignment = CA

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_zip(session_bytes: bytes, daywise: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("session_sheet.xlsx", session_bytes)
        for fname,fb in daywise.items():
            zf.writestr(f"attendance/{fname}", fb)
    buf.seek(0)
    return buf.getvalue()


# ═════════════════════════════════════════════════════════════════════
# ═══════════════════════  MAIN UI  ═══════════════════════════════════
# ═════════════════════════════════════════════════════════════════════

def main():
    st.markdown("""
    <div class="ttl">
      <h1>📋 DG Session & Attendance Sheet Generator</h1>
      <p>Sessions ki number batao → Syllabus auto-balance → Descriptions auto-generate → Download!</p>
    </div>""", unsafe_allow_html=True)

    # session state init
    for k in ["att_data","syl_data","preview_df","results","_att_nm","_syl_nm"]:
        if k not in st.session_state: st.session_state[k] = None

    # ═════════════════════════════════════════════════════════════════
    # STEP 1 — FILES
    # ═════════════════════════════════════════════════════════════════
    sh("📂 Step 1 — Files Upload Karo")
    f1,f2,f3 = st.columns(3)

    with f1:
        att_file = st.file_uploader("📅 Master Attendance Sheet (.xlsx)", type=["xlsx"], key="att")
        if att_file: st.markdown(f'<span class="bk">✅ {att_file.name}</span>', unsafe_allow_html=True)

    with f2:
        syl_file = st.file_uploader("📚 Syllabus File (.docx / .xlsx / .txt)", type=["docx","xlsx","txt"], key="syl")
        if syl_file: st.markdown(f'<span class="bk">✅ {syl_file.name}</span>', unsafe_allow_html=True)

    with f3:
        att_tpl = st.file_uploader("🗂️ DG Attendance Template (.xlsx)", type=["xlsx"], key="atpl")
        if att_tpl: st.markdown(f'<span class="bk">✅ {att_tpl.name}</span>', unsafe_allow_html=True)

    # Auto-parse
    if att_file and att_file.name != st.session_state._att_nm:
        try:
            st.session_state.att_data  = parse_attendance(att_file.read())
            st.session_state._att_nm   = att_file.name
            st.session_state.preview_df = None
            st.session_state.results    = None
        except Exception as ex: eb(f"Attendance error: {ex}")

    if syl_file and syl_file.name != st.session_state._syl_nm:
        try:
            st.session_state.syl_data = parse_syllabus(syl_file.read(), syl_file.name)
            st.session_state._syl_nm  = syl_file.name
        except Exception as ex: wb(f"Syllabus warning: {ex}")

    att_data = st.session_state.att_data
    syl_data = st.session_state.syl_data or {}
    dates    = att_data["dates"] if att_data else []

    if dates:
        sb(f"Attendance: <b>{len(dates)} dates</b> · <b>{len(att_data['students'])} students</b>")
    if syl_data:
        total_syl = sum(len(v) for v in syl_data.values())
        sb(f"Syllabus: <b>{len(syl_data)} units</b> · <b>{total_syl} topics</b> extracted")

    st.divider()

    # ═════════════════════════════════════════════════════════════════
    # STEP 2 — SESSION COUNT + BASIC INFO
    # ═════════════════════════════════════════════════════════════════
    sh("🔢 Step 2 — Kitne Sessions Banane Hain?")

    s1,s2,s3,s4 = st.columns(4)
    with s1:
        total_sessions = st.number_input(
            "📊 Total Sessions (Rows in Sheet)",
            min_value=1, max_value=500,
            value=min(len(dates), 50) if dates else 50,
            help="Yahi number of rows session sheet mein honge"
        )
    with s2:
        faculty_id = st.text_input("🧑‍🏫 Faculty Registration ID *",
                                   placeholder="e.g. IILMGG006412025")
    with s3:
        mandatory = st.selectbox("✅ Attendance Mandatory?", ["TRUE","FALSE"])
    with s4:
        tlo_max = st.number_input("📊 Max TLO Number", min_value=1, max_value=100, value=5,
                                  help="TLO1 se TLO(n) tak generate honge")

    all_tlos = [f"TLO{i}" for i in range(1, tlo_max+1)]

    if dates and total_sessions > len(dates):
        wb(f"Total sessions ({total_sessions}) > Attendance dates ({len(dates)}). "
           f"Sirf {len(dates)} dates available hain — baaki sessions blank dates honge.")

    st.divider()

    # ═════════════════════════════════════════════════════════════════
    # STEP 3 — LECTURE DAYS + TIMING
    # ═════════════════════════════════════════════════════════════════
    sh("⏰ Step 3 — Lecture Days & Timing Configure Karo")

    t_col1, t_col2 = st.columns([1,2])

    with t_col1:
        # Detect which days appear in attendance
        detected_days = []
        if dates:
            day_counts = {}
            for d in dates:
                if d: day_counts[d.strftime("%A")] = day_counts.get(d.strftime("%A"),0)+1
            detected_days = sorted(day_counts.keys(), key=lambda x: DAY_IDX.get(x,7))

        days_per_week = st.number_input("📅 Lectures per week (kitne din)", 
                                        min_value=1, max_value=7, value=len(detected_days) or 3)

        selected_days = st.multiselect(
            "📌 Kaun se din lecture hota hai?",
            options=WEEKDAYS,
            default=detected_days[:days_per_week] if detected_days else WEEKDAYS[:days_per_week],
            max_selections=days_per_week,
        )

    with t_col2:
        if selected_days:
            st.markdown("**⏱️ Har din ka timing (Start Time – End Time)**")
            day_timing: dict = {}
            cols_t = st.columns(min(len(selected_days), 4))
            for i, day in enumerate(selected_days):
                with cols_t[i % 4]:
                    st.markdown(f"**{day[:3]}**")
                    # Detect default timing from dates
                    default_start = "10:00"
                    default_end   = "11:00"
                    if dates:
                        day_dates = [d for d in dates if d and d.strftime("%A")==day]
                        # Use position in week to guess timing (common pattern)
                        idx_in_week = selected_days.index(day)
                        if idx_in_week == 1:
                            default_start, default_end = "12:20","13:20"
                        elif idx_in_week == 2:
                            default_start, default_end = "11:10","12:10"
                    s_t = st.text_input("Start", value=default_start, key=f"st_{day}",
                                        label_visibility="collapsed",
                                        placeholder="HH:MM")
                    e_t = st.text_input("End",   value=default_end,   key=f"et_{day}",
                                        label_visibility="collapsed",
                                        placeholder="HH:MM")
                    day_timing[day] = (s_t.strip(), e_t.strip())

        # Preview: dates with assigned timings
        if dates and selected_days and 'day_timing' in dir():
            valid_dates = [d for d in dates if d]
            preview_n   = min(total_sessions, len(valid_dates))
            use_dates   = valid_dates[:preview_n]

            preview_rows = []
            for d in use_dates[:8]:
                day_name = d.strftime("%A")
                s_t,e_t  = day_timing.get(day_name, ("10:00","11:00"))
                preview_rows.append({
                    "Date": d.strftime("%d %b %Y"),
                    "Day":  day_name,
                    "Start": s_t,
                    "End":   e_t,
                })
            if preview_rows:
                st.markdown("**📅 Date-Timing Preview (first 8):**")
                def style_row(row):
                    colors = {"Monday":"#fff8e1","Tuesday":"#e8f4fd","Wednesday":"#f0fff0",
                              "Thursday":"#fdf0f8","Friday":"#f0f0ff","Saturday":"#fdf5e6"}
                    bg = colors.get(row["Day"],"white")
                    return [f"background-color:{bg}"]*len(row)
                pv_df = pd.DataFrame(preview_rows)
                st.dataframe(pv_df.style.apply(style_row,axis=1),
                             use_container_width=True, hide_index=True, height=210)

    st.divider()

    # ═════════════════════════════════════════════════════════════════
    # STEP 4 — MODULE CONFIGURATION
    # ═════════════════════════════════════════════════════════════════
    sh("📚 Step 4 — Modules & Units Configure Karo")

    syl_units = list(syl_data.keys()) if syl_data else []
    if syl_units:
        ib(f"Syllabus se <b>{len(syl_units)} units</b> mili. Module names auto-filled — edit kar sakte ho.")

    num_modules = st.selectbox(
        "Kitne Modules/Units hain?",
        list(range(1,11)), index=min(len(syl_units)-1,4) if syl_units else 4,
        format_func=lambda x: f"{x} Module{'s' if x>1 else ''}"
    )

    modules_cfg = []
    m_cols = st.columns(min(num_modules,3))
    for i in range(num_modules):
        with m_cols[i % min(num_modules,3)]:
            with st.expander(f"📦 Module {i+1}", expanded=True):
                default_name = syl_units[i] if i < len(syl_units) else f"Module {i+1}"
                mname = st.text_input("Module Name *", value=default_name, key=f"mn_{i}")

                # TLO multiselect
                def_tlo = [all_tlos[i % len(all_tlos)]] if all_tlos else ["TLO1"]
                mtlos   = st.multiselect("TLOs (is module ke liye)", all_tlos,
                                         default=def_tlo, key=f"mt_{i}")
                tlo_str = " | ".join(mtlos) if mtlos else all_tlos[i % len(all_tlos)]

                # Show how many topics available in syllabus for this unit
                unit_key = syl_units[i] if i < len(syl_units) else ""
                avail    = len(syl_data.get(unit_key, []))
                if avail:
                    st.caption(f"📖 {avail} topics syllabus mein")

                modules_cfg.append({"name":mname,"tlo":tlo_str,"unit_key":unit_key})

    st.divider()

    # ═════════════════════════════════════════════════════════════════
    # STEP 5 — BUILD PREVIEW TABLE
    # ═════════════════════════════════════════════════════════════════
    sh("👀 Step 5 — Preview Table Banao & Edit Karo")

    ib("'Build Preview' dabao → auto-balanced titles + auto-generated descriptions dikhenge. "
       "Table mein seedha edit kar sakte ho before generating.")

    can_preview = bool(dates and modules_cfg and faculty_id.strip() if True else False)

    prev_btn = st.button("🔄 Build Preview Table",
                         type="secondary", use_container_width=True)

    if prev_btn:
        if not dates:
            eb("Attendance sheet upload karo pehle.")
        elif not syl_data:
            # No syllabus — create empty titles
            wb("Syllabus nahi diya — titles manually bharne padenge table mein.")
            flat = [{"module":m["name"],"tlo":m["tlo"],"title":f"Session {i+1}"}
                    for m in modules_cfg
                    for i in range(max(1, total_sessions//len(modules_cfg)))]
            flat = flat[:total_sessions]
        else:
            flat = balance_titles(syl_data, modules_cfg, total_sessions)

        # Build full preview df
        valid_dates = [d for d in dates if d]
        day_timing_local = {}
        if 'day_timing' in dir():
            day_timing_local = day_timing
        else:
            for d in selected_days if selected_days else []:
                day_timing_local[d] = ("10:00","11:00")

        rows = []
        for idx in range(total_sessions):
            # Date
            if idx < len(valid_dates):
                d = valid_dates[idx]
                date_str  = d.strftime("%Y-%m-%d")
                day_name  = d.strftime("%A")
                s_t, e_t  = day_timing_local.get(day_name, ("10:00","11:00"))
                start_dt  = f"{date_str} {s_t}:00"
                end_dt    = f"{date_str} {e_t}:00"
            else:
                date_str  = "TBD"
                start_dt  = "TBD"
                end_dt    = "TBD"

            sess = flat[idx] if idx < len(flat) else {
                "module":"Extra Session","tlo":all_tlos[0],"title":"Extra Session"}

            desc = auto_description(sess["title"], sess["module"])

            rows.append({
                "Sr":           idx+1,
                "Module Name*": sess["module"],
                "Start Date Time*": start_dt,
                "End Date Time*":   end_dt,
                "Title*":           sess["title"],
                "Description*":     desc,
                "Mandatory*":       mandatory,
                "TLO":              sess["tlo"],
                "Faculty Reg ID*":  faculty_id.strip(),
            })

        st.session_state.preview_df = pd.DataFrame(rows)
        st.session_state.results    = None

    # Show editable table
    if st.session_state.preview_df is not None:
        df = st.session_state.preview_df
        sb(f"Preview ready: <b>{len(df)} rows</b> · Edit karke changes save honge automatically.")

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="fixed",
            hide_index=True,
            key="preview_editor",
            column_config={
                "Sr":                st.column_config.NumberColumn("Sr", width=50, disabled=True),
                "Module Name*":      st.column_config.TextColumn("Module Name* ✏️",      width=220),
                "Start Date Time*":  st.column_config.TextColumn("Start DateTime* ✏️",   width=160),
                "End Date Time*":    st.column_config.TextColumn("End DateTime* ✏️",     width=160),
                "Title*":            st.column_config.TextColumn("Title* ✏️",            width=250),
                "Description*":      st.column_config.TextColumn("Description* ✏️",      width=350),
                "Mandatory*":        st.column_config.SelectboxColumn("Mandatory*",
                                         options=["TRUE","FALSE"], width=100),
                "TLO":               st.column_config.TextColumn("TLO ✏️",               width=120),
                "Faculty Reg ID*":   st.column_config.TextColumn("Faculty Reg ID* ✏️",   width=160),
            },
            height=450,
        )
        st.session_state.preview_df = edited_df

    st.divider()

    # ═════════════════════════════════════════════════════════════════
    # STEP 6 — GENERATE
    # ═════════════════════════════════════════════════════════════════
    sh("🚀 Step 6 — Generate Karo")

    can_gen = bool(
        st.session_state.preview_df is not None and
        att_data and faculty_id.strip()
    )
    if not can_gen:
        missing = []
        if not att_data:                          missing.append("Attendance Sheet")
        if not faculty_id.strip():                missing.append("Faculty Registration ID")
        if st.session_state.preview_df is None:  missing.append("Step 5 mein Preview Table banao")
        if missing: ib(f"Pehle karo: <b>{', '.join(missing)}</b>")

    gen_btn = st.button("⚡ Generate Session Sheet + All Attendance Files",
                        type="primary", disabled=not can_gen,
                        use_container_width=True)

    if gen_btn and can_gen:
        errors = []
        prog   = st.progress(0)
        log    = st.empty()

        try:
            df = st.session_state.preview_df

            # Build final session rows from edited table
            log.markdown('<div class="ib">📋 Session sheet rows prepare ho rahi hain…</div>',
                         unsafe_allow_html=True)
            session_rows = []
            for _, row in df.iterrows():
                session_rows.append({
                    "module":      str(row.get("Module Name*","")),
                    "start_dt":    str(row.get("Start Date Time*","")),
                    "end_dt":      str(row.get("End Date Time*","")),
                    "title":       str(row.get("Title*","")),
                    "description": str(row.get("Description*","")),
                    "mandatory":   str(row.get("Mandatory*","TRUE")),
                    "tlo":         str(row.get("TLO","TLO1")),
                    "faculty_id":  str(row.get("Faculty Reg ID*",faculty_id.strip())),
                })
            prog.progress(20)

            # Session sheet
            log.markdown('<div class="ib">📊 Session sheet generate ho rahi hai…</div>',
                         unsafe_allow_html=True)
            session_xlsx = generate_session_sheet(session_rows)
            sb(f"Session sheet: <b>{len(session_rows)} rows</b> written in DL.xlsx format.")
            prog.progress(50)

            # Day-wise attendance
            daywise = {}
            if att_tpl:
                log.markdown('<div class="ib">🗂️ Day-wise attendance files bana raha hoon…</div>',
                             unsafe_allow_html=True)
                att_tpl_bytes = att_tpl.read()
                valid_dates   = [d for d in dates if d][:total_sessions]
                bar = st.progress(0)
                for i, d in enumerate(valid_dates):
                    try:
                        fname = f"attendance_{d.strftime('%Y-%m-%d')}.xlsx"
                        daywise[fname] = generate_daywise_attendance(att_tpl_bytes, att_data, d)
                    except Exception as ex:
                        errors.append(f"{d}: {ex}")
                    bar.progress((i+1)/max(len(valid_dates),1))
                sb(f"Day-wise attendance: <b>{len(daywise)} files</b> ready.")
            else:
                wb("DG Attendance Template nahi diya — sirf session sheet download hogi.")

            prog.progress(85)

            # ZIP
            zip_bytes = build_zip(session_xlsx, daywise)
            prog.progress(100)
            log.markdown('<div class="sb">🎉 Sab ready hai! Neeche se download karo.</div>',
                         unsafe_allow_html=True)

            st.session_state.results = {
                "session": session_xlsx,
                "daywise": daywise,
                "zip":     zip_bytes,
                "errors":  errors,
                "n":       len(session_rows),
            }

        except Exception as ex:
            eb(f"Error: {ex}")
            import traceback; st.code(traceback.format_exc())

    # ═════════════════════════════════════════════════════════════════
    # STEP 7 — DOWNLOADS
    # ═════════════════════════════════════════════════════════════════
    if st.session_state.results:
        res = st.session_state.results
        st.divider()
        sh("📥 Step 7 — Download Karo")

        for e in res.get("errors",[]): wb(f"Skipped: {e}")

        d1,d2 = st.columns(2)
        with d1:
            st.download_button("📦 ⬇ Download ALL as ZIP",
                data=res["zip"],
                file_name=f"DG_Output_{datetime.now():%Y%m%d_%H%M}.zip",
                mime="application/zip", use_container_width=True)
        with d2:
            st.download_button("📋 ⬇ Download session_sheet.xlsx",
                data=res["session"],
                file_name="session_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        if res["daywise"]:
            st.markdown(f"#### 📅 Day-wise Attendance ({len(res['daywise'])} files)")
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
                            use_container_width=True, key=f"dl_{fname}")

        st.divider()
        if st.button("🔄 Reset"):
            for k in ["att_data","syl_data","preview_df","results","_att_nm","_syl_nm"]:
                st.session_state[k] = None
            st.rerun()

    st.markdown("<center style='color:#aaa;font-size:.72rem;margin-top:1rem'>"
                "🔒 In-memory · No data stored on server</center>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
