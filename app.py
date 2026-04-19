"""
DG Sheet Generator v8.0  –  Fully Generalized
===============================================
• Works with ANY attendance sheet format, ANY syllabus format
• Dynamic units — auto-detected from syllabus (3, 5, 10 — whatever)
• Module Name asked separately per unit → goes in "Module Name*" column
• Month checkboxes from actual attendance dates
• Big clear UI, professional fonts
"""

import io, re, zipfile, warnings
from copy import copy
from collections import defaultdict, OrderedDict
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ─────────────────────────────────── PAGE CONFIG ─────────────────────
st.set_page_config(
    page_title="DG Sheet Generator",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
/* ── Global ── */
html, body, [class*="css"] { font-size: 15px; }
.main { padding: 0.5rem 1.5rem 2rem; }

/* ── Title banner ── */
.app-banner {
    background: linear-gradient(135deg, #0d2137 0%, #1a5f96 100%);
    color: #fff; padding: 1.4rem 2rem; border-radius: 14px;
    margin-bottom: 1.2rem; text-align: center;
    box-shadow: 0 4px 18px rgba(26,95,150,.35);
}
.app-banner h1 { margin: 0; font-size: 1.9rem; font-weight: 900; letter-spacing: -.3px; }
.app-banner p  { margin: .3rem 0 0; font-size: .97rem; opacity: .88; }

/* ── Section header ── */
.sec-hdr {
    font-size: 1.15rem; font-weight: 800; color: #0d2137;
    border-left: 5px solid #1a5f96; padding: .25rem 0 .25rem .7rem;
    margin: 1.1rem 0 .6rem;
}

/* ── Info boxes ── */
.box-info  { background:#e8f4fb; border-left:4px solid #1a5f96; padding:.55rem 1rem; border-radius:0 8px 8px 0; margin:.3rem 0; font-size:.88rem; }
.box-ok    { background:#d4edda; border-left:4px solid #2ecc71; padding:.55rem 1rem; border-radius:0 8px 8px 0; margin:.3rem 0; font-size:.88rem; }
.box-warn  { background:#fff8e1; border-left:4px solid #f39c12; padding:.55rem 1rem; border-radius:0 8px 8px 0; margin:.3rem 0; font-size:.88rem; }
.box-err   { background:#fde8e8; border-left:4px solid #e74c3c; padding:.55rem 1rem; border-radius:0 8px 8px 0; margin:.3rem 0; font-size:.88rem; }

/* ── Month cards ── */
.month-card {
    background: #f4f8fd; border: 2px solid #c8ddf0;
    border-radius: 10px; padding: .7rem; margin-bottom: .5rem; text-align: center;
}
.month-card-on {
    background: #daeeff; border: 2.5px solid #1a5f96;
    border-radius: 10px; padding: .7rem; margin-bottom: .5rem; text-align: center;
}
.month-name  { font-size: 1rem; font-weight: 800; color: #0d2137; }
.month-count { font-size: .8rem; color: #1a5f96; font-weight: 600; }
.date-chip {
    display: inline-block; background: #1a5f96; color: #fff;
    border-radius: 12px; padding: .1rem .5rem; margin: .12rem;
    font-size: .72rem; font-weight: 600;
}

/* ── Unit card ── */
.unit-card {
    background: #f8fafd; border: 1px solid #c8ddf0;
    border-radius: 10px; padding: 1rem 1.1rem; margin-bottom: .7rem;
}
.unit-num {
    display: inline-block; background: #1a5f96; color: #fff;
    border-radius: 50%; width: 28px; height: 28px; line-height: 28px;
    text-align: center; font-weight: 800; font-size: .88rem; margin-right: .4rem;
}

/* ── Timing ── */
.day-enabled  { background:#e3f2fd; border:2px solid #1a5f96; border-radius:9px; padding:.7rem .5rem; text-align:center; }
.day-disabled { background:#f1f1f1; border:1px solid #ddd;    border-radius:9px; padding:.7rem .5rem; text-align:center; opacity:.55; }
.day-label    { font-size: .95rem; font-weight: 800; color: #0d2137; }
.day-count    { font-size: .75rem; color: #1a5f96; }
.timing-val   { font-size: .8rem; font-weight: 700; color: #0d2137; background:#fff; border-radius:6px; padding:.15rem .4rem; margin-top:.3rem; display:inline-block; }

/* ── Stat cards ── */
.stat-card { background:#fff; border:1px solid #d0e8f5; border-radius:10px; padding:.8rem 1rem; text-align:center; }
.stat-num  { font-size: 2rem; font-weight: 900; color: #1a5f96; line-height: 1; }
.stat-lbl  { font-size: .78rem; color: #555; margin-top: .25rem; }

/* ── Buttons ── */
.stButton>button { font-size: 1rem !important; font-weight: 700 !important; padding: .55rem 1.4rem !important; border-radius: 9px !important; }
.stDownloadButton>button {
    background: #0d2137 !important; color: #fff !important;
    border-radius: 9px !important; font-weight: 700 !important;
    font-size: .95rem !important; width: 100% !important;
}
.stDownloadButton>button:hover { background: #1a5f96 !important; }

/* ── Badge ── */
.badge { display:inline-block; background:#d4edda; color:#155724; padding:.15rem .6rem; border-radius:12px; font-size:.78rem; font-weight:700; }

/* ── Divider ── */
hr { border: none; border-top: 2px solid #e0edf8; margin: 1rem 0; }
</style>
""", unsafe_allow_html=True)

# ── Helper render functions ──────────────────────────────────────────
def ib(m):  st.markdown(f'<div class="box-info">ℹ️  {m}</div>',  unsafe_allow_html=True)
def sb(m):  st.markdown(f'<div class="box-ok">✅  {m}</div>',   unsafe_allow_html=True)
def wb(m):  st.markdown(f'<div class="box-warn">⚠️  {m}</div>', unsafe_allow_html=True)
def eb(m):  st.markdown(f'<div class="box-err">❌  {m}</div>',  unsafe_allow_html=True)
def sh(t):  st.markdown(f'<div class="sec-hdr">{t}</div>',      unsafe_allow_html=True)

ALL_DAYS  = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
DAY_SHORT = {d: d[:3] for d in ALL_DAYS}

# ─────────────────────────────────── UTILITIES ───────────────────────

def safe_isna(v) -> bool:
    try: return bool(pd.isna(v))
    except: return False


def parse_date(raw) -> date | None:
    if raw is None or safe_isna(raw): return None
    if isinstance(raw, datetime): return raw.date()
    if isinstance(raw, date):     return raw
    s = str(raw).strip()
    if not s or s.lower() in ("nat","nan","none","pd.nat"): return None
    s = re.sub(r"\([^)]*\)","",s)
    s = re.sub(r"\b(sub|Sub|SUB)\b","",s)
    s = re.sub(r"\s+"," ",s).strip()
    if not s: return None
    for fmt in ("%d %b %Y","%d %B %Y","%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%d %b %y"):
        try: return datetime.strptime(s,fmt).date()
        except: pass
    try:
        r = pd.to_datetime(s, dayfirst=True, errors="coerce")
        return None if pd.isna(r) else r.date()
    except: return None


def to_24h(h: int, m: str, ampm: str) -> str:
    hh = int(h)
    if ampm == "PM" and hh != 12: hh += 12
    if ampm == "AM" and hh == 12: hh = 0
    return f"{hh:02d}:{m}"


def pick_best_sheet(fb: bytes, hints: list = None) -> pd.DataFrame:
    """Pick the most data-rich sheet, optionally hint by name keywords."""
    xf = pd.ExcelFile(io.BytesIO(fb))
    ns = xf.sheet_names
    if len(ns) == 1:
        return pd.read_excel(io.BytesIO(fb), sheet_name=ns[0], header=None)
    if hints:
        for kw in hints:
            for s in ns:
                if kw.lower() in s.lower():
                    return pd.read_excel(io.BytesIO(fb), sheet_name=s, header=None)
    best, most = ns[0], -1
    for s in ns:
        df = pd.read_excel(io.BytesIO(fb), sheet_name=s, header=None)
        n  = int(df.notna().sum().sum())
        if n > most: most, best = n, s
    return pd.read_excel(io.BytesIO(fb), sheet_name=best, header=None)


# ─────────────────────────────── ATTENDANCE PARSER ───────────────────

def parse_attendance(fb: bytes) -> dict:
    """
    Robust parser — works with any wide-format attendance sheet.
    Finds the row with the most date-like values → that's the date row.
    Returns: {dates, students}
    """
    df = pick_best_sheet(fb, hints=["attendance","student","sheet"])

    # ── Find date row ─────────────────────────────────────────────────
    best_row_idx, best_count, best_first = None, 0, None
    for i, row in df.iterrows():
        hits, first = [], None
        for j, v in enumerate(row):
            d = parse_date(v)
            if d:
                hits.append(d)
                if first is None: first = j
        if len(hits) > best_count:
            best_count, best_row_idx, best_first = len(hits), i, first

    if best_row_idx is None or best_count < 2:
        raise ValueError(
            "❌ Could not find session dates. "
            "Make sure your attendance sheet has dates in a single row."
        )

    # ── Build column → date map ───────────────────────────────────────
    date_row = df.iloc[best_row_idx]
    col_to_date, dates_ordered, seen = {}, [], set()
    for j in range(best_first, df.shape[1]):
        d = parse_date(date_row.iloc[j])
        if d:
            col_to_date[j] = d
            if d not in seen: dates_ordered.append(d); seen.add(d)

    # ── Find student data columns ─────────────────────────────────────
    name_col = enroll_col = None
    for i in range(best_row_idx):
        row = df.iloc[i]
        rs  = " ".join(str(v).lower() for v in row if not safe_isna(v))
        if any(k in rs for k in ["name","enrol","roll","student"]):
            for j, v in enumerate(row):
                s = str(v).lower().strip()
                if "name" in s and name_col is None:      name_col   = j
                elif any(k in s for k in ["enrol","roll"]) and enroll_col is None:
                    enroll_col = j
            break

    # ── Parse student rows ────────────────────────────────────────────
    students = []
    for i in range(best_row_idx + 1, df.shape[0]):
        row = df.iloc[i]
        if row.notna().sum() < 3: continue
        name = (str(row.iloc[name_col]).strip()
                if name_col is not None and not safe_isna(row.iloc[name_col]) else "")
        if not name or name.lower() in ("nan","none",""): continue
        enroll = (str(row.iloc[enroll_col]).strip()
                  if enroll_col is not None and not safe_isna(row.iloc[enroll_col]) else "")
        att = {}
        for col_j, d in col_to_date.items():
            val = row.iloc[col_j]
            try:    att[d] = int(float(val)) if not safe_isna(val) else None
            except: att[d] = None
        students.append({"name": name, "enrollment": enroll, "att": att})

    return {"dates": dates_ordered, "students": students}


def group_by_month(dates: list) -> OrderedDict:
    """{ (year, month): [date, ...] } sorted chronologically."""
    result = defaultdict(list)
    for d in dates:
        if d: result[(d.year, d.month)].append(d)
    return OrderedDict(sorted(result.items()))


# ──────────────────────────────── SYLLABUS PARSER ────────────────────

def parse_syllabus(fb: bytes, filename: str) -> OrderedDict:
    """
    Generalized parser — detects any heading pattern.
    Returns OrderedDict: {unit_heading: [topic_str, ...]}
    Works with docx, xlsx, txt, csv.
    """
    ext = filename.lower().rsplit(".", 1)[-1]
    result = OrderedDict()

    # ── Common unit-heading patterns ──────────────────────────────────
    UNIT_RE = re.compile(
        r"^(UNIT|MODULE|CHAPTER|TOPIC|SECTION|PART|BLOCK)\s*[-:.]?\s*\d+",
        re.IGNORECASE
    )

    # Citation / legal case patterns to skip
    SKIP_RE = re.compile(
        r"\bv\.\b|\bAIR\b|\bILR\b|\bSCC\b|\bSCR\b|https?://|\(\d{4}\)\s+\d+|"
        r"^\s*\d+\.\s+[A-Z][a-z]+\s+v\.\s+|^\s*Case\s+Law",
        re.IGNORECASE
    )

    def add_topic(unit, text):
        t = text.strip()
        if not t or len(t) < 4: return
        if SKIP_RE.search(t): return
        result.setdefault(unit, []).append(t)

    if ext == "docx":
        import docx as _d
        doc = _d.Document(io.BytesIO(fb))
        cur = None
        for p in doc.paragraphs:
            t = p.text.strip()
            if not t: continue
            if UNIT_RE.match(t):
                cur = t
                result[cur] = []
            elif cur:
                add_topic(cur, t)
        # Also scan tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    t = cell.text.strip()
                    if not t: continue
                    if UNIT_RE.match(t):
                        cur = t; result[cur] = []
                    elif cur:
                        add_topic(cur, t)

    elif ext in ("xlsx", "xls"):
        df = pick_best_sheet(fb)
        cur = None
        for _, row in df.iterrows():
            for val in row:
                s = str(val).strip()
                if not s or s.lower() in ("nan","none",""): continue
                if UNIT_RE.match(s):
                    cur = s; result[cur] = []; break
                elif cur and len(s) > 4:
                    add_topic(cur, s); break

    elif ext in ("txt", "csv"):
        text = fb.decode("utf-8","ignore")
        cur  = None
        for line in text.splitlines():
            t = line.strip()
            if not t: continue
            if UNIT_RE.match(t):
                cur = t; result[cur] = []
            elif cur:
                add_topic(cur, t)

    # If no unit headings detected — treat whole file as one module
    if not result:
        result["Module 1"] = []
        if ext in ("txt","csv"):
            for line in fb.decode("utf-8","ignore").splitlines():
                add_topic("Module 1", line)
        elif ext == "docx":
            import docx as _d
            doc = _d.Document(io.BytesIO(fb))
            for p in doc.paragraphs:
                add_topic("Module 1", p.text)
        elif ext in ("xlsx","xls"):
            df = pick_best_sheet(fb)
            for _, row in df.iterrows():
                for val in row:
                    s = str(val).strip()
                    if s and s.lower() not in ("nan","none","") and len(s)>4:
                        add_topic("Module 1", s); break

    return OrderedDict((k,v) for k,v in result.items() if v)


# ───────────────────────────── AUTO DESCRIPTION ──────────────────────

def auto_desc(title: str, module_name: str) -> str:
    t   = title.strip()
    mod = re.sub(r"^(UNIT|MODULE|CHAPTER)\s*\d+[:\-.]?\s*","",module_name,flags=re.IGNORECASE).strip() or module_name
    tl  = t.lower()

    if any(w in tl for w in ["definition","define","meaning"]):
        core = re.sub(r"^(definition\s+(of\s+)?|define\s+)","",t,flags=re.IGNORECASE).strip()
        return (f"This session covers the statutory definition and essential elements of {core}. "
                f"Students will examine the legal provisions, judicial interpretations, and "
                f"practical significance within the framework of {mod}.")
    if any(w in tl for w in ["rights","duties","liability","liabilities","obligation"]):
        return (f"This session examines the rights, duties, and liabilities under {t}. "
                f"Students will analyse statutory provisions, landmark judgments, and the legal "
                f"consequences for parties in {mod}.")
    if any(w in tl for w in ["distinction","difference","compare","versus"," and "," vs "]):
        return (f"This session provides a comparative analysis of {t}. Students will identify "
                f"key distinctions through statutory provisions and case-law to apply "
                f"differential reasoning in {mod}.")
    if any(w in tl for w in ["type","kind","classif","categor","form","nature"]):
        return (f"This session explores the types, categories, and conceptual framework of {t}. "
                f"Students will study the legal significance of each category and their "
                f"practical application within {mod}.")
    if any(w in tl for w in ["termination","discharge","revocation","dissolution","end"]):
        return (f"This session covers the modes of {t} and legal consequences that follow. "
                f"Students will study statutory provisions, conditions, and judicial "
                f"precedents governing this aspect of {mod}.")
    if any(w in tl for w in ["creation","formation","essential","element","requisite"]):
        return (f"This session discusses the formation process and requisite legal elements "
                f"of {t}. Students will examine statutory requirements, judicial "
                f"interpretations, and practical illustrations relevant to {mod}.")
    if any(w in tl for w in ["remedy","remedies","compensation","damages","relief"]):
        return (f"This session discusses available remedies in {t}. Students will examine "
                f"statutory and equitable remedies, judicial approaches, and "
                f"computation of relief in {mod}.")
    if any(w in tl for w in ["case","judgment","ruling","court"," v."]):
        return (f"This session analyses {t} as a landmark judicial decision. Students will "
                f"examine the facts, legal issues, court reasoning, and the precedential "
                f"value of this ruling in {mod}.")
    return (f"This session provides a comprehensive study of {t} within the domain of {mod}. "
            f"Students will analyse relevant statutory provisions, judicial precedents, "
            f"and practical applications through structured discussion and case-based learning.")


# ──────────────────────────── TOPIC BALANCER ─────────────────────────

def balance_topics(unit_configs: list, total_sessions: int) -> list:
    """
    unit_configs: [{"module_name":str, "tlo":str, "topics":[str,...]}, ...]
    Returns list length = total_sessions: [{"module","tlo","title"}, ...]
    """
    all_items = []
    for uc in unit_configs:
        for t in uc["topics"]:
            all_items.append({"module": uc["module_name"],
                               "tlo":    uc["tlo"],
                               "title":  t})

    n = len(all_items)
    if n == 0:
        return [{"module":"Session","tlo":"TLO1","title":f"Session {i+1}"}
                for i in range(total_sessions)]

    result = []
    if n <= total_sessions:
        result = list(all_items)
        extra  = total_sessions - n
        for i in range(extra):
            base = all_items[i % n]
            result.append({"module": base["module"], "tlo": base["tlo"],
                           "title": base["title"] + f" — Revision {i//n + 2}"})
    else:
        # Proportional trim per unit
        total_avail = sum(len(uc["topics"]) for uc in unit_configs)
        for uc in unit_configs:
            n_take = max(1, round(total_sessions * len(uc["topics"]) / total_avail))
            n_take = min(n_take, len(uc["topics"]))
            for t in uc["topics"][:n_take]:
                result.append({"module": uc["module_name"], "tlo": uc["tlo"], "title": t})
        # Fine-tune
        while len(result) < total_sessions:
            result.append(result[-1] | {"title": result[-1]["title"]+" (Cont.)"})
        result = result[:total_sessions]

    return result


# ───────────────────────────── EXCEL GENERATORS ──────────────────────

HEADERS = ["Module Name*","Start Date Time*","End Date Time*","Title*",
           "Description*","Attendance Mandatory*","TLO","Teaching Faculty Registration ID*"]
WIDTHS  = [30, 22, 22, 42, 60, 22, 18, 32]
HF = PatternFill("solid", start_color="1F4E79")
HT = Font(bold=True, color="FFFFFF", size=11)
DA = PatternFill("solid", start_color="D6EAF8")
DB = PatternFill("solid", start_color="EBF5FB")
TN = Side(style="thin", color="BBBBBB")
BD = Border(left=TN, right=TN, top=TN, bottom=TN)
CT = Alignment(horizontal="center", vertical="center", wrap_text=True)
LF = Alignment(horizontal="left",   vertical="center", wrap_text=True)


def gen_session_sheet(rows: list) -> bytes:
    wb = Workbook(); ws = wb.active; ws.title = "Session Sheet"
    aligns = [LF,CT,CT,LF,LF,CT,CT,CT]
    for c,(h,w) in enumerate(zip(HEADERS,WIDTHS),1):
        cell = ws.cell(1,c,h)
        cell.fill=HF; cell.font=HT; cell.border=BD; cell.alignment=CT
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 24
    for i,row in enumerate(rows,2):
        fill = DA if i%2==0 else DB
        vals = [row.get(k,"") for k in
                ["module","start_dt","end_dt","title","description","mandatory","tlo","faculty_id"]]
        for c,(v,al) in enumerate(zip(vals,aligns),1):
            cell = ws.cell(i,c,v)
            cell.fill=fill; cell.border=BD; cell.alignment=al; cell.font=Font(size=10)
        ws.row_dimensions[i].height = 30
    ws.freeze_panes = "A2"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def gen_daywise_att(tpl_bytes: bytes, att_data: dict, session_date: date) -> bytes:
    wb = load_workbook(io.BytesIO(tpl_bytes)); ws = wb.active
    ec = rc = ac = None
    for c in range(1, ws.max_column+1):
        hv = str(ws.cell(1,c).value or "").lower().strip()
        if "email" in hv:                                      ec = c
        elif "registration" in hv or ("reg" in hv and "id" in hv): rc = c
        elif "attendance" in hv:                               ac = c
    if not all([ec,rc,ac]): ec,rc,ac = 1,2,3

    enroll_att = {str(st["enrollment"]).strip():
                  (1 if st["att"].get(session_date)==1 else 0)
                  for st in att_data["students"]}

    GF=PatternFill("solid",start_color="C6EFCE"); RF=PatternFill("solid",start_color="FFC7CE")
    GT=Font(color="006100",bold=True,size=10);    RT=Font(color="9C0006",bold=True,size=10)
    CA=Alignment(horizontal="center",vertical="center")

    for r in range(2, ws.max_row+1):
        reg = str(ws.cell(r,rc).value or "").strip()
        if not reg: continue
        status = "PRESENT" if enroll_att.get(reg,0)==1 else "ABSENT"
        cell = ws.cell(r,ac)
        cell.value=status; cell.fill=GF if status=="PRESENT" else RF
        cell.font=GT if status=="PRESENT" else RT; cell.alignment=CA

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def build_zip(session_bytes: bytes, daywise: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("session_sheet.xlsx", session_bytes)
        for fn,fb in daywise.items():
            zf.writestr(f"attendance/{fn}", fb)
    buf.seek(0); return buf.getvalue()


# ══════════════════════════════ MAIN UI ══════════════════════════════

def main():
    st.markdown("""
    <div class="app-banner">
      <h1>📋 DG Session & Attendance Sheet Generator</h1>
      <p>Any attendance format · Any syllabus format · Full/Custom topic selection · AM/PM timing</p>
    </div>""", unsafe_allow_html=True)

    # ── session state ──────────────────────────────────────────────────
    for k in ["att","syl","preview","results","_anm","_snm"]:
        if k not in st.session_state: st.session_state[k] = None

    # ══════════════════════════════════════════════════════════════════
    # STEP 1 — FILE UPLOADS
    # ══════════════════════════════════════════════════════════════════
    sh("📂  Step 1 — Files Upload Karo")
    f1,f2,f3 = st.columns(3)

    with f1:
        st.markdown("##### 📅 Master Attendance Sheet")
        att_file = st.file_uploader("",type=["xlsx","xls"],key="att_up",
                                    label_visibility="collapsed")
        if att_file:
            st.markdown(f'<span class="badge">✅ {att_file.name}</span>',
                        unsafe_allow_html=True)
    with f2:
        st.markdown("##### 📚 Syllabus File")
        syl_file = st.file_uploader("",type=["docx","xlsx","xls","txt","csv"],key="syl_up",
                                    label_visibility="collapsed")
        if syl_file:
            st.markdown(f'<span class="badge">✅ {syl_file.name}</span>',
                        unsafe_allow_html=True)
    with f3:
        st.markdown("##### 🗂️ DG Attendance Template")
        att_tpl  = st.file_uploader("",type=["xlsx"],key="atpl_up",
                                    label_visibility="collapsed")
        if att_tpl:
            st.markdown(f'<span class="badge">✅ {att_tpl.name}</span>',
                        unsafe_allow_html=True)

    # Auto-parse attendance
    if att_file and att_file.name != st.session_state._anm:
        try:
            st.session_state.att   = parse_attendance(att_file.read())
            st.session_state._anm  = att_file.name
            st.session_state.preview = None
            st.session_state.results = None
        except Exception as ex:
            eb(f"Attendance parse error: {ex}")

    # Auto-parse syllabus
    if syl_file and syl_file.name != st.session_state._snm:
        try:
            st.session_state.syl  = parse_syllabus(syl_file.read(), syl_file.name)
            st.session_state._snm = syl_file.name
        except Exception as ex:
            wb(f"Syllabus parse warning: {ex}")

    att_data  = st.session_state.att
    syl_data  = st.session_state.syl or OrderedDict()
    all_dates = att_data["dates"] if att_data else []
    month_map = group_by_month(all_dates)
    syl_units = list(syl_data.keys())

    # Quick stat strip
    if all_dates or syl_data:
        st.markdown("<br>",unsafe_allow_html=True)
        sc = st.columns(4)
        with sc[0]:
            st.markdown(f'<div class="stat-card"><div class="stat-num">{len(all_dates)}</div>'
                        f'<div class="stat-lbl">Session Dates in Sheet</div></div>',
                        unsafe_allow_html=True)
        with sc[1]:
            st.markdown(f'<div class="stat-card"><div class="stat-num">{len(month_map)}</div>'
                        f'<div class="stat-lbl">Months Available</div></div>',
                        unsafe_allow_html=True)
        with sc[2]:
            n_st = len(att_data["students"]) if att_data else 0
            st.markdown(f'<div class="stat-card"><div class="stat-num">{n_st}</div>'
                        f'<div class="stat-lbl">Students in Attendance</div></div>',
                        unsafe_allow_html=True)
        with sc[3]:
            n_tp = sum(len(v) for v in syl_data.values())
            st.markdown(f'<div class="stat-card"><div class="stat-num">{n_tp}</div>'
                        f'<div class="stat-lbl">Topics in Syllabus</div></div>',
                        unsafe_allow_html=True)

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 2 — SELECT MONTHS
    # ══════════════════════════════════════════════════════════════════
    sh("📅  Step 2 — Months Select Karo (Attendance Sheet se)")

    if not month_map:
        ib("Pehle Attendance Sheet upload karo — months yahan automatically dikhenge.")
        selected_dates = []
    else:
        ib("Jin months ka session banana hai unhe <b>tick karo</b>. "
           "Har month ke actual session dates chip mein dikhenge.")

        n_months = len(month_map)
        cols_per_row = min(n_months, 4)
        month_keys   = list(month_map.keys())
        selected_dates_all = []

        # Render month checkboxes in rows of 4
        for row_start in range(0, n_months, cols_per_row):
            row_keys = month_keys[row_start:row_start+cols_per_row]
            mcols = st.columns(cols_per_row)
            for ci,(ym) in enumerate(row_keys):
                ds    = month_map[ym]
                label = datetime(ym[0],ym[1],1).strftime("%B %Y")
                with mcols[ci]:
                    checked = st.checkbox(
                        f"**{label}**",
                        value=True,
                        key=f"m_{ym[0]}_{ym[1]}"
                    )
                    # Date chips
                    chips = "".join(
                        f'<span class="date-chip">{d.strftime("%d")} '
                        f'{d.strftime("%a")}</span>' for d in ds
                    )
                    card_cls = "month-card-on" if checked else "month-card"
                    st.markdown(
                        f'<div class="{card_cls}">'
                        f'<div class="month-name">{label}</div>'
                        f'<div class="month-count">🗓️ {len(ds)} sessions</div>'
                        f'<div style="margin-top:.35rem">{chips}</div>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                    if checked:
                        selected_dates_all.extend(ds)

        selected_dates = sorted(set(selected_dates_all))

        if selected_dates:
            sb(f"<b>{len(selected_dates)} dates</b> selected  "
               f"({selected_dates[0].strftime('%d %b %Y')} → "
               f"{selected_dates[-1].strftime('%d %b %Y')})")
        else:
            wb("Kam se kam ek month select karo.")

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 3 — BASIC SETTINGS
    # ══════════════════════════════════════════════════════════════════
    sh("⚙️  Step 3 — Basic Settings")

    bc = st.columns([1,1,1,1,1])
    with bc[0]:
        override = st.checkbox("✏️ Override session count", value=False)
        if override:
            total_sessions = st.number_input("Sessions", min_value=1, max_value=500,
                                              value=len(selected_dates) or 50)
        else:
            total_sessions = len(selected_dates)
            st.metric("📊 Total Sessions", total_sessions)

    with bc[1]:
        faculty_id = st.text_input("🧑‍🏫 Faculty Registration ID *",
                                   placeholder="e.g. IILMGG006412025")
    with bc[2]:
        mandatory = st.selectbox("✅ Mandatory?", ["TRUE","FALSE"])
    with bc[3]:
        tlo_max = st.number_input("📊 TLO Max (e.g. 5)", min_value=1, max_value=100, value=5)
    with bc[4]:
        tlo_prefix = st.text_input("TLO Prefix", value="TLO",
                                   help="e.g. 'TLO' → TLO1, TLO2 | 'CO' → CO1, CO2")

    all_tlos = [f"{tlo_prefix}{i}" for i in range(1, tlo_max+1)]

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 4 — TIMING (7 days, AM/PM)
    # ══════════════════════════════════════════════════════════════════
    sh("⏰  Step 4 — Lecture Timing (Din ke Hisab se)")
    ib("Jo din class hoti hai unhe <b>enable</b> karo. "
       "Start aur End time AM/PM format mein bharo.")

    detected_days = {}
    for d in selected_dates:
        dn = d.strftime("%A")
        detected_days[dn] = detected_days.get(dn,0)+1

    day_timing = {}
    cols7 = st.columns(7)

    for i, day in enumerate(ALL_DAYS):
        count = detected_days.get(day, 0)
        with cols7[i]:
            enabled = st.checkbox(
                f"**{DAY_SHORT[day]}**",
                value=(count > 0),
                key=f"de_{day}"
            )
            st.markdown(
                f'<div class="{"day-enabled" if enabled else "day-disabled"}">'
                f'<div class="day-label">{day[:3]}</div>'
                f'<div class="day-count">{"🔵 "+str(count)+" sessions" if count else "—"}</div>',
                unsafe_allow_html=True
            )

            if enabled:
                h_opt  = list(range(1,13))
                m_opt  = ["00","05","10","15","20","25","30","35","40","45","50","55"]
                ap_opt = ["AM","PM"]

                h1,m1,a1 = st.columns([3,3,3])
                with h1: sv_h = st.selectbox("",h_opt,  index=9, key=f"sh_{day}",label_visibility="collapsed")
                with m1: sv_m = st.selectbox("",m_opt,  index=0, key=f"sm_{day}",label_visibility="collapsed")
                with a1: sv_a = st.selectbox("",ap_opt, index=1, key=f"sa_{day}",label_visibility="collapsed")
                s24 = to_24h(sv_h, sv_m, sv_a)

                h2,m2,a2 = st.columns([3,3,3])
                with h2: ev_h = st.selectbox("",h_opt,  index=10,key=f"eh_{day}",label_visibility="collapsed")
                with m2: ev_m = st.selectbox("",m_opt,  index=0, key=f"em_{day}",label_visibility="collapsed")
                with a2: ev_a = st.selectbox("",ap_opt, index=1, key=f"ea_{day}",label_visibility="collapsed")
                e24 = to_24h(ev_h, ev_m, ev_a)

                day_timing[day] = (s24, e24)
                st.markdown(
                    f'<div class="timing-val">⏱ {s24} – {e24}</div></div>',
                    unsafe_allow_html=True)
            else:
                st.markdown('</div>', unsafe_allow_html=True)

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 5 — MODULE & TOPIC CONFIGURATION
    # (dynamic: as many units as syllabus has)
    # ══════════════════════════════════════════════════════════════════
    sh("📚  Step 5 — Module Configuration & Topic Selection")

    if not syl_units:
        wb("Syllabus upload nahi hua — topics manually bharne padenge.")
        unit_configs = []
    else:
        ib(f"Syllabus mein <b>{len(syl_units)} unit(s)</b> mili. "
           f"Har unit ke liye <b>Module Name</b>, <b>TLO</b>, aur <b>Topics</b> configure karo. "
           f"Full unit lo ya custom selection karo.")

        unit_configs = []
        for i, unit_key in enumerate(syl_units):
            avail = syl_data[unit_key]

            with st.expander(
                f"{'🟢' if i<len(all_tlos) else '⚪'} "
                f"Unit {i+1}  —  {unit_key[:60]}  ({len(avail)} topics)",
                expanded=(i < 5)   # open first 5, rest collapsed
            ):
                row1 = st.columns([3,2])
                with row1[0]:
                    # Module Name — THIS goes in "Module Name*" column of sheet
                    mod_name = st.text_input(
                        "📦 Module Name* (goes in session sheet)",
                        value=unit_key,
                        key=f"mn_{i}",
                        help="This exact text will appear in the 'Module Name*' column"
                    )
                with row1[1]:
                    def_tlo = [all_tlos[i % len(all_tlos)]] if all_tlos else [f"{tlo_prefix}1"]
                    sel_tlos = st.multiselect(
                        f"📊 TLOs for this module",
                        options=all_tlos,
                        default=def_tlo,
                        key=f"tl_{i}"
                    )
                    tlo_str = " | ".join(sel_tlos) if sel_tlos else all_tlos[i%len(all_tlos)]

                # Topic coverage
                cov_mode = st.radio(
                    "Topics:",
                    ["📖 Full Unit — pura syllabus",
                     "✂️ Custom — specific topics choose karo"],
                    key=f"cov_{i}",
                    horizontal=True
                )

                if "Custom" in cov_mode:
                    chosen = st.multiselect(
                        f"Topics select karo (Unit {i+1}):",
                        options=avail,
                        default=avail,
                        key=f"tp_{i}",
                        help="Sab pre-selected hain — jo nahi chahiye unhe hata do"
                    )
                    final_topics = chosen
                    st.caption(f"✅ {len(final_topics)} / {len(avail)} topics selected")
                else:
                    final_topics = avail
                    st.caption(f"📖 All {len(avail)} topics included")
                    with st.expander(f"See all {len(avail)} topics →"):
                        for j,t in enumerate(avail,1):
                            st.markdown(f"`{j}.` {t}")

                unit_configs.append({
                    "module_name": mod_name,
                    "tlo":         tlo_str,
                    "topics":      final_topics,
                })

        # Summary across all units
        total_topics_selected = sum(len(u["topics"]) for u in unit_configs)
        st.markdown("<br>",unsafe_allow_html=True)
        sc2 = st.columns(3)
        with sc2[0]:
            st.markdown(f'<div class="stat-card"><div class="stat-num">{len(unit_configs)}</div>'
                        f'<div class="stat-lbl">Modules Configured</div></div>',
                        unsafe_allow_html=True)
        with sc2[1]:
            st.markdown(f'<div class="stat-card"><div class="stat-num">{total_topics_selected}</div>'
                        f'<div class="stat-lbl">Topics Selected</div></div>',
                        unsafe_allow_html=True)
        with sc2[2]:
            st.markdown(f'<div class="stat-card"><div class="stat-num">{total_sessions}</div>'
                        f'<div class="stat-lbl">Sessions to Generate</div></div>',
                        unsafe_allow_html=True)
        st.markdown("<br>",unsafe_allow_html=True)

        if total_topics_selected == total_sessions:
            sb("<b>Perfect match!</b> Topics = Sessions — 1 topic per session.")
        elif total_topics_selected < total_sessions:
            wb(f"Topics ({total_topics_selected}) < Sessions ({total_sessions}). "
               f"Last {total_sessions-total_topics_selected} sessions = revision entries.")
        else:
            wb(f"Topics ({total_topics_selected}) > Sessions ({total_sessions}). "
               f"Topics will be trimmed proportionally per unit.")

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 6 — BUILD PREVIEW
    # ══════════════════════════════════════════════════════════════════
    sh("👀  Step 6 — Preview Table Banao & Edit Karo")
    ib("'Build Preview' dabao → auto-balanced titles + auto-generated descriptions dikhenge. "
       "Table mein directly kuch bhi edit kar sakte ho before generating.")

    prev_btn = st.button("🔄  Build Preview Table", type="secondary",
                         use_container_width=True)

    if prev_btn:
        if not selected_dates:
            eb("Step 2 mein months select karo pehle.")
        elif not faculty_id.strip():
            eb("Faculty Registration ID daalo (Step 3).")
        elif not unit_configs and syl_units:
            eb("Step 5 mein module config complete karo.")
        else:
            if unit_configs:
                flat = balance_topics(unit_configs, total_sessions)
            else:
                flat = [{"module":"Module","tlo":f"{tlo_prefix}1","title":f"Session {i+1}"}
                        for i in range(total_sessions)]

            use_dates = selected_dates[:total_sessions]
            rows = []
            for idx in range(total_sessions):
                if idx < len(use_dates):
                    d   = use_dates[idx]
                    ds  = d.strftime("%Y-%m-%d")
                    dn  = d.strftime("%A")
                    s,e = day_timing.get(dn, ("10:00","11:00"))
                    start_dt = f"{ds} {s}:00"
                    end_dt   = f"{ds} {e}:00"
                else:
                    start_dt = end_dt = "TBD"

                sess = flat[idx] if idx < len(flat) else {
                    "module":"Extra","tlo":f"{tlo_prefix}1",
                    "title":f"Extra Session {idx+1}"}

                rows.append({
                    "Sr":                idx+1,
                    "Module Name*":      sess["module"],
                    "Start Date Time*":  start_dt,
                    "End Date Time*":    end_dt,
                    "Title*":            sess["title"],
                    "Description*":      auto_desc(sess["title"], sess["module"]),
                    "Mandatory*":        mandatory,
                    "TLO":               sess["tlo"],
                    "Faculty Reg ID*":   faculty_id.strip(),
                })

            st.session_state.preview = pd.DataFrame(rows)
            st.session_state.results = None
            sb(f"Preview ready: <b>{len(rows)} rows</b>")

    if st.session_state.preview is not None:
        df = st.session_state.preview
        edited = st.data_editor(
            df, use_container_width=True, num_rows="fixed",
            hide_index=True, key="ed_prev",
            column_config={
                "Sr":               st.column_config.NumberColumn("Sr",width=55,disabled=True),
                "Module Name*":     st.column_config.TextColumn("Module Name* ✏️",    width=230),
                "Start Date Time*": st.column_config.TextColumn("Start DateTime* ✏️", width=168),
                "End Date Time*":   st.column_config.TextColumn("End DateTime* ✏️",   width=168),
                "Title*":           st.column_config.TextColumn("Title* ✏️",          width=265),
                "Description*":     st.column_config.TextColumn("Description* ✏️",    width=380),
                "Mandatory*":       st.column_config.SelectboxColumn("Mandatory*",
                                        options=["TRUE","FALSE"],width=110),
                "TLO":              st.column_config.TextColumn("TLO ✏️",             width=125),
                "Faculty Reg ID*":  st.column_config.TextColumn("Faculty Reg ID* ✏️", width=168),
            },
            height=480,
        )
        st.session_state.preview = edited

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 7 — GENERATE
    # ══════════════════════════════════════════════════════════════════
    sh("🚀  Step 7 — Generate Karo")

    can_gen = bool(
        st.session_state.preview is not None and
        att_data and faculty_id.strip()
    )
    if not can_gen:
        miss = []
        if not att_data:                        miss.append("Attendance Sheet")
        if not faculty_id.strip():              miss.append("Faculty Registration ID")
        if st.session_state.preview is None:   miss.append("Step 6: Preview Table banao")
        if miss: ib(f"Pehle karo: <b>{', '.join(miss)}</b>")

    gen_btn = st.button(
        "⚡  Generate Session Sheet + All Attendance Files",
        type="primary", disabled=not can_gen, use_container_width=True
    )

    if gen_btn and can_gen:
        errors = []; prog = st.progress(0); log = st.empty()
        try:
            df = st.session_state.preview
            session_rows = [{
                "module":      str(r.get("Module Name*","")),
                "start_dt":    str(r.get("Start Date Time*","")),
                "end_dt":      str(r.get("End Date Time*","")),
                "title":       str(r.get("Title*","")),
                "description": str(r.get("Description*","")),
                "mandatory":   str(r.get("Mandatory*","TRUE")),
                "tlo":         str(r.get("TLO",f"{tlo_prefix}1")),
                "faculty_id":  str(r.get("Faculty Reg ID*", faculty_id.strip())),
            } for _,r in df.iterrows()]
            prog.progress(20)

            log.markdown('<div class="box-info">📊 Session sheet generating…</div>',
                         unsafe_allow_html=True)
            session_xlsx = gen_session_sheet(session_rows)
            sb(f"Session sheet: <b>{len(session_rows)} rows</b> — DL.xlsx exact format ✅")
            prog.progress(50)

            daywise = {}
            if att_tpl:
                log.markdown('<div class="box-info">🗂️ Day-wise attendance generating…</div>',
                             unsafe_allow_html=True)
                att_tpl_b = att_tpl.read()
                bar = st.progress(0)
                use_d = selected_dates[:total_sessions]
                for i,d in enumerate(use_d):
                    try:
                        fname = f"attendance_{d.strftime('%Y-%m-%d')}.xlsx"
                        daywise[fname] = gen_daywise_att(att_tpl_b, att_data, d)
                    except Exception as ex: errors.append(f"{d}: {ex}")
                    bar.progress((i+1)/max(len(use_d),1))
                sb(f"Day-wise attendance: <b>{len(daywise)} files</b> ✅")
            else:
                wb("Attendance Template nahi diya — sirf session sheet download hogi.")

            prog.progress(85)
            zip_bytes = build_zip(session_xlsx, daywise)
            prog.progress(100)
            log.markdown('<div class="box-ok">🎉 Sab ready hai! Download karo.</div>',
                         unsafe_allow_html=True)

            st.session_state.results = {
                "session": session_xlsx, "daywise": daywise,
                "zip": zip_bytes, "errors": errors, "n": len(session_rows),
            }
        except Exception as ex:
            eb(f"Error: {ex}")
            import traceback; st.code(traceback.format_exc())

    # ══════════════════════════════════════════════════════════════════
    # STEP 8 — DOWNLOADS
    # ══════════════════════════════════════════════════════════════════
    if st.session_state.results:
        res = st.session_state.results
        st.divider()
        sh("📥  Step 8 — Download Karo")
        for e in res.get("errors",[]): wb(f"Skipped: {e}")

        dl1,dl2 = st.columns(2)
        with dl1:
            st.download_button(
                "📦  ⬇ Download ALL as ZIP",
                data=res["zip"],
                file_name=f"DG_Output_{datetime.now():%Y%m%d_%H%M}.zip",
                mime="application/zip",
                use_container_width=True
            )
        with dl2:
            st.download_button(
                "📋  ⬇ Download session_sheet.xlsx",
                data=res["session"], file_name="session_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        if res["daywise"]:
            st.markdown(f"#### 📅 Day-wise Attendance ({len(res['daywise'])} files)")
            files = list(res["daywise"].items())
            for ri in range(0, len(files), 4):
                chunk = files[ri:ri+4]; cols = st.columns(4)
                for ci,(fn,fb) in enumerate(chunk):
                    with cols[ci]:
                        dp = fn.replace("attendance_","").replace(".xlsx","")
                        try:    lbl = datetime.strptime(dp,"%Y-%m-%d").strftime("%d %b %Y")
                        except: lbl = dp
                        st.download_button(f"📅 {lbl}", data=fb, file_name=fn,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True, key=f"dl_{fn}")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 Reset — Start Over"):
            for k in ["att","syl","preview","results","_anm","_snm"]:
                st.session_state[k] = None
            st.rerun()

    st.markdown(
        "<center style='color:#aaa;font-size:.73rem;margin-top:1.2rem'>"
        "🔒  All processing in-memory — no data stored on server  |  "
        "Works with any attendance format & any syllabus format"
        "</center>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
