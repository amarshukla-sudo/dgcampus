"""
DG Sheet Generator  v11.0
Created by Dr. Amar Shukla · IILM University, Gurugram
=========================================================
FIXES:
  1. ATTENDANCE LOGIC (CORRECT):
     - Student in template AND in master → use master value (1=PRESENT, 0=ABSENT)
     - Student in template but NOT in master → ABSENT (no record = not present)
     - Works for any section's template + any master attendance sheet
  2. TIMING UI → simple text inputs per day (no ugly dropdowns)
  3. SESSION SHEET → exact DL.xlsx format, all cells as TEXT
  4. Works with ANY file format for attendance and syllabus
"""

import io, re, zipfile, warnings
from collections import defaultdict, OrderedDict
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DG Sheet Generator — Dr. Amar Shukla",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
html, body, [class*="css"] { font-size: 16px !important; }
.main { padding: 0.5rem 2rem 3rem; }

.banner {
    background: linear-gradient(135deg,#0a1f35,#1565c0);
    color:#fff; padding:1.4rem 2.2rem; border-radius:14px; margin-bottom:1.1rem;
}
.banner h1  { margin:0; font-size:1.95rem; font-weight:900; }
.banner p   { margin:.35rem 0 0; font-size:.97rem; opacity:.88; }
.banner small { font-size:.82rem; opacity:.65; font-style:italic; }

.sec { font-size:1.15rem; font-weight:800; color:#0a1f35;
       border-left:6px solid #1565c0; padding:.28rem 0 .28rem .7rem;
       margin:1.2rem 0 .6rem; }

.ib { background:#e3f2fd; border-left:5px solid #1565c0; border-radius:0 8px 8px 0;
      padding:.55rem 1rem; margin:.3rem 0; font-size:.93rem; color:#0d47a1; }
.ok { background:#e8f5e9; border-left:5px solid #2e7d32; border-radius:0 8px 8px 0;
      padding:.55rem 1rem; margin:.3rem 0; font-size:.93rem; color:#1b5e20; }
.wn { background:#fff8e1; border-left:5px solid #f57f17; border-radius:0 8px 8px 0;
      padding:.55rem 1rem; margin:.3rem 0; font-size:.93rem; color:#e65100; }
.er { background:#ffebee; border-left:5px solid #c62828; border-radius:0 8px 8px 0;
      padding:.55rem 1rem; margin:.3rem 0; font-size:.93rem; color:#b71c1c; }

.stat-row { display:flex; gap:.8rem; margin:.6rem 0; flex-wrap:wrap; }
.sc { background:#fff; border:2px solid #e3f2fd; border-radius:12px;
      padding:.85rem 1.1rem; text-align:center; flex:1; min-width:110px; }
.sn { font-size:2.1rem; font-weight:900; color:#1565c0; line-height:1.1; }
.sl { font-size:.8rem; color:#546e7a; margin-top:.2rem; font-weight:600; }

.date-chip { display:inline-block; background:#1565c0; color:#fff;
             border-radius:14px; padding:.14rem .52rem; margin:.1rem;
             font-size:.75rem; font-weight:700; }

.stButton>button { font-size:1rem !important; font-weight:700 !important;
                   padding:.58rem 1.4rem !important; border-radius:10px !important; }
.stDownloadButton>button { background:#0a1f35 !important; color:#fff !important;
    border-radius:10px !important; font-weight:700 !important;
    font-size:.92rem !important; width:100% !important; }
.stDownloadButton>button:hover { background:#1565c0 !important; }

label { font-size:1rem !important; font-weight:600 !important; }
.streamlit-expanderHeader { font-size:1rem !important; font-weight:700 !important; }
.badge { display:inline-block; background:#e8f5e9; color:#2e7d32;
         padding:.13rem .6rem; border-radius:12px; font-size:.78rem; font-weight:700; }
hr { border:none; border-top:2.5px solid #e3f2fd; margin:1.1rem 0; }
.footer { text-align:center; color:#90a4ae; font-size:.8rem; margin-top:1.8rem;
          padding-top:1rem; border-top:2px solid #e3f2fd; }
</style>
""", unsafe_allow_html=True)

def ib(m): st.markdown(f'<div class="ib">ℹ️  {m}</div>', unsafe_allow_html=True)
def ok(m): st.markdown(f'<div class="ok">✅  {m}</div>', unsafe_allow_html=True)
def wn(m): st.markdown(f'<div class="wn">⚠️  {m}</div>', unsafe_allow_html=True)
def er(m): st.markdown(f'<div class="er">❌  {m}</div>', unsafe_allow_html=True)
def sec(t): st.markdown(f'<div class="sec">{t}</div>', unsafe_allow_html=True)

ALL_DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]

# ══════════════════════════════════════════════════════════════════════
# UTILITIES
# ══════════════════════════════════════════════════════════════════════

def safe_isna(v):
    try: return bool(pd.isna(v))
    except: return False

def parse_date(raw) -> date | None:
    if raw is None or safe_isna(raw): return None
    if isinstance(raw, datetime): return raw.date()
    if isinstance(raw, date):     return raw
    s = str(raw).strip()
    if not s or s.lower() in ("nat","nan","none","pd.nat"): return None
    s = re.sub(r"\([^)]*\)","",s); s = re.sub(r"\b(sub|Sub|SUB)\b","",s)
    s = re.sub(r"\s+"," ",s).strip()
    if not s: return None
    for fmt in ("%d %b %Y","%d %B %Y","%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%d %b %y"):
        try: return datetime.strptime(s, fmt).date()
        except: pass
    try:
        r = pd.to_datetime(s, dayfirst=True, errors="coerce")
        return None if pd.isna(r) else r.date()
    except: return None

def parse_time(s: str) -> tuple:
    """Parse '10:00' or '10:00 AM' or '10.00' → (HH, MM)"""
    s = s.strip().upper()
    am_pm = None
    if "AM" in s or "PM" in s:
        am_pm = "PM" if "PM" in s else "AM"
        s = s.replace("AM","").replace("PM","").strip()
    s = s.replace(".",":").replace("-",":")
    parts = re.split(r"[:h]", s)
    try:
        h, m = int(parts[0]), int(parts[1]) if len(parts)>1 else 0
        if am_pm == "PM" and h != 12: h += 12
        if am_pm == "AM" and h == 12: h = 0
        return (h, m)
    except: return (10, 0)

def fmt_dt(d: date, h: int, m: int) -> str:
    return f"{d.strftime('%Y-%m-%d')} {h:02d}:{m:02d}:00"

def best_sheet(fb: bytes, hints=None) -> pd.DataFrame:
    xf = pd.ExcelFile(io.BytesIO(fb)); ns = xf.sheet_names
    if len(ns)==1: return pd.read_excel(io.BytesIO(fb), sheet_name=ns[0], header=None)
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

# ══════════════════════════════════════════════════════════════════════
# ATTENDANCE PARSER — reads master sheet
# ══════════════════════════════════════════════════════════════════════

def parse_attendance(fb: bytes) -> dict:
    df = best_sheet(fb, hints=["attendance","student"])

    # Find row with most date values
    best_row, best_cnt, best_first = None, 0, None
    for i, row in df.iterrows():
        hits, first = [], None
        for j, v in enumerate(row):
            d = parse_date(v)
            if d: hits.append(d); first = (j if first is None else first)
        if len(hits) > best_cnt:
            best_cnt, best_row, best_first = len(hits), i, first

    if best_row is None or best_cnt < 2:
        raise ValueError("Could not find session dates in the attendance sheet. "
                         "Ensure dates appear in a single row.")

    # Build date → column map
    date_row = df.iloc[best_row]
    col_to_date, dates_ordered, seen = {}, [], set()
    for j in range(best_first, df.shape[1]):
        d = parse_date(date_row.iloc[j])
        if d:
            col_to_date[j] = d
            if d not in seen: dates_ordered.append(d); seen.add(d)

    # Find enrollment column
    enroll_col = None
    for i in range(best_row):
        row = df.iloc[i]
        rs  = " ".join(str(v).lower() for v in row if not safe_isna(v))
        if any(k in rs for k in ["enrol","roll","student id","registration"]):
            for j, v in enumerate(row):
                s = str(v).lower().strip()
                if any(k in s for k in ["enrol","roll"]) and enroll_col is None:
                    enroll_col = j
            if enroll_col is not None: break

    # Also try column 1 as fallback (typical pattern)
    if enroll_col is None: enroll_col = 1

    # Parse per-date attendance per student: {date: {enrollment: 0/1}}
    date_att: dict[date, dict[str, int]] = {d: {} for d in dates_ordered}

    for i in range(best_row+1, df.shape[0]):
        row = df.iloc[i]
        if row.notna().sum() < 3: continue
        enroll = str(row.iloc[enroll_col]).strip() if not safe_isna(row.iloc[enroll_col]) else ""
        # Remove decimal if numeric enrollment (e.g. 2571034.0)
        enroll = re.sub(r'\.0+$', '', enroll)
        if not enroll or enroll.lower() in ("nan","none",""): continue

        for col_j, d in col_to_date.items():
            val = row.iloc[col_j]
            try:    att_val = int(float(val)) if not safe_isna(val) else 0
            except: att_val = 0
            date_att[d][enroll] = att_val

    return {"dates": dates_ordered, "date_att": date_att}

def group_by_month(dates: list) -> OrderedDict:
    res = defaultdict(list)
    for d in dates:
        if d: res[(d.year, d.month)].append(d)
    return OrderedDict(sorted(res.items()))

# ══════════════════════════════════════════════════════════════════════
# SYLLABUS PARSER — any format
# ══════════════════════════════════════════════════════════════════════

UNIT_RE = re.compile(
    r"^(UNIT|MODULE|CHAPTER|TOPIC|SECTION|PART|BLOCK)\s*[-:.]?\s*\d+",
    re.IGNORECASE
)
SKIP_RE = re.compile(
    r"\bv\.\b|\bAIR\b|\bILR\b|\bSCC\b|https?://|\(\d{4}\)\s+\d+|^\s*Case\s+Law",
    re.IGNORECASE
)

def parse_syllabus(fb: bytes, filename: str) -> OrderedDict:
    ext = filename.lower().rsplit(".", 1)[-1]
    result = OrderedDict()

    def add(unit, text):
        t = text.strip()
        if not t or len(t)<4 or SKIP_RE.search(t): return
        result.setdefault(unit, []).append(t)

    if ext == "docx":
        import docx as _d
        doc = _d.Document(io.BytesIO(fb)); cur = None
        for p in doc.paragraphs:
            t = p.text.strip()
            if not t: continue
            if UNIT_RE.match(t): cur = t; result[cur] = []
            elif cur: add(cur, t)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    t = cell.text.strip()
                    if not t: continue
                    if UNIT_RE.match(t): cur = t; result[cur] = []
                    elif cur: add(cur, t)

    elif ext in ("xlsx","xls"):
        df = best_sheet(fb); cur = None
        for _, row in df.iterrows():
            for val in row:
                s = str(val).strip()
                if not s or s.lower() in ("nan","none",""): continue
                if UNIT_RE.match(s): cur = s; result[cur] = []; break
                elif cur and len(s)>4: add(cur,s); break

    elif ext in ("txt","csv"):
        cur = None
        for line in fb.decode("utf-8","ignore").splitlines():
            t = line.strip()
            if not t: continue
            if UNIT_RE.match(t): cur = t; result[cur] = []
            elif cur: add(cur, t)

    if not result:
        result["Module 1"] = []
        if ext in ("txt","csv"):
            for line in fb.decode("utf-8","ignore").splitlines(): add("Module 1", line)
        elif ext == "docx":
            import docx as _d
            for p in _d.Document(io.BytesIO(fb)).paragraphs: add("Module 1", p.text)
        elif ext in ("xlsx","xls"):
            df = best_sheet(fb)
            for _, row in df.iterrows():
                for val in row:
                    s = str(val).strip()
                    if s and s.lower() not in ("nan","none","") and len(s)>4:
                        add("Module 1",s); break

    return OrderedDict((k,v) for k,v in result.items() if v)

# ══════════════════════════════════════════════════════════════════════
# AUTO-DESCRIPTION
# ══════════════════════════════════════════════════════════════════════

def auto_desc(title: str, module: str) -> str:
    t   = title.strip()
    mod = re.sub(r"^(UNIT|MODULE|CHAPTER)\s*\d+[:\-.]?\s*","",module,flags=re.IGNORECASE).strip() or module
    tl  = t.lower()
    if any(w in tl for w in ["definition","define","meaning"]):
        core = re.sub(r"^(definition\s+(of\s+)?|define\s+)","",t,flags=re.IGNORECASE).strip()
        return (f"This session covers the statutory definition and essential elements of {core}. "
                f"Students will examine the relevant legal provisions, judicial interpretations, "
                f"and practical significance within the framework of {mod}.")
    if any(w in tl for w in ["rights","duties","liability","liabilities","obligation"]):
        return (f"This session examines the rights, duties, and liabilities under {t}. "
                f"Students will analyse statutory provisions, landmark judgments, and the legal "
                f"consequences for parties involved in {mod}.")
    if any(w in tl for w in ["distinction","difference","compare","versus"]):
        return (f"This session provides a comparative analysis of {t}. Students will identify "
                f"key distinctions through statutory provisions and case-law, enabling "
                f"differential reasoning in {mod}.")
    if any(w in tl for w in ["type","kind","classif","categor","nature","form"]):
        return (f"This session explores the types, categories, and conceptual framework of {t}. "
                f"Students will study the legal significance of each category and its "
                f"practical application within {mod}.")
    if any(w in tl for w in ["termination","discharge","revocation","dissolution"]):
        return (f"This session covers the modes of {t} and the legal consequences that follow. "
                f"Students will study statutory provisions, conditions, and judicial "
                f"precedents governing this aspect of {mod}.")
    if any(w in tl for w in ["creation","formation","essential","element"]):
        return (f"This session discusses the formation process and requisite elements of {t}. "
                f"Students will examine statutory requirements, judicial interpretations, "
                f"and practical illustrations relevant to {mod}.")
    if any(w in tl for w in ["remedy","remedies","damages","compensation"]):
        return (f"This session discusses remedies available in {t}. Students will examine "
                f"statutory and equitable remedies, judicial approaches, and "
                f"computation of relief in {mod}.")
    if any(w in tl for w in ["case","judgment"," v."]):
        return (f"This session analyses {t} as a significant judicial decision. Students will "
                f"examine the facts, legal issues, court reasoning, and the precedential "
                f"value of this ruling in {mod}.")
    return (f"This session provides a comprehensive study of {t} within {mod}. "
            f"Students will analyse relevant statutory provisions, judicial precedents, "
            f"and practical applications through structured discussion and case-based learning.")

# ══════════════════════════════════════════════════════════════════════
# TOPIC BALANCER
# ══════════════════════════════════════════════════════════════════════

def balance_topics(unit_configs: list, total_sessions: int) -> list:
    all_items = []
    for uc in unit_configs:
        for t in uc["topics"]:
            all_items.append({"module":uc["module_name"],"tlo":uc["tlo"],"title":t})
    n = len(all_items)
    if n == 0:
        return [{"module":"Session","tlo":"TLO1","title":f"Session {i+1}"}
                for i in range(total_sessions)]
    if n <= total_sessions:
        result = list(all_items)
        for i in range(total_sessions-n):
            base = all_items[i%n]
            result.append({"module":base["module"],"tlo":base["tlo"],
                           "title":base["title"]+" — Revision"})
        return result
    total_avail = sum(len(uc["topics"]) for uc in unit_configs)
    result = []
    for uc in unit_configs:
        n_take = max(1, round(total_sessions*len(uc["topics"])/total_avail))
        n_take = min(n_take, len(uc["topics"]))
        for t in uc["topics"][:n_take]:
            result.append({"module":uc["module_name"],"tlo":uc["tlo"],"title":t})
    while len(result) < total_sessions:
        result.append(result[-1]|{"title":result[-1]["title"]+" (Cont.)"})
    return result[:total_sessions]

# ══════════════════════════════════════════════════════════════════════
# SESSION SHEET GENERATOR — exact DL.xlsx format, all cells TEXT
# ══════════════════════════════════════════════════════════════════════

HEADERS = ["Module Name*","Start Date Time*","End Date Time*","Title*",
           "Description*","Attendance Mandatory*","TLO","Teaching Faculty Registration ID*"]
COL_W   = [30,22,22,42,60,22,18,32]
HF = PatternFill("solid", start_color="1F4E79")
HT = Font(bold=True, color="FFFFFF", size=11)
DA = PatternFill("solid", start_color="D6EAF8")
DB = PatternFill("solid", start_color="EBF5FB")
TN = Side(style="thin", color="BBBBBB")
BD = Border(left=TN,right=TN,top=TN,bottom=TN)
CT = Alignment(horizontal="center",vertical="center",wrap_text=True)
LF = Alignment(horizontal="left",  vertical="center",wrap_text=True)

def gen_session_sheet(rows: list) -> bytes:
    wb = Workbook(); ws = wb.active; ws.title = "Session Sheet"
    aligns = [LF,CT,CT,LF,LF,CT,CT,CT]
    for c,(h,w) in enumerate(zip(HEADERS,COL_W),1):
        cell=ws.cell(1,c,str(h)); cell.fill=HF; cell.font=HT
        cell.border=BD; cell.alignment=CT; cell.number_format="@"
        ws.column_dimensions[get_column_letter(c)].width=w
    ws.row_dimensions[1].height=24
    for i,row in enumerate(rows,2):
        fill=DA if i%2==0 else DB
        vals=[row.get(k,"") for k in
              ["module","start_dt","end_dt","title","description","mandatory","tlo","faculty_id"]]
        for c,(v,al) in enumerate(zip(vals,aligns),1):
            cell=ws.cell(i,c,str(v)); cell.fill=fill
            cell.border=BD; cell.alignment=al
            cell.font=Font(size=10); cell.number_format="@"
        ws.row_dimensions[i].height=30
    ws.freeze_panes="A2"
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════
# DAY-WISE ATTENDANCE — CORRECT LOGIC
#
# RULE (per student in DG template):
#   reg in master AND master value == 1  →  PRESENT
#   reg in master AND master value == 0  →  ABSENT
#   reg NOT in master                    →  ABSENT  (no record = not present)
# ══════════════════════════════════════════════════════════════════════

def gen_daywise_att(tpl_bytes: bytes, att_data: dict, session_date: date) -> bytes:
    wb = load_workbook(io.BytesIO(tpl_bytes)); ws = wb.active

    # Detect header columns
    ec = rc = ac = None
    for c in range(1, ws.max_column+1):
        hv = str(ws.cell(1,c).value or "").lower().strip()
        if "email" in hv:                                          ec = c
        elif "registration" in hv or ("reg" in hv and "id" in hv): rc = c
        elif "attendance" in hv:                                   ac = c
    if not all([ec,rc,ac]): ec,rc,ac=1,2,3

    # Get master attendance for this date
    day_att: dict[str,int] = att_data["date_att"].get(session_date, {})

    GF = PatternFill("solid", start_color="C6EFCE")
    RF = PatternFill("solid", start_color="FFC7CE")
    GN = Font(color="006100", bold=True, size=10)
    RN = Font(color="9C0006", bold=True, size=10)
    CA = Alignment(horizontal="center", vertical="center")

    for r in range(2, ws.max_row+1):
        raw_reg = ws.cell(r, rc).value
        if raw_reg is None: continue
        reg = re.sub(r'\.0+$', '', str(raw_reg).strip())
        if not reg: continue

        # CORRECT: only PRESENT if explicitly marked 1 in master
        att_val = day_att.get(reg)   # None if not in master
        if att_val == 1:
            status = "PRESENT"
        else:
            # att_val == 0 (marked absent) OR None (not in master) → ABSENT
            status = "ABSENT"

        cell = ws.cell(r, ac)
        cell.value     = status
        cell.fill      = GF if status == "PRESENT" else RF
        cell.font      = GN if status == "PRESENT" else RN
        cell.alignment = CA

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def build_zip(sess: bytes, dw: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("session_sheet.xlsx", sess)
        for fn,fb in dw.items(): zf.writestr(f"attendance/{fn}",fb)
    buf.seek(0); return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════
#                             MAIN UI
# ══════════════════════════════════════════════════════════════════════

def main():
    st.markdown("""
    <div class="banner">
      <h1>📋 DG Session & Attendance Sheet Generator</h1>
      <p>Upload files · Select dates · Configure · Generate exact DG format files</p>
      <small>Created by Dr. Amar Shukla · IILM University, Gurugram</small>
    </div>""", unsafe_allow_html=True)

    for k in ["att","syl","preview","results","_anm","_snm"]:
        if k not in st.session_state: st.session_state[k]=None

    att_data  = st.session_state.att
    syl_data  = st.session_state.syl or OrderedDict()
    all_dates = att_data["dates"] if att_data else []
    month_map = group_by_month(all_dates)
    syl_units = list(syl_data.keys())

    # Stats
    if all_dates or syl_data:
        n_tp = sum(len(v) for v in syl_data.values())
        st.markdown(f"""
        <div class="stat-row">
          <div class="sc"><div class="sn">{len(all_dates)}</div><div class="sl">Session Dates</div></div>
          <div class="sc"><div class="sn">{len(month_map)}</div><div class="sl">Months</div></div>
          <div class="sc"><div class="sn">{len(syl_units)}</div><div class="sl">Syllabus Units</div></div>
          <div class="sc"><div class="sn">{n_tp}</div><div class="sl">Topics Found</div></div>
        </div>""", unsafe_allow_html=True)

    st.divider()

    # ═══════════════════════════════════════════════════════════════
    # STEP 1 — FILES
    # ═══════════════════════════════════════════════════════════════
    sec("📂  Step 1 — Upload Files")
    c1,c2,c3 = st.columns(3)

    with c1:
        st.markdown("**📅 Master Attendance Sheet**")
        st.caption("Any Excel format — dates, enrollment IDs, and 0/1 attendance auto-detected")
        att_file = st.file_uploader("",type=["xlsx","xls"],key="u_att",label_visibility="collapsed")
        if att_file:
            st.markdown(f'<span class="badge">✅ {att_file.name}</span>',unsafe_allow_html=True)

    with c2:
        st.markdown("**📚 Syllabus File**")
        st.caption("Accepts .docx · .xlsx · .txt · .csv — units auto-detected")
        syl_file = st.file_uploader("",type=["docx","xlsx","xls","txt","csv"],key="u_syl",label_visibility="collapsed")
        if syl_file:
            st.markdown(f'<span class="badge">✅ {syl_file.name}</span>',unsafe_allow_html=True)

    with c3:
        st.markdown("**🗂️ DG Attendance Template**")
        st.caption("Section-specific template with student Email IDs and Registration IDs")
        att_tpl = st.file_uploader("",type=["xlsx"],key="u_atpl",label_visibility="collapsed")
        if att_tpl:
            st.markdown(f'<span class="badge">✅ {att_tpl.name}</span>',unsafe_allow_html=True)

    if att_file and att_file.name!=st.session_state._anm:
        try:
            st.session_state.att  = parse_attendance(att_file.read())
            st.session_state._anm = att_file.name
            st.session_state.preview=None; st.session_state.results=None
            att_data  = st.session_state.att
            all_dates = att_data["dates"]
            month_map = group_by_month(all_dates)
            ok(f"Attendance loaded — <b>{len(all_dates)} session dates</b>")
        except Exception as ex: er(f"Could not read attendance: {ex}")

    if syl_file and syl_file.name!=st.session_state._snm:
        try:
            st.session_state.syl  = parse_syllabus(syl_file.read(), syl_file.name)
            st.session_state._snm = syl_file.name
            syl_data  = st.session_state.syl
            syl_units = list(syl_data.keys())
            n_tp      = sum(len(v) for v in syl_data.values())
            ok(f"Syllabus loaded — <b>{len(syl_units)} unit(s)</b> · <b>{n_tp} topics</b>")
        except Exception as ex: wn(f"Syllabus warning: {ex}")

    st.divider()

    # ═══════════════════════════════════════════════════════════════
    # STEP 2 — SELECT DATES
    # ═══════════════════════════════════════════════════════════════
    sec("📅  Step 2 — Select Session Dates")

    selected_dates = []

    if not month_map:
        ib("Upload the Master Attendance Sheet — months and dates will appear here.")
    else:
        ib("Select months to include. Expand any month to pick individual dates.")

        for row_start in range(0, len(month_map), 4):
            keys_row = list(month_map.keys())[row_start:row_start+4]
            cols     = st.columns(len(keys_row))
            for ci, ym in enumerate(keys_row):
                ds    = month_map[ym]
                label = datetime(ym[0],ym[1],1).strftime("%B %Y")
                with cols[ci]:
                    mon_on = st.checkbox(f"**{label}**", value=True, key=f"m_{ym[0]}_{ym[1]}")
                    st.caption(f"🗓️ {len(ds)} session dates")
                    if mon_on:
                        with st.expander(f"Select dates in {label}", expanded=False):
                            ib("Uncheck any date to exclude it.")
                            for d in ds:
                                lbl = f"{d.strftime('%d %b %Y')}  ({d.strftime('%A')})"
                                if st.checkbox(lbl, value=True, key=f"d_{d.isoformat()}"):
                                    selected_dates.append(d)
                        chips = "".join(
                            f'<span class="date-chip">{d.strftime("%d")} {d.strftime("%a")}</span>'
                            for d in ds
                        )
                        st.markdown(chips, unsafe_allow_html=True)

        selected_dates = sorted(set(selected_dates))
        if selected_dates:
            ok(f"<b>{len(selected_dates)} dates selected</b> — "
               f"{selected_dates[0].strftime('%d %b %Y')} to "
               f"{selected_dates[-1].strftime('%d %b %Y')}")
        else:
            wn("Select at least one date.")

    st.divider()

    # ═══════════════════════════════════════════════════════════════
    # STEP 3 — BASIC SETTINGS
    # ═══════════════════════════════════════════════════════════════
    sec("⚙️  Step 3 — Basic Settings")

    s1,s2,s3,s4,s5 = st.columns([1.5,1.5,1,1,1])

    with s1:
        st.markdown("**📊 Total Sessions**")
        st.caption(
            "Default = number of selected dates. "
            "Check the box only if you want a different number — "
            "e.g. 30 dates selected but only 25 sessions needed."
        )
        use_custom = st.checkbox("Enter custom session count", value=False)
        if use_custom:
            total_sessions = st.number_input("Sessions:", min_value=1, max_value=500,
                                              value=len(selected_dates) or 50)
        else:
            total_sessions = len(selected_dates)
            st.metric("Sessions (= selected dates)", total_sessions)

    with s2:
        faculty_id = st.text_input(
            "🧑‍🏫 Faculty Registration ID",
            placeholder="e.g. IILMGG006412025"
        )

    with s3:
        mandatory = st.selectbox("✅ Attendance Mandatory?", ["TRUE","FALSE"])

    with s4:
        tlo_max = st.number_input("Max TLO Number", min_value=1, max_value=100, value=5)

    with s5:
        tlo_prefix = st.text_input("TLO Prefix", value="TLO",
                                   help="'TLO' → TLO1,TLO2 | 'CO' → CO1,CO2")

    all_tlos = [f"{tlo_prefix}{i}" for i in range(1, tlo_max+1)]

    st.divider()

    # ═══════════════════════════════════════════════════════════════
    # STEP 4 — TIMING  (simple: one row, text inputs)
    # ═══════════════════════════════════════════════════════════════
    sec("⏰  Step 4 — Lecture Timing")
    ib("Enter start and end time for each day that has classes. "
       "Format: <b>HH:MM</b> or <b>10:00 AM</b> — both accepted. "
       "Days with classes in your selected dates are pre-filled.")

    # Detect which days have sessions
    detected = {}
    for d in selected_dates:
        dn = d.strftime("%A"); detected[dn] = detected.get(dn,0)+1

    day_timing: dict[str, tuple] = {}   # {day: (h, m, h_end, m_end)}

    # Show as a clean table: day | sessions | start | end
    col_hdr = st.columns([2,1,2,2])
    col_hdr[0].markdown("**Day**"); col_hdr[1].markdown("**Sessions**")
    col_hdr[2].markdown("**Start Time**"); col_hdr[3].markdown("**End Time**")

    for day in ALL_DAYS:
        cnt = detected.get(day, 0)
        cols = st.columns([2,1,2,2])
        with cols[0]:
            enabled = st.checkbox(f"{day}", value=(cnt>0), key=f"dc_{day}")
        with cols[1]:
            st.caption(f"🔵 {cnt}" if cnt else "—")
        if enabled:
            with cols[2]:
                st_raw = st.text_input("", value="10:00", key=f"st_{day}",
                                        label_visibility="collapsed",
                                        placeholder="10:00")
            with cols[3]:
                en_raw = st.text_input("", value="11:00", key=f"en_{day}",
                                        label_visibility="collapsed",
                                        placeholder="11:00")
            sh, sm = parse_time(st_raw)
            eh, em = parse_time(en_raw)
            day_timing[day] = (sh, sm, eh, em)

    st.divider()

    # ═══════════════════════════════════════════════════════════════
    # STEP 5 — UNITS & MODULES
    # ═══════════════════════════════════════════════════════════════
    sec("📚  Step 5 — Units to Include & Module Configuration")

    unit_configs = []

    if not syl_units:
        wn("No syllabus uploaded. Titles will be 'Session 1', 'Session 2'... "
           "Upload a syllabus to auto-fill topics and descriptions.")
    else:
        ib(f"<b>{len(syl_units)} unit(s)</b> found in syllabus. "
           "Choose how many to include. For each: set Module Name, TLOs, and topic coverage.")

        n_include = st.number_input(
            f"How many units to include?  (Syllabus has {len(syl_units)})",
            min_value=1, max_value=len(syl_units), value=len(syl_units),
            help="Example: syllabus has 4 units but you want only 2 in this sheet → type 2"
        )

        for i, unit_key in enumerate(syl_units[:n_include]):
            avail = syl_data[unit_key]
            with st.expander(
                f"Unit {i+1}  ·  {unit_key[:65]}  ({len(avail)} topics)",
                expanded=(i<3)
            ):
                r1,r2 = st.columns([3,2])
                with r1:
                    mod_name = st.text_input(
                        "Module Name*  (appears in 'Module Name*' column of session sheet)",
                        value=unit_key, key=f"mn_{i}"
                    )
                with r2:
                    def_tlo  = [all_tlos[i%len(all_tlos)]] if all_tlos else [f"{tlo_prefix}1"]
                    sel_tlos = st.multiselect("TLOs", all_tlos, default=def_tlo, key=f"tl_{i}",
                                              help="Multiple TLOs → 'TLO1 | TLO2' in output")
                    tlo_str  = " | ".join(sel_tlos) if sel_tlos else all_tlos[i%len(all_tlos)]

                cov = st.radio("Topic coverage:",
                               ["📖 Full Unit — all topics",
                                "✂️ Custom — select specific topics"],
                               key=f"cov_{i}", horizontal=True)
                if "Custom" in cov:
                    chosen = st.multiselect("Select topics:", avail, default=avail, key=f"tp_{i}")
                    final_topics = chosen
                    st.caption(f"✅ {len(final_topics)} / {len(avail)} topics selected")
                else:
                    final_topics = avail
                    with st.expander(f"Preview {len(avail)} topics →"):
                        for j,t in enumerate(avail,1): st.write(f"`{j}.` {t}")

                unit_configs.append({"module_name":mod_name,"tlo":tlo_str,"topics":final_topics})

        n_tp = sum(len(u["topics"]) for u in unit_configs)
        if total_sessions and n_tp:
            if n_tp == total_sessions: ok("Perfect — one topic per session.")
            elif n_tp < total_sessions:
                wn(f"Topics ({n_tp}) < Sessions ({total_sessions}) — "
                   f"last {total_sessions-n_tp} rows will be revision entries.")
            else:
                wn(f"Topics ({n_tp}) > Sessions ({total_sessions}) — "
                   f"trimmed proportionally per unit.")

    st.divider()

    # ═══════════════════════════════════════════════════════════════
    # STEP 6 — PREVIEW
    # ═══════════════════════════════════════════════════════════════
    sec("👀  Step 6 — Preview & Edit")
    ib("Click <b>Build Preview</b> → all rows generated with balanced topics "
       "and auto-descriptions. Edit any cell before generating files.")

    if st.button("🔄  Build Preview Table", type="secondary", use_container_width=True):
        if not selected_dates:
            er("Select at least one date in Step 2.")
        elif not faculty_id.strip():
            er("Enter Faculty Registration ID in Step 3.")
        else:
            flat = (balance_topics(unit_configs, total_sessions)
                    if unit_configs else
                    [{"module":"Module","tlo":f"{tlo_prefix}1","title":f"Session {i+1}"}
                     for i in range(total_sessions)])

            rows = []
            use_dates = selected_dates[:total_sessions]
            for idx in range(total_sessions):
                if idx < len(use_dates):
                    d  = use_dates[idx]
                    dn = d.strftime("%A")
                    if dn in day_timing:
                        sh,sm,eh,em = day_timing[dn]
                    else:
                        sh,sm,eh,em = 10,0,11,0
                    sdt = fmt_dt(d,sh,sm)
                    edt = fmt_dt(d,eh,em)
                else:
                    sdt = edt = "TBD"

                sess = flat[idx] if idx<len(flat) else {
                    "module":"Extra","tlo":f"{tlo_prefix}1",
                    "title":f"Extra Session {idx+1}"}

                rows.append({
                    "Sr":               idx+1,
                    "Module Name*":     sess["module"],
                    "Start Date Time*": sdt,
                    "End Date Time*":   edt,
                    "Title*":           sess["title"],
                    "Description*":     auto_desc(sess["title"],sess["module"]),
                    "Mandatory*":       mandatory,
                    "TLO":              sess["tlo"],
                    "Faculty Reg ID*":  faculty_id.strip(),
                })

            st.session_state.preview = pd.DataFrame(rows)
            st.session_state.results = None
            ok(f"Preview ready — <b>{len(rows)} rows</b>. Edit below if needed.")

    if st.session_state.preview is not None:
        edited = st.data_editor(
            st.session_state.preview,
            use_container_width=True, num_rows="fixed",
            hide_index=True, key="ed_prev",
            column_config={
                "Sr":               st.column_config.NumberColumn("Sr",width=55,disabled=True),
                "Module Name*":     st.column_config.TextColumn("Module Name* ✏️",    width=230),
                "Start Date Time*": st.column_config.TextColumn("Start DateTime* ✏️", width=170),
                "End Date Time*":   st.column_config.TextColumn("End DateTime* ✏️",   width=170),
                "Title*":           st.column_config.TextColumn("Title* ✏️",          width=265),
                "Description*":     st.column_config.TextColumn("Description* ✏️",    width=380),
                "Mandatory*":       st.column_config.SelectboxColumn("Mandatory*",
                                        options=["TRUE","FALSE"],width=110),
                "TLO":              st.column_config.TextColumn("TLO ✏️",             width=130),
                "Faculty Reg ID*":  st.column_config.TextColumn("Faculty Reg ID* ✏️", width=170),
            },
            height=480,
        )
        st.session_state.preview = edited

    st.divider()

    # ═══════════════════════════════════════════════════════════════
    # STEP 7 — GENERATE
    # ═══════════════════════════════════════════════════════════════
    sec("🚀  Step 7 — Generate Files")

    can_gen = bool(st.session_state.preview is not None and att_data and faculty_id.strip())
    if not can_gen:
        miss=[]
        if not att_data:                       miss.append("Attendance Sheet (Step 1)")
        if not faculty_id.strip():             miss.append("Faculty Registration ID (Step 3)")
        if st.session_state.preview is None:  miss.append("Preview Table (Step 6)")
        if miss: ib(f"Complete first: <b>{' · '.join(miss)}</b>")

    if st.button("⚡  Generate Session Sheet + Day-wise Attendance Files",
                 type="primary", disabled=not can_gen, use_container_width=True):

        errors=[]; prog=st.progress(0); log=st.empty()
        try:
            df = st.session_state.preview
            session_rows=[{
                "module":      str(r.get("Module Name*","")),
                "start_dt":    str(r.get("Start Date Time*","")),
                "end_dt":      str(r.get("End Date Time*","")),
                "title":       str(r.get("Title*","")),
                "description": str(r.get("Description*","")),
                "mandatory":   str(r.get("Mandatory*","TRUE")),
                "tlo":         str(r.get("TLO",f"{tlo_prefix}1")),
                "faculty_id":  str(r.get("Faculty Reg ID*",faculty_id.strip())),
            } for _,r in df.iterrows()]
            prog.progress(20)

            log.markdown('<div class="ib">📊 Generating session sheet…</div>',unsafe_allow_html=True)
            session_xlsx = gen_session_sheet(session_rows)
            ok(f"Session sheet — <b>{len(session_rows)} rows</b> in exact DL.xlsx format ✅")
            prog.progress(50)

            daywise={}
            if att_tpl:
                log.markdown('<div class="ib">🗂️ Generating day-wise attendance…</div>',
                             unsafe_allow_html=True)
                tpl_b = att_tpl.read()
                bar   = st.progress(0)
                use_d = selected_dates[:total_sessions]
                for i,d in enumerate(use_d):
                    try:
                        fn = f"attendance_{d.strftime('%Y-%m-%d')}.xlsx"
                        daywise[fn] = gen_daywise_att(tpl_b, att_data, d)
                    except Exception as ex: errors.append(f"{d}: {ex}")
                    bar.progress((i+1)/max(len(use_d),1))

                # Count sample stats from first file
                if daywise:
                    first_fn = list(daywise.keys())[0]
                    from openpyxl import load_workbook as lw2
                    ws_check = lw2(io.BytesIO(list(daywise.values())[0])).active
                    p_cnt=sum(1 for r in range(2,ws_check.max_row+1) if ws_check.cell(r,3).value=="PRESENT")
                    a_cnt=sum(1 for r in range(2,ws_check.max_row+1) if ws_check.cell(r,3).value=="ABSENT")
                    ok(f"Day-wise attendance — <b>{len(daywise)} files</b>. "
                       f"Sample ({list(use_d)[0].strftime('%d %b %Y')}): "
                       f"PRESENT={p_cnt}, ABSENT={a_cnt}")
            else:
                wn("No DG Attendance Template uploaded — only session sheet generated.")

            prog.progress(90)
            zip_b = build_zip(session_xlsx, daywise)
            prog.progress(100)
            log.markdown('<div class="ok">🎉 Done! Download files below.</div>',
                         unsafe_allow_html=True)

            st.session_state.results={
                "session":session_xlsx,"daywise":daywise,
                "zip":zip_b,"errors":errors,"n":len(session_rows)
            }
        except Exception as ex:
            er(f"Error: {ex}")
            import traceback; st.code(traceback.format_exc())

    # ═══════════════════════════════════════════════════════════════
    # STEP 8 — DOWNLOADS
    # ═══════════════════════════════════════════════════════════════
    if st.session_state.results:
        res=st.session_state.results
        st.divider()
        sec("📥  Step 8 — Download")
        for e in res.get("errors",[]): wn(f"Skipped: {e}")

        d1,d2=st.columns(2)
        with d1:
            st.download_button("📦  Download Everything as ZIP",
                data=res["zip"],
                file_name=f"DG_Output_{datetime.now():%Y%m%d_%H%M}.zip",
                mime="application/zip",use_container_width=True)
        with d2:
            st.download_button("📋  Download session_sheet.xlsx",
                data=res["session"],file_name="session_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        if res["daywise"]:
            st.markdown(f"#### 📅  Day-wise Attendance ({len(res['daywise'])} files)")
            files=list(res["daywise"].items())
            for ri in range(0,len(files),4):
                chunk=files[ri:ri+4]; cols=st.columns(4)
                for ci,(fn,fb) in enumerate(chunk):
                    with cols[ci]:
                        dp=fn.replace("attendance_","").replace(".xlsx","")
                        try:    lbl=datetime.strptime(dp,"%Y-%m-%d").strftime("%d %b %Y")
                        except: lbl=dp
                        st.download_button(f"📅  {lbl}",data=fb,file_name=fn,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,key=f"dl_{fn}")

        st.markdown("<br>",unsafe_allow_html=True)
        if st.button("🔄  Reset — Start Over"):
            for k in ["att","syl","preview","results","_anm","_snm"]:
                st.session_state[k]=None
            st.rerun()

    st.markdown("""
    <div class="footer">
      Created by <b>Dr. Amar Shukla</b> · IILM University, Gurugram ·
      All processing in-memory — no data stored on server
    </div>""", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
