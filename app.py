"""
DG Sheet Generator  v9.0
Created by Dr. Amar Shukla
==========================
• Month → individual date selection
• All English UI
• Large clear fonts
• Full/Custom topic selection per unit
• Dynamic units from any syllabus format
• Module Name separate field
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

# ──────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DG Sheet Generator — Dr. Amar Shukla",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
/* === Base === */
html, body, [class*="css"] {
    font-size: 16px !important;
    font-family: 'Segoe UI', Arial, sans-serif !important;
}
.main { padding: 0.6rem 2rem 3rem; }

/* === Banner === */
.banner {
    background: linear-gradient(135deg, #0a1f35 0%, #1565c0 100%);
    color: #fff; padding: 1.6rem 2.5rem; border-radius: 16px;
    margin-bottom: 1.4rem; position: relative; overflow: hidden;
}
.banner::after {
    content: ""; position: absolute; top: -40px; right: -40px;
    width: 200px; height: 200px; border-radius: 50%;
    background: rgba(255,255,255,.06);
}
.banner h1 { margin: 0; font-size: 2.1rem; font-weight: 900; letter-spacing: -.4px; }
.banner .sub { margin: .4rem 0 0; font-size: 1.05rem; opacity: .88; }
.banner .credit {
    margin: .6rem 0 0; font-size: .88rem; opacity: .7;
    font-style: italic; letter-spacing: .3px;
}

/* === Section Header === */
.sec {
    font-size: 1.2rem; font-weight: 800; color: #0a1f35;
    border-left: 6px solid #1565c0; padding: .3rem 0 .3rem .75rem;
    margin: 1.3rem 0 .7rem; letter-spacing: -.2px;
}

/* === Alerts === */
.al-info { background:#e3f2fd; border-left:5px solid #1565c0; border-radius:0 8px 8px 0;
           padding:.65rem 1.1rem; margin:.35rem 0; font-size:.95rem; color:#0d47a1; }
.al-ok   { background:#e8f5e9; border-left:5px solid #2e7d32; border-radius:0 8px 8px 0;
           padding:.65rem 1.1rem; margin:.35rem 0; font-size:.95rem; color:#1b5e20; }
.al-warn { background:#fff8e1; border-left:5px solid #f57f17; border-radius:0 8px 8px 0;
           padding:.65rem 1.1rem; margin:.35rem 0; font-size:.95rem; color:#e65100; }
.al-err  { background:#ffebee; border-left:5px solid #c62828; border-radius:0 8px 8px 0;
           padding:.65rem 1.1rem; margin:.35rem 0; font-size:.95rem; color:#b71c1c; }

/* === Stat Cards === */
.stat-row { display:flex; gap:1rem; margin: .8rem 0; flex-wrap:wrap; }
.stat-card {
    background:#fff; border:2px solid #e3f2fd; border-radius:12px;
    padding:1rem 1.4rem; text-align:center; min-width:130px; flex:1;
    box-shadow: 0 2px 8px rgba(21,101,192,.08);
}
.stat-num { font-size:2.3rem; font-weight:900; color:#1565c0; line-height:1.1; }
.stat-lbl { font-size:.85rem; color:#546e7a; margin-top:.3rem; font-weight:600; }

/* === Month Card === */
.month-grid { display:flex; flex-wrap:wrap; gap:.9rem; margin:.6rem 0; }
.month-box {
    background:#f5f9ff; border:2px solid #bbdefb; border-radius:12px;
    padding:.8rem; min-width:180px; flex:1;
}
.month-box.active { border-color:#1565c0; background:#e3f2fd; }
.month-title { font-size:1.05rem; font-weight:800; color:#0a1f35; margin-bottom:.3rem; }
.month-count { font-size:.82rem; color:#1565c0; font-weight:700; margin-bottom:.4rem; }
.date-chip {
    display:inline-block; background:#1565c0; color:#fff;
    border-radius:14px; padding:.15rem .55rem; margin:.12rem;
    font-size:.78rem; font-weight:700; letter-spacing:.2px;
}
.date-chip.selected { background:#1565c0; }
.date-chip.deselected { background:#90a4ae; }

/* === Day Timing Card === */
.day-grid { display:flex; gap:.6rem; flex-wrap:wrap; margin:.5rem 0; }
.day-card {
    background:#f5f9ff; border:2px solid #bbdefb; border-radius:10px;
    padding:.7rem .6rem; text-align:center; min-width:100px; flex:1;
}
.day-card.on { border-color:#1565c0; background:#e8f4ff; }
.day-card.off { opacity:.5; }
.day-name { font-size:1rem; font-weight:800; color:#0a1f35; }
.day-sess { font-size:.78rem; color:#1565c0; font-weight:600; margin:.15rem 0; }
.time-tag { font-size:.82rem; font-weight:700; background:#1565c0; color:#fff;
            border-radius:6px; padding:.2rem .5rem; margin-top:.3rem; display:inline-block; }

/* === Unit Card === */
.unit-hdr {
    display:flex; align-items:center; gap:.6rem;
    font-size:1rem; font-weight:700; color:#0a1f35;
}
.unit-badge {
    background:#1565c0; color:#fff; border-radius:50%;
    width:30px; height:30px; display:inline-flex;
    align-items:center; justify-content:center;
    font-size:.9rem; font-weight:900; flex-shrink:0;
}

/* === Badge === */
.badge { display:inline-block; background:#e8f5e9; color:#2e7d32;
         padding:.15rem .65rem; border-radius:12px; font-size:.8rem; font-weight:700; }

/* === Buttons === */
.stButton>button {
    font-size: 1.05rem !important; font-weight: 700 !important;
    padding: .6rem 1.5rem !important; border-radius: 10px !important;
    letter-spacing: .2px !important;
}
.stDownloadButton>button {
    background: #0a1f35 !important; color: #fff !important;
    border-radius: 10px !important; font-weight: 700 !important;
    font-size: 1rem !important; width: 100% !important;
    padding: .55rem 1rem !important;
}
.stDownloadButton>button:hover { background: #1565c0 !important; }

/* === Form labels === */
label { font-size: 1rem !important; font-weight: 600 !important; color: #1a2e3b !important; }
.stTextInput>div>div>input,
.stNumberInput>div>div>input,
.stSelectbox>div>div>div {
    font-size: 1rem !important;
}
/* expander header */
.streamlit-expanderHeader { font-size: 1rem !important; font-weight: 700 !important; }

/* === Divider === */
hr { border: none; border-top: 2.5px solid #e3f2fd; margin: 1.2rem 0; }

/* === Footer === */
.footer { text-align:center; color:#90a4ae; font-size:.82rem; margin-top:2rem;
          padding-top:1rem; border-top:2px solid #e3f2fd; }
</style>
""", unsafe_allow_html=True)

def al_info(m): st.markdown(f'<div class="al-info">ℹ️  {m}</div>', unsafe_allow_html=True)
def al_ok(m):   st.markdown(f'<div class="al-ok">✅  {m}</div>',   unsafe_allow_html=True)
def al_warn(m): st.markdown(f'<div class="al-warn">⚠️  {m}</div>', unsafe_allow_html=True)
def al_err(m):  st.markdown(f'<div class="al-err">❌  {m}</div>',  unsafe_allow_html=True)
def sec(t):     st.markdown(f'<div class="sec">{t}</div>',          unsafe_allow_html=True)

ALL_DAYS  = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
SHORT_DAY = {d: d[:3] for d in ALL_DAYS}

# ─────────────────────────────────── UTILITIES ────────────────────────

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
        try: return datetime.strptime(s, fmt).date()
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

def best_sheet(fb: bytes, hints: list = None) -> pd.DataFrame:
    xf = pd.ExcelFile(io.BytesIO(fb)); ns = xf.sheet_names
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

# ─────────────────────────────── ATTENDANCE PARSER ────────────────────

def parse_attendance(fb: bytes) -> dict:
    df = best_sheet(fb, hints=["attendance","student"])
    best_row_idx, best_count, best_first = None, 0, None
    for i, row in df.iterrows():
        hits, first = [], None
        for j, v in enumerate(row):
            d = parse_date(v)
            if d: hits.append(d); first = (j if first is None else first)
        if len(hits) > best_count:
            best_count, best_row_idx, best_first = len(hits), i, first
    if best_row_idx is None or best_count < 2:
        raise ValueError("Could not find session dates. Ensure dates appear in a single row.")

    date_row = df.iloc[best_row_idx]
    col_to_date, dates_ordered, seen = {}, [], set()
    for j in range(best_first, df.shape[1]):
        d = parse_date(date_row.iloc[j])
        if d:
            col_to_date[j] = d
            if d not in seen: dates_ordered.append(d); seen.add(d)

    name_col = enroll_col = None
    for i in range(best_row_idx):
        row = df.iloc[i]
        rs  = " ".join(str(v).lower() for v in row if not safe_isna(v))
        if any(k in rs for k in ["name","enrol","roll","student"]):
            for j, v in enumerate(row):
                s = str(v).lower().strip()
                if "name" in s and name_col is None:                               name_col   = j
                elif any(k in s for k in ["enrol","roll"]) and enroll_col is None: enroll_col = j
            break

    students = []
    for i in range(best_row_idx+1, df.shape[0]):
        row = df.iloc[i]
        if row.notna().sum() < 3: continue
        name   = (str(row.iloc[name_col]).strip()
                  if name_col   is not None and not safe_isna(row.iloc[name_col])   else "")
        enroll = (str(row.iloc[enroll_col]).strip()
                  if enroll_col is not None and not safe_isna(row.iloc[enroll_col]) else "")
        if not name or name.lower() in ("nan","none",""): continue
        att = {}
        for col_j, d in col_to_date.items():
            val = row.iloc[col_j]
            try:    att[d] = int(float(val)) if not safe_isna(val) else None
            except: att[d] = None
        students.append({"name":name, "enrollment":enroll, "att":att})

    return {"dates": dates_ordered, "students": students}

def group_by_month(dates: list) -> OrderedDict:
    res = defaultdict(list)
    for d in dates:
        if d: res[(d.year, d.month)].append(d)
    return OrderedDict(sorted(res.items()))

# ──────────────────────────────── SYLLABUS PARSER ─────────────────────

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
        if not t or len(t) < 4 or SKIP_RE.search(t): return
        result.setdefault(unit, []).append(t)

    if ext == "docx":
        import docx as _d
        doc = _d.Document(io.BytesIO(fb))
        cur = None
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
        df = best_sheet(fb)
        cur = None
        for _, row in df.iterrows():
            for val in row:
                s = str(val).strip()
                if not s or s.lower() in ("nan","none",""): continue
                if UNIT_RE.match(s): cur = s; result[cur] = []; break
                elif cur and len(s) > 4: add(cur, s); break

    elif ext in ("txt","csv"):
        cur = None
        for line in fb.decode("utf-8","ignore").splitlines():
            t = line.strip()
            if not t: continue
            if UNIT_RE.match(t): cur = t; result[cur] = []
            elif cur: add(cur, t)

    # Fallback: whole file as one module
    if not result:
        result["Module 1"] = []
        if ext in ("txt","csv"):
            for line in fb.decode("utf-8","ignore").splitlines():
                add("Module 1", line)
        elif ext == "docx":
            import docx as _d
            for p in _d.Document(io.BytesIO(fb)).paragraphs: add("Module 1", p.text)
        elif ext in ("xlsx","xls"):
            df = best_sheet(fb)
            for _, row in df.iterrows():
                for val in row:
                    s = str(val).strip()
                    if s and s.lower() not in ("nan","none","") and len(s)>4:
                        add("Module 1", s); break

    return OrderedDict((k,v) for k,v in result.items() if v)

# ─────────────────────────────── AUTO-DESCRIPTION ─────────────────────

def auto_desc(title: str, module: str) -> str:
    t   = title.strip()
    mod = re.sub(r"^(UNIT|MODULE|CHAPTER)\s*\d+[:\-.]?\s*","",module,flags=re.IGNORECASE).strip() or module
    tl  = t.lower()
    if any(w in tl for w in ["definition","define","meaning"]):
        core = re.sub(r"^(definition\s+(of\s+)?|define\s+)","",t,flags=re.IGNORECASE).strip()
        return (f"This session covers the statutory definition and essential elements of {core}. "
                f"Students will examine relevant legal provisions, judicial interpretations, and practical significance within {mod}.")
    if any(w in tl for w in ["rights","duties","liability","liabilities","obligation"]):
        return (f"This session examines the rights, duties, and liabilities under {t}. "
                f"Students will analyse statutory provisions, landmark judgments, and legal consequences for parties in {mod}.")
    if any(w in tl for w in ["distinction","difference","compare","versus"]):
        return (f"This session provides a comparative analysis of {t}. "
                f"Students will identify key distinctions through statutory provisions and case-law to apply differential reasoning in {mod}.")
    if any(w in tl for w in ["type","kind","classif","categor","form","nature"]):
        return (f"This session explores the types, categories, and conceptual framework of {t}. "
                f"Students will study the legal significance of each category and its application within {mod}.")
    if any(w in tl for w in ["termination","discharge","revocation","dissolution"]):
        return (f"This session covers the modes of {t} and the legal consequences that follow. "
                f"Students will study statutory provisions, conditions, and judicial precedents governing this aspect of {mod}.")
    if any(w in tl for w in ["creation","formation","essential","element"]):
        return (f"This session discusses the formation process and requisite legal elements of {t}. "
                f"Students will examine statutory requirements, judicial interpretations, and practical illustrations in {mod}.")
    if any(w in tl for w in ["remedy","remedies","damages","compensation"]):
        return (f"This session discusses remedies available in {t}. "
                f"Students will examine statutory and equitable remedies, judicial approaches, and computation of relief in {mod}.")
    if any(w in tl for w in ["case","judgment","ruling"," v."]):
        return (f"This session analyses {t} as a significant judicial decision. "
                f"Students will examine the facts, legal issues, court reasoning, and precedential value of this ruling in {mod}.")
    return (f"This session provides a comprehensive study of {t} within {mod}. "
            f"Students will analyse relevant statutory provisions, judicial precedents, and practical applications through structured discussion.")

# ────────────────────────────── TOPIC BALANCER ────────────────────────

def balance_topics(unit_configs: list, total_sessions: int) -> list:
    """unit_configs: [{"module_name","tlo","topics":[]}]  →  list[{"module","tlo","title"}]"""
    all_items = []
    for uc in unit_configs:
        for t in uc["topics"]:
            all_items.append({"module": uc["module_name"], "tlo": uc["tlo"], "title": t})
    n = len(all_items)
    if n == 0:
        return [{"module":"Session","tlo":"TLO1","title":f"Session {i+1}"}
                for i in range(total_sessions)]

    result = []
    if n <= total_sessions:
        result = list(all_items)
        for i in range(total_sessions - n):
            base = all_items[i % n]
            result.append({"module":base["module"],"tlo":base["tlo"],
                           "title":base["title"]+" — Revision"})
    else:
        total_avail = sum(len(uc["topics"]) for uc in unit_configs)
        for uc in unit_configs:
            n_take = max(1, round(total_sessions * len(uc["topics"]) / total_avail))
            n_take = min(n_take, len(uc["topics"]))
            for t in uc["topics"][:n_take]:
                result.append({"module":uc["module_name"],"tlo":uc["tlo"],"title":t})
        while len(result) < total_sessions:
            result.append(result[-1] | {"title": result[-1]["title"]+" (Cont.)"})
        result = result[:total_sessions]
    return result

# ────────────────────────────── EXCEL GENERATORS ──────────────────────

HEADERS = ["Module Name*","Start Date Time*","End Date Time*","Title*",
           "Description*","Attendance Mandatory*","TLO","Teaching Faculty Registration ID*"]
WIDTHS  = [30,22,22,42,60,22,18,32]
HF=PatternFill("solid",start_color="1F4E79"); HT=Font(bold=True,color="FFFFFF",size=11)
DA=PatternFill("solid",start_color="D6EAF8"); DB=PatternFill("solid",start_color="EBF5FB")
TN=Side(style="thin",color="BBBBBB")
BD=Border(left=TN,right=TN,top=TN,bottom=TN)
CT=Alignment(horizontal="center",vertical="center",wrap_text=True)
LF=Alignment(horizontal="left",  vertical="center",wrap_text=True)

def gen_session_sheet(rows: list) -> bytes:
    wb = Workbook(); ws = wb.active; ws.title="Session Sheet"
    aligns = [LF,CT,CT,LF,LF,CT,CT,CT]
    for c,(h,w) in enumerate(zip(HEADERS,WIDTHS),1):
        cell=ws.cell(1,c,h); cell.fill=HF; cell.font=HT
        cell.border=BD; cell.alignment=CT
        ws.column_dimensions[get_column_letter(c)].width=w
    ws.row_dimensions[1].height=24
    for i,row in enumerate(rows,2):
        fill=DA if i%2==0 else DB
        vals=[row.get(k,"") for k in
              ["module","start_dt","end_dt","title","description","mandatory","tlo","faculty_id"]]
        for c,(v,al) in enumerate(zip(vals,aligns),1):
            cell=ws.cell(i,c,v); cell.fill=fill
            cell.border=BD; cell.alignment=al; cell.font=Font(size=10)
        ws.row_dimensions[i].height=30
    ws.freeze_panes="A2"
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

def gen_daywise_att(tpl: bytes, att_data: dict, d: date) -> bytes:
    wb=load_workbook(io.BytesIO(tpl)); ws=wb.active
    ec=rc=ac=None
    for c in range(1,ws.max_column+1):
        hv=str(ws.cell(1,c).value or "").lower().strip()
        if "email" in hv:                                       ec=c
        elif "registration" in hv or ("reg" in hv and "id" in hv): rc=c
        elif "attendance" in hv:                                ac=c
    if not all([ec,rc,ac]): ec,rc,ac=1,2,3
    ea={str(s["enrollment"]).strip():(1 if s["att"].get(d)==1 else 0)
        for s in att_data["students"]}
    GF=PatternFill("solid",start_color="C6EFCE"); RF=PatternFill("solid",start_color="FFC7CE")
    GT=Font(color="006100",bold=True,size=10);    RT=Font(color="9C0006",bold=True,size=10)
    CA=Alignment(horizontal="center",vertical="center")
    for r in range(2,ws.max_row+1):
        reg=str(ws.cell(r,rc).value or "").strip()
        if not reg: continue
        s="PRESENT" if ea.get(reg,0)==1 else "ABSENT"
        c=ws.cell(r,ac); c.value=s
        c.fill=GF if s=="PRESENT" else RF; c.font=GT if s=="PRESENT" else RT; c.alignment=CA
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

def build_zip(sess: bytes, dw: dict) -> bytes:
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("session_sheet.xlsx",sess)
        for fn,fb in dw.items(): zf.writestr(f"attendance/{fn}",fb)
    buf.seek(0); return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════
#                              MAIN  UI
# ══════════════════════════════════════════════════════════════════════

def main():
    # ── Banner ─────────────────────────────────────────────────────────
    st.markdown("""
    <div class="banner">
      <h1>📋 DG Session & Attendance Sheet Generator</h1>
      <div class="sub">Upload files · Select dates · Configure units · Generate — any format, any syllabus</div>
      <div class="credit">Created by Dr. Amar Shukla · IILM University, Gurugram</div>
    </div>""", unsafe_allow_html=True)

    for k in ["att","syl","preview","results","_anm","_snm"]:
        if k not in st.session_state: st.session_state[k]=None

    att_data  = st.session_state.att
    syl_data  = st.session_state.syl or OrderedDict()
    all_dates = att_data["dates"] if att_data else []
    month_map = group_by_month(all_dates)
    syl_units = list(syl_data.keys())

    # ── Quick stats ─────────────────────────────────────────────────────
    if all_dates or syl_data:
        n_st = len(att_data["students"]) if att_data else 0
        n_tp = sum(len(v) for v in syl_data.values())
        st.markdown(f"""
        <div class="stat-row">
          <div class="stat-card"><div class="stat-num">{len(all_dates)}</div><div class="stat-lbl">Session Dates</div></div>
          <div class="stat-card"><div class="stat-num">{len(month_map)}</div><div class="stat-lbl">Months Available</div></div>
          <div class="stat-card"><div class="stat-num">{n_st}</div><div class="stat-lbl">Students</div></div>
          <div class="stat-card"><div class="stat-num">{len(syl_units)}</div><div class="stat-lbl">Units in Syllabus</div></div>
          <div class="stat-card"><div class="stat-num">{n_tp}</div><div class="stat-lbl">Topics Found</div></div>
        </div>""", unsafe_allow_html=True)

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 1 — FILE UPLOADS
    # ══════════════════════════════════════════════════════════════════
    sec("📂  Step 1 — Upload Files")
    c1,c2,c3 = st.columns(3)

    with c1:
        st.markdown("**📅 Master Attendance Sheet**")
        st.caption("Supports any Excel format — dates auto-detected")
        att_file = st.file_uploader("",type=["xlsx","xls"],key="u_att",
                                    label_visibility="collapsed")
        if att_file:
            st.markdown(f'<span class="badge">✅ {att_file.name}</span>',
                        unsafe_allow_html=True)

    with c2:
        st.markdown("**📚 Syllabus File**")
        st.caption("Supports .docx · .xlsx · .txt · .csv — units auto-detected")
        syl_file = st.file_uploader("",type=["docx","xlsx","xls","txt","csv"],key="u_syl",
                                    label_visibility="collapsed")
        if syl_file:
            st.markdown(f'<span class="badge">✅ {syl_file.name}</span>',
                        unsafe_allow_html=True)

    with c3:
        st.markdown("**🗂️ DG Attendance Template**")
        st.caption("The pre-filled template with student Email & Registration IDs")
        att_tpl = st.file_uploader("",type=["xlsx"],key="u_atpl",
                                   label_visibility="collapsed")
        if att_tpl:
            st.markdown(f'<span class="badge">✅ {att_tpl.name}</span>',
                        unsafe_allow_html=True)

    # Auto-parse
    if att_file and att_file.name != st.session_state._anm:
        try:
            st.session_state.att  = parse_attendance(att_file.read())
            st.session_state._anm = att_file.name
            st.session_state.preview = None
            st.session_state.results = None
            att_data  = st.session_state.att
            all_dates = att_data["dates"]
            month_map = group_by_month(all_dates)
            al_ok(f"Attendance loaded — {len(all_dates)} dates · {len(att_data['students'])} students")
        except Exception as ex: al_err(f"Attendance parse error: {ex}")

    if syl_file and syl_file.name != st.session_state._snm:
        try:
            st.session_state.syl  = parse_syllabus(syl_file.read(), syl_file.name)
            st.session_state._snm = syl_file.name
            syl_data  = st.session_state.syl
            syl_units = list(syl_data.keys())
            total_t   = sum(len(v) for v in syl_data.values())
            al_ok(f"Syllabus loaded — {len(syl_units)} unit(s) · {total_t} topics extracted")
        except Exception as ex: al_warn(f"Syllabus warning: {ex}")

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 2 — SELECT MONTHS & DATES
    # ══════════════════════════════════════════════════════════════════
    sec("📅  Step 2 — Select Session Months & Dates")

    if not month_map:
        al_info("Upload the attendance sheet first — months and dates will appear here automatically.")
        selected_dates = []
    else:
        al_info(
            "Select which <b>months</b> to include (tick the checkbox). "
            "Then <b>expand each month</b> to pick individual dates. "
            "All dates are pre-selected — uncheck any date to exclude it."
        )

        selected_dates = []
        month_keys = list(month_map.keys())

        for row_start in range(0, len(month_keys), 4):
            row_ks = month_keys[row_start:row_start+4]
            mcols  = st.columns(len(row_ks))

            for ci, ym in enumerate(row_ks):
                ds    = month_map[ym]
                label = datetime(ym[0],ym[1],1).strftime("%B %Y")

                with mcols[ci]:
                    # Month-level toggle
                    month_on = st.checkbox(
                        f"**{label}**",
                        value=True,
                        key=f"m_{ym[0]}_{ym[1]}"
                    )
                    st.caption(f"🗓️ {len(ds)} session dates available")

                    if month_on:
                        # Individual date selection inside expander
                        with st.expander(f"Choose dates — {label}", expanded=False):
                            st.caption("Uncheck any date to exclude it from sessions.")
                            for d in ds:
                                d_label = f"{d.strftime('%d %b %Y')}  ({d.strftime('%A')})"
                                chosen  = st.checkbox(
                                    d_label,
                                    value=True,
                                    key=f"d_{d.isoformat()}"
                                )
                                if chosen:
                                    selected_dates.append(d)
                        # Show chip summary
                        chips = "".join(
                            f'<span class="date-chip">{d.strftime("%d")} '
                            f'{d.strftime("%a")}</span>'
                            for d in ds
                        )
                        st.markdown(chips, unsafe_allow_html=True)

        selected_dates = sorted(set(selected_dates))

        if selected_dates:
            al_ok(
                f"<b>{len(selected_dates)} dates selected</b> — "
                f"{selected_dates[0].strftime('%d %b %Y')} to "
                f"{selected_dates[-1].strftime('%d %b %Y')}"
            )
        else:
            al_warn("Select at least one date to continue.")

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 3 — BASIC SETTINGS
    # ══════════════════════════════════════════════════════════════════
    sec("⚙️  Step 3 — Basic Settings")

    bc = st.columns([1.4,1.5,1,1,1])

    with bc[0]:
        st.markdown("**📊 Total Sessions to Generate**")
        st.caption(
            "By default, the number of sessions equals the number of selected dates above. "
            "Enable 'Custom count' only if you want a different number — "
            "e.g., you selected 27 dates but want only 25 sessions."
        )
        use_custom = st.checkbox("Use custom session count", value=False, key="cust_sess")
        if use_custom:
            total_sessions = st.number_input(
                "Enter custom session count",
                min_value=1, max_value=500,
                value=len(selected_dates) or 50,
                help="This sets exactly how many rows the session sheet will have."
            )
        else:
            total_sessions = len(selected_dates)
            st.metric("Sessions (= selected dates)", total_sessions)

    with bc[1]:
        faculty_id = st.text_input(
            "🧑‍🏫 Faculty Registration ID",
            placeholder="e.g. IILMGG006412025",
            help="This appears in every row of the session sheet."
        )

    with bc[2]:
        mandatory = st.selectbox(
            "✅ Attendance Mandatory?",
            ["TRUE","FALSE"],
            help="Value for the 'Attendance Mandatory*' column."
        )

    with bc[3]:
        tlo_max = st.number_input(
            "Max TLO Number",
            min_value=1, max_value=100, value=5,
            help="e.g. 5 creates TLO1 through TLO5 as options."
        )

    with bc[4]:
        tlo_prefix = st.text_input(
            "TLO Prefix",
            value="TLO",
            help="'TLO' → TLO1, TLO2 … | 'CO' → CO1, CO2 …"
        )

    all_tlos = [f"{tlo_prefix}{i}" for i in range(1, tlo_max+1)]

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 4 — TIMING (7 days, AM/PM)
    # ══════════════════════════════════════════════════════════════════
    sec("⏰  Step 4 — Lecture Timing (Day-wise)")
    al_info(
        "Enable the days on which lectures are scheduled. "
        "Set Start and End time for each enabled day using the AM/PM dropdowns. "
        "Days with sessions in your selected dates are auto-detected."
    )

    detected = {}
    for d in selected_dates:
        dn = d.strftime("%A"); detected[dn] = detected.get(dn,0)+1

    day_timing = {}
    cols7 = st.columns(7)

    for i, day in enumerate(ALL_DAYS):
        cnt = detected.get(day,0)
        with cols7[i]:
            enabled = st.checkbox(
                f"**{SHORT_DAY[day]}**",
                value=(cnt>0),
                key=f"de_{day}"
            )
            card_cls = "day-card on" if enabled else "day-card off"
            st.markdown(
                f'<div class="{card_cls}">'
                f'<div class="day-name">{day[:3]}</div>'
                f'<div class="day-sess">{"🔵 "+str(cnt)+" sessions" if cnt else "—"}</div>',
                unsafe_allow_html=True
            )
            if enabled:
                h_list = list(range(1,13))
                m_list = ["00","05","10","15","20","25","30","35","40","45","50","55"]
                ap     = ["AM","PM"]
                h1,m1,a1 = st.columns([3,3,3])
                with h1: sv_h=st.selectbox("",h_list, index=9, key=f"sh_{day}",label_visibility="collapsed")
                with m1: sv_m=st.selectbox("",m_list, index=0, key=f"sm_{day}",label_visibility="collapsed")
                with a1: sv_a=st.selectbox("",ap,     index=1, key=f"sa_{day}",label_visibility="collapsed")
                s24=to_24h(sv_h,sv_m,sv_a)
                h2,m2,a2 = st.columns([3,3,3])
                with h2: ev_h=st.selectbox("",h_list, index=10,key=f"eh_{day}",label_visibility="collapsed")
                with m2: ev_m=st.selectbox("",m_list, index=0, key=f"em_{day}",label_visibility="collapsed")
                with a2: ev_a=st.selectbox("",ap,     index=1, key=f"ea_{day}",label_visibility="collapsed")
                e24=to_24h(ev_h,ev_m,ev_a)
                day_timing[day]=(s24,e24)
                st.markdown(
                    f'<div class="time-tag">⏱ {s24}–{e24}</div></div>',
                    unsafe_allow_html=True)
            else:
                st.markdown('</div>', unsafe_allow_html=True)

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 5 — MODULE & TOPIC CONFIGURATION
    # ══════════════════════════════════════════════════════════════════
    sec("📚  Step 5 — Module & Topic Configuration")

    unit_configs = []

    if not syl_units:
        al_warn(
            "No syllabus uploaded. You can still generate — "
            "topic titles will be 'Session 1', 'Session 2', etc. "
            "Upload a syllabus file to auto-fill titles and descriptions."
        )
    else:
        al_info(
            f"<b>{len(syl_units)} unit(s)</b> detected in your syllabus. "
            "For each unit: enter the <b>Module Name</b> (appears in the 'Module Name*' column), "
            "assign <b>TLOs</b>, then choose <b>Full Unit</b> (all topics) or "
            "<b>Custom Selection</b> (pick specific topics)."
        )

        for i, unit_key in enumerate(syl_units):
            avail = syl_data[unit_key]

            with st.expander(
                f"Unit {i+1}  ·  {unit_key[:70]}  ({len(avail)} topics)",
                expanded=(i < 4)
            ):
                r1c1, r1c2 = st.columns([3,2])

                with r1c1:
                    mod_name = st.text_input(
                        "Module Name*  (this text goes into the session sheet)",
                        value=unit_key,
                        key=f"mn_{i}",
                        help="Exactly what appears in the 'Module Name*' column of the output sheet."
                    )

                with r1c2:
                    def_tlo  = [all_tlos[i % len(all_tlos)]] if all_tlos else [f"{tlo_prefix}1"]
                    sel_tlos = st.multiselect(
                        f"TLOs for this module",
                        options=all_tlos,
                        default=def_tlo,
                        key=f"tl_{i}",
                        help="Select one or more TLOs. Multiple will appear as 'TLO1 | TLO2'."
                    )
                    tlo_str  = " | ".join(sel_tlos) if sel_tlos else all_tlos[i%len(all_tlos)]

                cov = st.radio(
                    "Topic Coverage:",
                    ["📖 Full Unit — include all topics from this unit",
                     "✂️ Custom Selection — I want to choose specific topics"],
                    key=f"cov_{i}",
                    horizontal=True
                )

                if "Custom" in cov:
                    chosen = st.multiselect(
                        "Select topics to include:",
                        options=avail,
                        default=avail,
                        key=f"tp_{i}",
                        help="All topics are pre-selected. Remove any you don't want."
                    )
                    final_topics = chosen
                    st.caption(
                        f"✅ {len(final_topics)} of {len(avail)} topics selected for this module."
                    )
                else:
                    final_topics = avail
                    with st.expander(f"Preview all {len(avail)} topics →"):
                        for j,t in enumerate(avail,1):
                            st.write(f"`{j}.` {t}")

                unit_configs.append({
                    "module_name": mod_name,
                    "tlo":         tlo_str,
                    "topics":      final_topics,
                })

        # Summary
        n_selected_topics = sum(len(u["topics"]) for u in unit_configs)
        st.markdown("<br>", unsafe_allow_html=True)
        sum_cols = st.columns(3)
        with sum_cols[0]:
            st.markdown(
                f'<div class="stat-card"><div class="stat-num">{len(unit_configs)}</div>'
                f'<div class="stat-lbl">Modules Configured</div></div>',
                unsafe_allow_html=True)
        with sum_cols[1]:
            st.markdown(
                f'<div class="stat-card"><div class="stat-num">{n_selected_topics}</div>'
                f'<div class="stat-lbl">Topics Selected</div></div>',
                unsafe_allow_html=True)
        with sum_cols[2]:
            st.markdown(
                f'<div class="stat-card"><div class="stat-num">{total_sessions}</div>'
                f'<div class="stat-lbl">Sessions to Generate</div></div>',
                unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        if n_selected_topics == total_sessions:
            al_ok("Perfect match — one topic per session.")
        elif n_selected_topics < total_sessions:
            al_warn(
                f"Topics ({n_selected_topics}) < Sessions ({total_sessions}). "
                f"The last {total_sessions-n_selected_topics} sessions will be revision entries."
            )
        else:
            al_warn(
                f"Topics ({n_selected_topics}) > Sessions ({total_sessions}). "
                f"Topics will be trimmed proportionally across modules."
            )

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 6 — BUILD PREVIEW TABLE
    # ══════════════════════════════════════════════════════════════════
    sec("👀  Step 6 — Preview & Edit Table")
    al_info(
        "Click <b>Build Preview Table</b> to generate the full session table. "
        "Topics are auto-balanced across modules. Descriptions are auto-generated. "
        "Every cell is editable — make changes directly before generating files."
    )

    prev_btn = st.button("🔄  Build Preview Table", type="secondary",
                         use_container_width=True)

    if prev_btn:
        if not selected_dates:
            al_err("Select at least one date in Step 2.")
        elif not faculty_id.strip():
            al_err("Enter Faculty Registration ID in Step 3.")
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
                    sdt = f"{ds} {s}:00"
                    edt = f"{ds} {e}:00"
                else:
                    sdt = edt = "TBD"

                sess = flat[idx] if idx < len(flat) else {
                    "module":"Extra","tlo":f"{tlo_prefix}1",
                    "title":f"Extra Session {idx+1}"}

                rows.append({
                    "Sr":                idx+1,
                    "Module Name*":      sess["module"],
                    "Start Date Time*":  sdt,
                    "End Date Time*":    edt,
                    "Title*":            sess["title"],
                    "Description*":      auto_desc(sess["title"],sess["module"]),
                    "Mandatory*":        mandatory,
                    "TLO":               sess["tlo"],
                    "Faculty Reg ID*":   faculty_id.strip(),
                })

            st.session_state.preview = pd.DataFrame(rows)
            st.session_state.results = None
            al_ok(f"Preview built — <b>{len(rows)} rows</b>. Edit anything in the table below.")

    if st.session_state.preview is not None:
        df  = st.session_state.preview
        edf = st.data_editor(
            df, use_container_width=True, num_rows="fixed",
            hide_index=True, key="ed_prev",
            column_config={
                "Sr":               st.column_config.NumberColumn("Sr",width=55,disabled=True),
                "Module Name*":     st.column_config.TextColumn("Module Name* ✏️",    width=235),
                "Start Date Time*": st.column_config.TextColumn("Start DateTime* ✏️", width=170),
                "End Date Time*":   st.column_config.TextColumn("End DateTime* ✏️",   width=170),
                "Title*":           st.column_config.TextColumn("Title* ✏️",          width=270),
                "Description*":     st.column_config.TextColumn("Description* ✏️",    width=390),
                "Mandatory*":       st.column_config.SelectboxColumn(
                                        "Mandatory*",options=["TRUE","FALSE"],width=115),
                "TLO":              st.column_config.TextColumn("TLO ✏️",             width=130),
                "Faculty Reg ID*":  st.column_config.TextColumn("Faculty Reg ID* ✏️", width=170),
            },
            height=500,
        )
        st.session_state.preview = edf

    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # STEP 7 — GENERATE
    # ══════════════════════════════════════════════════════════════════
    sec("🚀  Step 7 — Generate Files")

    can_gen = bool(
        st.session_state.preview is not None and
        att_data and faculty_id.strip()
    )

    if not can_gen:
        miss = []
        if not att_data:                        miss.append("Attendance Sheet (Step 1)")
        if not faculty_id.strip():              miss.append("Faculty Registration ID (Step 3)")
        if st.session_state.preview is None:   miss.append("Preview Table (Step 6)")
        if miss: al_info(f"Complete first: <b>{' · '.join(miss)}</b>")

    gen_btn = st.button(
        "⚡  Generate Session Sheet + All Day-wise Attendance Files",
        type="primary", disabled=not can_gen, use_container_width=True
    )

    if gen_btn and can_gen:
        errors=[]; prog=st.progress(0); log=st.empty()
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
                "faculty_id":  str(r.get("Faculty Reg ID*",faculty_id.strip())),
            } for _,r in df.iterrows()]
            prog.progress(20)

            log.markdown('<div class="al-info">📊  Generating session sheet…</div>',
                         unsafe_allow_html=True)
            session_xlsx = gen_session_sheet(session_rows)
            al_ok(f"Session sheet ready — <b>{len(session_rows)} rows</b> in DL.xlsx format.")
            prog.progress(50)

            daywise = {}
            if att_tpl:
                log.markdown('<div class="al-info">🗂️  Generating day-wise attendance files…</div>',
                             unsafe_allow_html=True)
                tpl_b = att_tpl.read()
                use_d = selected_dates[:total_sessions]
                bar   = st.progress(0)
                for i,d in enumerate(use_d):
                    try:
                        fn = f"attendance_{d.strftime('%Y-%m-%d')}.xlsx"
                        daywise[fn] = gen_daywise_att(tpl_b, att_data, d)
                    except Exception as ex: errors.append(f"{d}: {ex}")
                    bar.progress((i+1)/max(len(use_d),1))
                al_ok(f"Day-wise attendance ready — <b>{len(daywise)} files</b>.")
            else:
                al_warn("No attendance template uploaded — only session sheet will be available.")

            prog.progress(90)
            zip_b = build_zip(session_xlsx, daywise)
            prog.progress(100)
            log.markdown(
                '<div class="al-ok">🎉  All files generated successfully! Download below.</div>',
                unsafe_allow_html=True)

            st.session_state.results = {
                "session":session_xlsx,"daywise":daywise,
                "zip":zip_b,"errors":errors,"n":len(session_rows)
            }
        except Exception as ex:
            al_err(f"Error: {ex}")
            import traceback; st.code(traceback.format_exc())

    # ══════════════════════════════════════════════════════════════════
    # STEP 8 — DOWNLOADS
    # ══════════════════════════════════════════════════════════════════
    if st.session_state.results:
        res = st.session_state.results
        st.divider()
        sec("📥  Step 8 — Download")
        for e in res.get("errors",[]): al_warn(f"Skipped: {e}")

        d1,d2 = st.columns(2)
        with d1:
            st.download_button(
                "📦  Download Everything as ZIP",
                data=res["zip"],
                file_name=f"DG_Output_{datetime.now():%Y%m%d_%H%M}.zip",
                mime="application/zip",
                use_container_width=True
            )
        with d2:
            st.download_button(
                "📋  Download session_sheet.xlsx only",
                data=res["session"], file_name="session_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        if res["daywise"]:
            st.markdown(f"#### 📅  Day-wise Attendance Files  ({len(res['daywise'])} files)")
            files = list(res["daywise"].items())
            for ri in range(0,len(files),4):
                chunk=files[ri:ri+4]; cols=st.columns(4)
                for ci,(fn,fb) in enumerate(chunk):
                    with cols[ci]:
                        dp = fn.replace("attendance_","").replace(".xlsx","")
                        try:    lbl=datetime.strptime(dp,"%Y-%m-%d").strftime("%d %b %Y")
                        except: lbl=dp
                        st.download_button(
                            f"📅  {lbl}", data=fb, file_name=fn,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True, key=f"dl_{fn}"
                        )

        st.markdown("<br>",unsafe_allow_html=True)
        if st.button("🔄  Reset — Start Over", type="secondary"):
            for k in ["att","syl","preview","results","_anm","_snm"]:
                st.session_state[k]=None
            st.rerun()

    # ── Footer ─────────────────────────────────────────────────────────
    st.markdown("""
    <div class="footer">
      Created by <b>Dr. Amar Shukla</b>  ·  IILM University, Gurugram  ·
      All processing is in-memory — no data is stored on any server
    </div>""", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
