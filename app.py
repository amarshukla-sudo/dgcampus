"""
DG Sheet Generator v12.0
Created by Dr. Amar Shukla · IILM University, Gurugram
=========================================================
TWO SEPARATE TOOLS:
  1. Day-wise Attendance Generator
  2. Session Sheet Generator

ATTENDANCE LOGIC (correct):
  - Row 5  = session dates
  - Col 2  = Enrollment No.
  - Row 6+ = student attendance (0/1)
  - Match DG template Registration Id with master Enrollment No.
  - 1 → PRESENT (green), 0 or not found → ABSENT (red)
"""

import io, re, zipfile, warnings
from collections import OrderedDict, defaultdict
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="DG Sheet Generator — Dr. Amar Shukla",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
html,[class*="css"]{font-size:16px!important}
.main{padding:.5rem 2rem 3rem}
.banner{background:linear-gradient(135deg,#0a1f35,#1565c0);color:#fff;
        padding:1.3rem 2rem;border-radius:14px;margin-bottom:1rem}
.banner h1{margin:0;font-size:1.9rem;font-weight:900}
.banner p{margin:.3rem 0 0;font-size:.96rem;opacity:.87}
.banner small{font-size:.82rem;opacity:.65;font-style:italic}
.sec{font-size:1.12rem;font-weight:800;color:#0a1f35;
     border-left:6px solid #1565c0;padding:.28rem 0 .28rem .7rem;
     margin:1.1rem 0 .55rem}
.ib{background:#e3f2fd;border-left:5px solid #1565c0;border-radius:0 8px 8px 0;
    padding:.5rem 1rem;margin:.3rem 0;font-size:.93rem;color:#0d47a1}
.ok{background:#e8f5e9;border-left:5px solid #2e7d32;border-radius:0 8px 8px 0;
    padding:.5rem 1rem;margin:.3rem 0;font-size:.93rem;color:#1b5e20}
.wn{background:#fff8e1;border-left:5px solid #f57f17;border-radius:0 8px 8px 0;
    padding:.5rem 1rem;margin:.3rem 0;font-size:.93rem;color:#e65100}
.er{background:#ffebee;border-left:5px solid #c62828;border-radius:0 8px 8px 0;
    padding:.5rem 1rem;margin:.3rem 0;font-size:.93rem;color:#b71c1c}
.stat-row{display:flex;gap:.8rem;margin:.6rem 0;flex-wrap:wrap}
.sc{background:#fff;border:2px solid #e3f2fd;border-radius:12px;
    padding:.8rem 1rem;text-align:center;flex:1;min-width:110px}
.sn{font-size:2rem;font-weight:900;color:#1565c0;line-height:1.1}
.sl{font-size:.8rem;color:#546e7a;margin-top:.2rem;font-weight:600}
.date-chip{display:inline-block;background:#1565c0;color:#fff;
           border-radius:14px;padding:.13rem .5rem;margin:.1rem;
           font-size:.75rem;font-weight:700}
.stButton>button{font-size:1rem!important;font-weight:700!important;
                 padding:.55rem 1.4rem!important;border-radius:10px!important}
.stDownloadButton>button{background:#0a1f35!important;color:#fff!important;
    border-radius:10px!important;font-weight:700!important;
    font-size:.92rem!important;width:100%!important}
.stDownloadButton>button:hover{background:#1565c0!important}
label{font-size:1rem!important;font-weight:600!important}
.streamlit-expanderHeader{font-size:1rem!important;font-weight:700!important}
.badge{display:inline-block;background:#e8f5e9;color:#2e7d32;
       padding:.13rem .6rem;border-radius:12px;font-size:.78rem;font-weight:700}
hr{border:none;border-top:2.5px solid #e3f2fd;margin:1rem 0}
.footer{text-align:center;color:#90a4ae;font-size:.8rem;margin-top:1.5rem;
        padding-top:1rem;border-top:2px solid #e3f2fd}
</style>
""", unsafe_allow_html=True)

def ib(m): st.markdown(f'<div class="ib">ℹ️  {m}</div>', unsafe_allow_html=True)
def ok(m): st.markdown(f'<div class="ok">✅  {m}</div>', unsafe_allow_html=True)
def wn(m): st.markdown(f'<div class="wn">⚠️  {m}</div>', unsafe_allow_html=True)
def er(m): st.markdown(f'<div class="er">❌  {m}</div>', unsafe_allow_html=True)
def sec(t): st.markdown(f'<div class="sec">{t}</div>', unsafe_allow_html=True)

ALL_DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]

# ══════════════════════════════════════════════════════════════════════
# CORE: READ MASTER ATTENDANCE (openpyxl - reads exact cell values)
# Structure:  Row 5 = dates, Col 2 = Enrollment, Row 6+ = 0/1 data
# ══════════════════════════════════════════════════════════════════════

def clean_date_str(raw) -> str:
    """Strip day name, 'Sub', extra whitespace from date string."""
    s = str(raw).strip()
    s = re.sub(r"\([^)]*\)", "", s)           # remove (Friday) etc
    s = re.sub(r"\bSub\b", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_date_str(raw) -> date | None:
    if raw is None: return None
    s = clean_date_str(raw)
    if not s or s.lower() in ("nan","none",""): return None
    for fmt in ("%d %b %Y", "%d %b %y", "%d %B %Y",
                "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try: return datetime.strptime(s, fmt).date()
        except: pass
    # fallback pandas
    try:
        r = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(r): return r.date()
    except: pass
    return None


def read_master_attendance(fb: bytes) -> dict:
    """
    Read master attendance sheet using openpyxl (exact cell values).

    STRUCTURE (auto-detected):
      - Scans for the row containing the most parseable dates → DATE ROW
      - Col 2 = Enrollment No. (hardcoded based on standard format,
        but also auto-detected by header scan)
      - Rows after date row = student data

    Returns:
      {
        'dates': [date, ...],              # ordered unique session dates
        'col_date': {col_idx: date},       # column number → date
        'enroll_col': int,                 # column index of enrollment
        'date_att': {date: {enroll: 0/1}} # attendance per date
      }
    """
    wb = load_workbook(io.BytesIO(fb), data_only=True)
    # Pick densest sheet
    best_ws = None; best_n = -1
    for sn in wb.sheetnames:
        ws = wb[sn]
        n = sum(1 for row in ws.iter_rows() for c in row if c.value is not None)
        if n > best_n: best_n, best_ws = n, ws
    ws = best_ws

    # ── Find date row ─────────────────────────────────────────────────
    date_row_num = None
    date_col_start = None
    best_hits = 0

    for r in range(1, min(ws.max_row+1, 20)):
        hits = 0; first = None
        for c in range(1, ws.max_column+1):
            v = ws.cell(r, c).value
            if v and parse_date_str(v):
                hits += 1
                if first is None: first = c
        if hits > best_hits:
            best_hits = hits
            date_row_num = r
            date_col_start = first

    if not date_row_num or best_hits < 2:
        raise ValueError(
            "Could not find session dates in the attendance sheet.\n"
            "Expected: a row with multiple dates like '2 Jan 2026 (Friday)'."
        )

    # ── Build col→date map ────────────────────────────────────────────
    col_date: dict[int, date] = {}
    dates_ordered: list[date] = []
    seen: set[date] = set()

    for c in range(date_col_start, ws.max_column+1):
        v = ws.cell(date_row_num, c).value
        if v is None: continue
        d = parse_date_str(v)
        if d:
            col_date[c] = d
            if d not in seen:
                dates_ordered.append(d)
                seen.add(d)

    # ── Find enrollment column ────────────────────────────────────────
    enroll_col = None
    for r in range(1, date_row_num):
        for c in range(1, date_col_start):
            v = str(ws.cell(r, c).value or "").lower()
            if "enrol" in v or "roll no" in v or "registration" in v:
                enroll_col = c; break
        if enroll_col: break
    if enroll_col is None: enroll_col = 2  # default: column 2

    # ── Parse student attendance ──────────────────────────────────────
    date_att: dict[date, dict[str, int]] = {d: {} for d in dates_ordered}

    for r in range(date_row_num+1, ws.max_row+1):
        raw_enroll = ws.cell(r, enroll_col).value
        if raw_enroll is None: continue
        # Normalize enrollment: remove trailing .0
        enroll = re.sub(r"\.0+$", "", str(raw_enroll).strip())
        if not enroll or enroll.lower() in ("nan","none",""): continue

        for c, d in col_date.items():
            val = ws.cell(r, c).value
            try:    att = int(float(val)) if val is not None else 0
            except: att = 0
            date_att[d][enroll] = att

    return {
        "dates":      dates_ordered,
        "col_date":   col_date,
        "enroll_col": enroll_col,
        "date_att":   date_att,
    }


def group_by_month(dates: list) -> OrderedDict:
    res = defaultdict(list)
    for d in dates:
        if d: res[(d.year, d.month)].append(d)
    return OrderedDict(sorted(res.items()))


# ══════════════════════════════════════════════════════════════════════
# GENERATE DAY-WISE ATTENDANCE
# ══════════════════════════════════════════════════════════════════════

def gen_one_attendance(tpl_bytes: bytes, day_att: dict[str, int]) -> bytes:
    """
    Fill one DG attendance template for a single date.
    day_att = {enrollment_str: 0/1}
    Rule: 1=PRESENT, 0 or not-found=ABSENT
    """
    wb = load_workbook(io.BytesIO(tpl_bytes))
    ws = wb.active

    # Detect columns from row 1 header
    rc = ac = None
    for c in range(1, ws.max_column+1):
        hv = str(ws.cell(1, c).value or "").lower()
        if "registration" in hv or ("reg" in hv and "id" in hv): rc = c
        elif "attendance" in hv: ac = c
    if rc is None: rc = 2
    if ac is None: ac = 3

    GF = PatternFill("solid", start_color="C6EFCE")
    RF = PatternFill("solid", start_color="FFC7CE")
    GN = Font(color="006100", bold=True, size=10)
    RN = Font(color="9C0006", bold=True, size=10)
    CA = Alignment(horizontal="center", vertical="center")

    for r in range(2, ws.max_row+1):
        raw = ws.cell(r, rc).value
        if raw is None: continue
        reg = re.sub(r"\.0+$", "", str(raw).strip())
        if not reg: continue

        # 1 in master → PRESENT, 0 or missing → ABSENT
        att_val = day_att.get(reg, 0)
        status  = "PRESENT" if att_val == 1 else "ABSENT"

        cell = ws.cell(r, ac)
        cell.value     = status
        cell.fill      = GF if status == "PRESENT" else RF
        cell.font      = GN if status == "PRESENT" else RN
        cell.alignment = CA

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════
# GENERATE SESSION SHEET (exact DL.xlsx format)
# ══════════════════════════════════════════════════════════════════════

UNIT_RE = re.compile(
    r"^(UNIT|MODULE|CHAPTER|TOPIC|SECTION|PART|BLOCK)\s*[-:.]?\s*\d+", re.IGNORECASE)
SKIP_RE = re.compile(
    r"\bv\.\b|\bAIR\b|\bILR\b|\bSCC\b|https?://|\(\d{4}\)\s+\d+|^\s*Case\s+Law",
    re.IGNORECASE)


def best_sheet_df(fb: bytes, hints=None) -> pd.DataFrame:
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


def parse_syllabus(fb: bytes, filename: str) -> OrderedDict:
    ext = filename.lower().rsplit(".",1)[-1]
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
            if UNIT_RE.match(t): cur=t; result[cur]=[]
            elif cur: add(cur, t)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    t = cell.text.strip()
                    if not t: continue
                    if UNIT_RE.match(t): cur=t; result[cur]=[]
                    elif cur: add(cur, t)
    elif ext in ("xlsx","xls"):
        df = best_sheet_df(fb); cur=None
        for _,row in df.iterrows():
            for val in row:
                s=str(val).strip()
                if not s or s.lower() in ("nan","none",""): continue
                if UNIT_RE.match(s): cur=s; result[cur]=[]; break
                elif cur and len(s)>4: add(cur,s); break
    elif ext in ("txt","csv"):
        cur=None
        for line in fb.decode("utf-8","ignore").splitlines():
            t=line.strip()
            if not t: continue
            if UNIT_RE.match(t): cur=t; result[cur]=[]
            elif cur: add(cur,t)

    if not result:
        result["Module 1"]=[]
        if ext in ("txt","csv"):
            for line in fb.decode("utf-8","ignore").splitlines(): add("Module 1",line)
        elif ext=="docx":
            import docx as _d
            for p in _d.Document(io.BytesIO(fb)).paragraphs: add("Module 1",p.text)
        elif ext in ("xlsx","xls"):
            df=best_sheet_df(fb)
            for _,row in df.iterrows():
                for val in row:
                    s=str(val).strip()
                    if s and s.lower() not in ("nan","none","") and len(s)>4:
                        add("Module 1",s); break

    return OrderedDict((k,v) for k,v in result.items() if v)


def auto_desc(title:str, module:str) -> str:
    t=title.strip()
    mod=re.sub(r"^(UNIT|MODULE|CHAPTER)\s*\d+[:\-.]?\s*","",module,flags=re.IGNORECASE).strip() or module
    tl=t.lower()
    if any(w in tl for w in ["definition","define","meaning"]):
        core=re.sub(r"^(definition\s+(of\s+)?|define\s+)","",t,flags=re.IGNORECASE).strip()
        return (f"This session covers the statutory definition and essential elements of {core}. "
                f"Students will examine the relevant legal provisions, judicial interpretations, "
                f"and practical significance within the framework of {mod}.")
    if any(w in tl for w in ["rights","duties","liability","liabilities"]):
        return (f"This session examines the rights, duties, and liabilities under {t}. "
                f"Students will analyse statutory provisions, landmark judgments, "
                f"and the legal consequences for parties involved in {mod}.")
    if any(w in tl for w in ["distinction","difference","compare","versus"]):
        return (f"This session provides a comparative analysis of {t}. Students will identify "
                f"key distinctions through statutory provisions and case-law in {mod}.")
    if any(w in tl for w in ["type","kind","classif","categor","nature"]):
        return (f"This session explores the types and conceptual framework of {t}. "
                f"Students will study the legal significance of each category within {mod}.")
    if any(w in tl for w in ["termination","discharge","revocation","dissolution"]):
        return (f"This session covers the modes of {t} and their legal consequences. "
                f"Students will study statutory provisions and judicial precedents in {mod}.")
    if any(w in tl for w in ["creation","formation","essential","element"]):
        return (f"This session discusses the formation process and requisite elements of {t}. "
                f"Students will examine statutory requirements and practical illustrations in {mod}.")
    if any(w in tl for w in ["remedy","remedies","damages"]):
        return (f"This session discusses remedies available in {t}. Students will examine "
                f"statutory remedies, judicial approaches, and computation of relief in {mod}.")
    if any(w in tl for w in ["case","judgment"," v."]):
        return (f"This session analyses {t} as a significant judicial decision. Students will "
                f"examine the facts, legal issues, court reasoning, and precedential value in {mod}.")
    return (f"This session provides a comprehensive study of {t} within {mod}. "
            f"Students will analyse statutory provisions, judicial precedents, "
            f"and practical applications through structured discussion.")


def balance_topics(unit_configs:list, total:int) -> list:
    items=[]
    for uc in unit_configs:
        for t in uc["topics"]:
            items.append({"module":uc["module_name"],"tlo":uc["tlo"],"title":t})
    n=len(items)
    if n==0:
        return [{"module":"Session","tlo":"TLO1","title":f"Session {i+1}"} for i in range(total)]
    if n<=total:
        result=list(items)
        for i in range(total-n):
            b=items[i%n]
            result.append({"module":b["module"],"tlo":b["tlo"],"title":b["title"]+" — Revision"})
        return result
    ta=sum(len(uc["topics"]) for uc in unit_configs)
    result=[]
    for uc in unit_configs:
        n_take=max(1,round(total*len(uc["topics"])/ta))
        n_take=min(n_take,len(uc["topics"]))
        for t in uc["topics"][:n_take]:
            result.append({"module":uc["module_name"],"tlo":uc["tlo"],"title":t})
    while len(result)<total: result.append(result[-1]|{"title":result[-1]["title"]+" (Cont.)"})
    return result[:total]


HEADERS=["Module Name*","Start Date Time*","End Date Time*","Title*",
         "Description*","Attendance Mandatory*","TLO","Teaching Faculty Registration ID*"]
COL_W  =[30,22,22,42,60,22,18,32]
HF=PatternFill("solid",start_color="1F4E79"); HT=Font(bold=True,color="FFFFFF",size=11)
DA=PatternFill("solid",start_color="D6EAF8"); DB=PatternFill("solid",start_color="EBF5FB")
TN=Side(style="thin",color="BBBBBB"); BD=Border(left=TN,right=TN,top=TN,bottom=TN)
CT=Alignment(horizontal="center",vertical="center",wrap_text=True)
LF=Alignment(horizontal="left",  vertical="center",wrap_text=True)


def gen_session_sheet(rows:list) -> bytes:
    wb=Workbook(); ws=wb.active; ws.title="Session Sheet"
    aligns=[LF,CT,CT,LF,LF,CT,CT,CT]
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
            cell.border=BD; cell.alignment=al; cell.font=Font(size=10); cell.number_format="@"
        ws.row_dimensions[i].height=30
    ws.freeze_panes="A2"
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()


def parse_time(s:str) -> tuple:
    s=s.strip().upper(); am_pm=None
    if "AM" in s or "PM" in s:
        am_pm="PM" if "PM" in s else "AM"
        s=s.replace("AM","").replace("PM","").strip()
    s=re.sub(r"[.h\-]",":",s); parts=re.split(r":",s)
    try:
        h,m=int(parts[0]),int(parts[1]) if len(parts)>1 else 0
        if am_pm=="PM" and h!=12: h+=12
        if am_pm=="AM" and h==12: h=0
        return h,m
    except: return 10,0


def fmt_dt(d:date,h:int,m:int)->str:
    return f"{d.strftime('%Y-%m-%d')} {h:02d}:{m:02d}:00"


def build_zip(files:dict) -> bytes:
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as zf:
        for fn,fb in files.items(): zf.writestr(fn,fb)
    buf.seek(0); return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════
#                           MAIN UI
# ══════════════════════════════════════════════════════════════════════

def main():
    st.markdown("""
    <div class="banner">
      <h1>📋 DG Sheet Generator</h1>
      <p>Two separate tools — Attendance Generator &amp; Session Sheet Generator</p>
      <small>Created by Dr. Amar Shukla · IILM University, Gurugram</small>
    </div>""", unsafe_allow_html=True)

    for k in ["att_data","_att_nm","syl_data","_syl_nm","sess_preview","sess_results"]:
        if k not in st.session_state: st.session_state[k]=None

    # TOOL SELECTOR
    tool = st.radio(
        "**Select Tool:**",
        ["📅 Day-wise Attendance Generator", "📋 Session Sheet Generator"],
        horizontal=True
    )
    st.divider()

    # ══════════════════════════════════════════════════════════════════
    # TOOL 1: DAY-WISE ATTENDANCE GENERATOR
    # ══════════════════════════════════════════════════════════════════
    if "Attendance" in tool:
        sec("📅  Day-wise Attendance Generator")
        ib(
            "Upload your <b>Master Attendance Sheet</b> and the <b>DG Attendance Template</b>. "
            "Select which dates to generate. "
            "PRESENT/ABSENT is determined from the master sheet — "
            "students marked <b>1 = PRESENT</b>, <b>0 or not found = ABSENT</b>."
        )

        # Files
        a1,a2 = st.columns(2)
        with a1:
            st.markdown("**📅 Master Attendance Sheet**")
            st.caption("Excel file with dates in one row and 0/1 attendance per student")
            att_file = st.file_uploader("",type=["xlsx","xls"],key="att_up",
                                        label_visibility="collapsed")
            if att_file:
                st.markdown(f'<span class="badge">✅ {att_file.name}</span>',
                            unsafe_allow_html=True)
        with a2:
            st.markdown("**🗂️ DG Attendance Template**")
            st.caption("Pre-filled template with Email Id, Registration Id, Attendance* columns")
            dg_tpl = st.file_uploader("",type=["xlsx"],key="dg_tpl_up",
                                      label_visibility="collapsed")
            if dg_tpl:
                st.markdown(f'<span class="badge">✅ {dg_tpl.name}</span>',
                            unsafe_allow_html=True)

        # Parse master attendance
        att_data = st.session_state.att_data
        if att_file and att_file.name != st.session_state._att_nm:
            try:
                att_data = read_master_attendance(att_file.read())
                st.session_state.att_data = att_data
                st.session_state._att_nm  = att_file.name
                n_dates = len(att_data["dates"])
                n_stud  = len(next(iter(att_data["date_att"].values()),{}))
                ok(f"Attendance loaded — <b>{n_dates} session dates</b> detected")
            except Exception as ex:
                er(f"Could not read attendance sheet: {ex}")
                att_data = None

        if att_data is None and att_file is None:
            return

        if att_data:
            all_dates = att_data["dates"]
            month_map = group_by_month(all_dates)

            # Stats
            sample_day = all_dates[0] if all_dates else None
            n_stud = len(att_data["date_att"].get(sample_day,{})) if sample_day else 0
            st.markdown(f"""
            <div class="stat-row">
              <div class="sc"><div class="sn">{len(all_dates)}</div>
                <div class="sl">Session Dates</div></div>
              <div class="sc"><div class="sn">{len(month_map)}</div>
                <div class="sl">Months</div></div>
              <div class="sc"><div class="sn">{n_stud}</div>
                <div class="sl">Students in Master</div></div>
            </div>""", unsafe_allow_html=True)

            st.divider()

            # Date selection
            sec("Select Dates to Generate")
            ib("Tick months to include. Expand each month to select/deselect individual dates.")

            selected_dates = []
            for row_start in range(0, len(month_map), 4):
                keys_row = list(month_map.keys())[row_start:row_start+4]
                cols     = st.columns(len(keys_row))
                for ci, ym in enumerate(keys_row):
                    ds    = month_map[ym]
                    label = datetime(ym[0],ym[1],1).strftime("%B %Y")
                    with cols[ci]:
                        mon_on = st.checkbox(f"**{label}**", value=True,
                                             key=f"am_{ym[0]}_{ym[1]}")
                        st.caption(f"🗓️ {len(ds)} dates")
                        if mon_on:
                            with st.expander(f"Dates in {label}", expanded=False):
                                for d in ds:
                                    lbl = f"{d.strftime('%d %b %Y')} ({d.strftime('%A')})"
                                    if st.checkbox(lbl, value=True,
                                                   key=f"ad_{d.isoformat()}"):
                                        selected_dates.append(d)
                            chips = "".join(
                                f'<span class="date-chip">'
                                f'{d.strftime("%d")} {d.strftime("%a")}'
                                f'</span>' for d in ds
                            )
                            st.markdown(chips, unsafe_allow_html=True)

            selected_dates = sorted(set(selected_dates))
            if selected_dates:
                ok(f"<b>{len(selected_dates)} dates selected</b>")
            else:
                wn("Select at least one date.")

            st.divider()

            # Generate button
            can_gen = bool(selected_dates and dg_tpl)
            if not can_gen:
                if not dg_tpl:        ib("Upload the DG Attendance Template above.")
                if not selected_dates: ib("Select at least one date.")

            if st.button("⚡  Generate Day-wise Attendance Files",
                         type="primary", disabled=not can_gen,
                         use_container_width=True):

                tpl_bytes = dg_tpl.read()
                all_files: dict[str,bytes] = {}
                errors    = []
                bar       = st.progress(0)
                log       = st.empty()

                for i, d in enumerate(selected_dates):
                    try:
                        fn       = f"attendance_{d.strftime('%Y-%m-%d')}.xlsx"
                        day_att  = att_data["date_att"].get(d, {})
                        fb_out   = gen_one_attendance(tpl_bytes, day_att)
                        all_files[fn] = fb_out
                    except Exception as ex:
                        errors.append(f"{d.strftime('%d %b %Y')}: {ex}")
                    bar.progress((i+1)/len(selected_dates))

                if errors:
                    for e in errors: wn(f"Skipped: {e}")

                # Verify sample
                if all_files:
                    first_fb = list(all_files.values())[0]
                    ws_chk   = load_workbook(io.BytesIO(first_fb)).active
                    p_cnt    = sum(1 for r in range(2,ws_chk.max_row+1)
                                   if ws_chk.cell(r,3).value=="PRESENT")
                    a_cnt    = sum(1 for r in range(2,ws_chk.max_row+1)
                                   if ws_chk.cell(r,3).value=="ABSENT")
                    first_d  = list(selected_dates)[0]
                    ok(f"<b>{len(all_files)} files generated.</b> "
                       f"Sample ({first_d.strftime('%d %b %Y')}): "
                       f"PRESENT = {p_cnt} · ABSENT = {a_cnt}")

                    # ZIP
                    zip_b = build_zip(all_files)

                    st.divider()
                    sec("Download")

                    d1, d2 = st.columns(2)
                    with d1:
                        st.download_button(
                            "📦  Download All as ZIP",
                            data=zip_b,
                            file_name=f"Attendance_{datetime.now():%Y%m%d_%H%M}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )

                    # Individual files
                    files_list = list(all_files.items())
                    for ri in range(0, len(files_list), 4):
                        chunk = files_list[ri:ri+4]
                        cols  = st.columns(4)
                        for ci,(fn,fb) in enumerate(chunk):
                            with cols[ci]:
                                dp = fn.replace("attendance_","").replace(".xlsx","")
                                try:    lbl=datetime.strptime(dp,"%Y-%m-%d").strftime("%d %b %Y")
                                except: lbl=dp
                                st.download_button(
                                    f"📅 {lbl}", data=fb, file_name=fn,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    key=f"att_dl_{fn}"
                                )

    # ══════════════════════════════════════════════════════════════════
    # TOOL 2: SESSION SHEET GENERATOR
    # ══════════════════════════════════════════════════════════════════
    else:
        sec("📋  Session Sheet Generator")
        ib(
            "Enter how many sessions to create, paste syllabus unit-wise, "
            "configure timing and faculty details. "
            "Topics are extracted from your pasted text and auto-distributed across sessions. "
            "Descriptions are auto-generated per topic."
        )

        st.divider()

        # ── STEP A: Basic Info ────────────────────────────────────────
        sec("Step 1 — Basic Information")

        sa1,sa2,sa3,sa4,sa5 = st.columns([1.2,1.8,1,1,1])
        with sa1:
            total_sessions = st.number_input(
                "📊 Total Sessions to Create",
                min_value=1, max_value=500, value=40,
                help="This is the total number of rows in the output sheet."
            )
        with sa2:
            faculty_id = st.text_input(
                "🧑‍🏫 Faculty Registration ID",
                placeholder="e.g. IILMGG006412025"
            )
        with sa3:
            mandatory = st.selectbox("Mandatory?", ["TRUE","FALSE"])
        with sa4:
            tlo_max = st.number_input("Max TLO", min_value=1, max_value=100, value=5)
        with sa5:
            tlo_pfx = st.text_input("TLO Prefix", value="TLO",
                                     help="'TLO'→TLO1,TLO2 | 'CO'→CO1,CO2")

        all_tlos = [f"{tlo_pfx}{i}" for i in range(1, tlo_max+1)]

        st.divider()

        # ── STEP B: Dates (optional from attendance) ──────────────────
        sec("Step 2 — Session Dates  (Optional)")
        ib(
            "Upload your Master Attendance Sheet to auto-fill dates in the session sheet. "
            "If not uploaded, the Start/End Date Time columns will be left as TBD — "
            "you can fill them manually in the preview table."
        )

        att_file_s = st.file_uploader(
            "📅 Master Attendance Sheet (optional)",
            type=["xlsx","xls"], key="att_s_up"
        )
        sess_dates = []
        if att_file_s:
            try:
                att_s      = read_master_attendance(att_file_s.read())
                sess_dates = att_s["dates"]
                ok(f"Dates loaded — <b>{len(sess_dates)} session dates</b> from attendance sheet")
            except Exception as ex:
                wn(f"Could not read dates: {ex}")

        st.divider()

        # ── STEP C: Timing ────────────────────────────────────────────
        sec("Step 3 — Lecture Timing (Day-wise)")
        ib("Enter Start and End time for each lecture day. Format: 10:00 or 10:00 AM — both work.")

        detected = {}
        for d in sess_dates:
            dn = d.strftime("%A"); detected[dn] = detected.get(dn,0)+1

        day_timing = {}
        hdr = st.columns([2,1,2,2])
        hdr[0].markdown("**Day**")
        hdr[1].markdown("**Count**")
        hdr[2].markdown("**Start**")
        hdr[3].markdown("**End**")

        for day in ALL_DAYS:
            cnt = detected.get(day, 0)
            cs  = st.columns([2,1,2,2])
            with cs[0]: enabled = st.checkbox(day, value=(cnt>0), key=f"sc_{day}")
            with cs[1]: st.caption(f"🔵 {cnt}" if cnt else "—")
            if enabled:
                with cs[2]:
                    sv = st.text_input("", value="10:00", key=f"sv_{day}",
                                       label_visibility="collapsed")
                with cs[3]:
                    ev = st.text_input("", value="11:00", key=f"ev_{day}",
                                       label_visibility="collapsed")
                sh,sm = parse_time(sv); eh,em = parse_time(ev)
                day_timing[day] = (sh,sm,eh,em)

        st.divider()

        # ── STEP D: Units & Paste Syllabus ────────────────────────────
        sec("Step 4 — Units & Syllabus Content")

        ib(
            "Enter how many units your course has. "
            "For each unit, paste or type the syllabus topics — "
            "<b>separated by commas, newlines, or semicolons</b>. "
            "The app will split them into individual session titles automatically."
        )

        n_units = st.number_input(
            "How many units does your course have?",
            min_value=1, max_value=20, value=4
        )

        unit_configs = []
        total_topics_entered = 0

        for i in range(n_units):
            with st.expander(f"Unit {i+1}", expanded=True):
                u1, u2 = st.columns([3, 2])

                with u1:
                    mod_name = st.text_input(
                        "Module Name*  (appears in 'Module Name*' column of sheet)",
                        key=f"mn_{i}",
                        placeholder=f"e.g. Unit {i+1}: Introduction to Contract Law"
                    )

                with u2:
                    def_tlo  = [all_tlos[i % len(all_tlos)]] if all_tlos else [f"{tlo_pfx}1"]
                    sel_tlos = st.multiselect(
                        "TLOs for this unit",
                        options=all_tlos,
                        default=def_tlo,
                        key=f"tl_{i}",
                        help="Multiple TLOs → 'TLO1 | TLO2' in output"
                    )
                    tlo_str = " | ".join(sel_tlos) if sel_tlos else all_tlos[i % len(all_tlos)]

                st.markdown("**Paste Syllabus Topics for this Unit:**")
                st.caption(
                    "Paste topics separated by commas or new lines. Example:\n"
                    "Definition of Contract, Offer and Acceptance, Consideration, "
                    "Capacity to Contract, Free Consent"
                )

                raw_text = st.text_area(
                    "",
                    key=f"syl_{i}",
                    height=120,
                    label_visibility="collapsed",
                    placeholder=(
                        "Paste topics here...\n"
                        "e.g.:\n"
                        "Definition of Indemnity, Rights of Indemnity Holder, "
                        "Position under English Law, Definition of Guarantee, "
                        "Nature and Extent of Surety's Liability, "
                        "Discharge of Surety from Liability"
                    )
                )

                # Parse topics from pasted text
                topics = []
                if raw_text.strip():
                    # Split by comma, semicolon, or newline
                    raw_split = re.split(r"[,;\n]+", raw_text)
                    for item in raw_split:
                        t = item.strip().strip("•-–—*").strip()
                        # Remove numbering like "1." or "1)" at start
                        t = re.sub(r"^\d+[\.\)]\s*", "", t).strip()
                        if t and len(t) > 3:
                            topics.append(t)

                if topics:
                    st.caption(f"✅ {len(topics)} topics parsed from your input:")
                    # Show as compact chips
                    chips = "  ·  ".join(f"`{t}`" for t in topics[:8])
                    if len(topics) > 8:
                        chips += f"  ...+{len(topics)-8} more"
                    st.markdown(chips)
                else:
                    st.caption("⚠️ No topics entered yet — paste or type topics above.")

                total_topics_entered += len(topics)
                unit_configs.append({
                    "module_name": mod_name or f"Unit {i+1}",
                    "tlo":         tlo_str,
                    "topics":      topics,
                })

        # Summary
        if total_topics_entered:
            st.markdown("<br>",unsafe_allow_html=True)
            sum_c = st.columns(3)
            with sum_c[0]:
                st.markdown(
                    f'<div class="sc"><div class="sn">{n_units}</div>'
                    f'<div class="sl">Units</div></div>', unsafe_allow_html=True)
            with sum_c[1]:
                st.markdown(
                    f'<div class="sc"><div class="sn">{total_topics_entered}</div>'
                    f'<div class="sl">Topics Entered</div></div>', unsafe_allow_html=True)
            with sum_c[2]:
                st.markdown(
                    f'<div class="sc"><div class="sn">{total_sessions}</div>'
                    f'<div class="sl">Sessions to Generate</div></div>', unsafe_allow_html=True)
            st.markdown("<br>",unsafe_allow_html=True)

            if total_topics_entered == total_sessions:
                ok("Perfect match — one topic per session.")
            elif total_topics_entered < total_sessions:
                wn(
                    f"Topics entered ({total_topics_entered}) < Sessions ({total_sessions}). "
                    f"Last {total_sessions - total_topics_entered} sessions "
                    f"will be revision/repeat entries."
                )
            else:
                wn(
                    f"Topics entered ({total_topics_entered}) > Sessions ({total_sessions}). "
                    f"Topics will be trimmed proportionally across units."
                )

        st.divider()

        # ── STEP E: Build Preview ─────────────────────────────────────
        sec("Step 5 — Preview & Edit")
        ib(
            "Click <b>Build Preview</b> to generate all rows. "
            "Topics are distributed across sessions proportionally per unit. "
            "Descriptions are auto-generated based on each topic title. "
            "Every cell is editable in the table below."
        )

        if st.button("🔄  Build Preview Table", type="secondary",
                     use_container_width=True):
            if not faculty_id.strip():
                er("Enter Faculty Registration ID in Step 1.")
            elif total_topics_entered == 0:
                er("Paste syllabus topics for at least one unit in Step 4.")
            else:
                flat = balance_topics(unit_configs, total_sessions)
                use_d = sess_dates[:total_sessions] if sess_dates else []
                rows  = []

                for idx in range(total_sessions):
                    if idx < len(use_d):
                        d   = use_d[idx]
                        dn  = d.strftime("%A")
                        sh,sm,eh,em = day_timing.get(dn, (10,0,11,0))
                        sdt = fmt_dt(d,sh,sm)
                        edt = fmt_dt(d,eh,em)
                    else:
                        sdt = edt = "TBD"

                    sess = flat[idx] if idx < len(flat) else {
                        "module": unit_configs[-1]["module_name"] if unit_configs else "Module",
                        "tlo":    f"{tlo_pfx}1",
                        "title":  f"Extra Session {idx+1}"
                    }

                    rows.append({
                        "Sr":               idx+1,
                        "Module Name*":     sess["module"],
                        "Start Date Time*": sdt,
                        "End Date Time*":   edt,
                        "Title*":           sess["title"],
                        "Description*":     auto_desc(sess["title"], sess["module"]),
                        "Mandatory*":       mandatory,
                        "TLO":              sess["tlo"],
                        "Faculty Reg ID*":  faculty_id.strip(),
                    })

                st.session_state.sess_preview = pd.DataFrame(rows)
                st.session_state.sess_results = None
                ok(f"Preview ready — <b>{len(rows)} rows</b>. "
                   "Edit any cell before generating.")

        if st.session_state.sess_preview is not None:
            edited = st.data_editor(
                st.session_state.sess_preview,
                use_container_width=True, num_rows="fixed",
                hide_index=True, key="ed_sess",
                column_config={
                    "Sr":               st.column_config.NumberColumn(
                                            "Sr", width=55, disabled=True),
                    "Module Name*":     st.column_config.TextColumn(
                                            "Module Name* ✏️", width=230),
                    "Start Date Time*": st.column_config.TextColumn(
                                            "Start DateTime* ✏️", width=168),
                    "End Date Time*":   st.column_config.TextColumn(
                                            "End DateTime* ✏️", width=168),
                    "Title*":           st.column_config.TextColumn(
                                            "Title* ✏️", width=265),
                    "Description*":     st.column_config.TextColumn(
                                            "Description* ✏️", width=380),
                    "Mandatory*":       st.column_config.SelectboxColumn(
                                            "Mandatory*",
                                            options=["TRUE","FALSE"], width=110),
                    "TLO":              st.column_config.TextColumn(
                                            "TLO ✏️", width=125),
                    "Faculty Reg ID*":  st.column_config.TextColumn(
                                            "Faculty Reg ID* ✏️", width=168),
                },
                height=480,
            )
            st.session_state.sess_preview = edited

            st.divider()

            if st.button("⚡  Generate session_sheet.xlsx",
                         type="primary", use_container_width=True):
                df = st.session_state.sess_preview
                rows_out = [{
                    "module":      str(r.get("Module Name*","")),
                    "start_dt":    str(r.get("Start Date Time*","")),
                    "end_dt":      str(r.get("End Date Time*","")),
                    "title":       str(r.get("Title*","")),
                    "description": str(r.get("Description*","")),
                    "mandatory":   str(r.get("Mandatory*","TRUE")),
                    "tlo":         str(r.get("TLO",f"{tlo_pfx}1")),
                    "faculty_id":  str(r.get("Faculty Reg ID*", faculty_id.strip())),
                } for _,r in df.iterrows()]

                xlsx = gen_session_sheet(rows_out)
                ok(
                    f"Session sheet generated — <b>{len(rows_out)} rows</b> "
                    f"in exact DL.xlsx format ✅"
                )

                st.download_button(
                    "📋  Download session_sheet.xlsx",
                    data=xlsx,
                    file_name="session_sheet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    st.markdown("""
    <div class="footer">
      Created by <b>Dr. Amar Shukla</b> · IILM University, Gurugram ·
      All processing in-memory — no data stored on server
    </div>""", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
