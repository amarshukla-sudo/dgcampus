"""
DG Sheet Generator v7.0
========================
• Month selection: actual months from attendance sheet shown as checkboxes
  → user picks which months → dates auto-collected → session count auto-set
• Syllabus per unit: Full Unit (all topics) OR Custom (pick specific topics)
• Timing: 7 days, AM/PM dropdowns
• Auto-balance selected topics across selected sessions
• Auto-generate descriptions
• Day-wise attendance fixed
"""

import io, re, zipfile, warnings
from copy import copy
from collections import defaultdict
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

st.set_page_config(page_title="DG Sheet Generator", page_icon="📋",
                   layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
.main{padding:.4rem 1.2rem}
.ttl{background:linear-gradient(135deg,#1a2f5a,#2471a3);color:#fff;
     padding:1rem 2rem;border-radius:12px;text-align:center;margin-bottom:.8rem}
.ttl h1{margin:0;font-size:1.5rem;font-weight:800}
.ttl p{margin:.2rem 0 0;opacity:.88;font-size:.82rem}
.sh{font-size:.93rem;font-weight:700;color:#1a2f5a;
    border-bottom:2px solid #2471a3;padding-bottom:.3rem;margin:.8rem 0 .5rem}
.ib{background:#eaf4fb;border-left:4px solid #2471a3;padding:.4rem .85rem;
    border-radius:0 6px 6px 0;margin:.25rem 0;font-size:.81rem}
.sb{background:#d4edda;border-left:4px solid #28a745;padding:.4rem .85rem;
    border-radius:0 6px 6px 0;margin:.25rem 0;font-size:.81rem}
.wbox{background:#fff8e1;border-left:4px solid #f39c12;padding:.4rem .85rem;
    border-radius:0 6px 6px 0;margin:.25rem 0;font-size:.81rem}
.ebox{background:#fde8e8;border-left:4px solid #e74c3c;padding:.4rem .85rem;
    border-radius:0 6px 6px 0;margin:.25rem 0;font-size:.81rem}
.month-card{background:#f8fafd;border:1px solid #c8ddf0;border-radius:8px;
            padding:.6rem .8rem;margin-bottom:.4rem;text-align:center}
.month-on{background:#e3f2fd;border:2px solid #2471a3;border-radius:8px;
          padding:.6rem .8rem;margin-bottom:.4rem;text-align:center}
.date-pill{display:inline-block;background:#e8f4fd;border:1px solid #aed6f1;
           border-radius:12px;padding:.1rem .45rem;margin:.15rem;font-size:.72rem;color:#1a5276}
.bk{background:#d4edda;color:#155724;padding:.1rem .5rem;
    border-radius:9px;font-size:.73rem;font-weight:700}
.stDownloadButton>button{background:#1a2f5a !important;color:#fff !important;
    border-radius:7px !important;font-weight:600 !important;width:100% !important}
.stDownloadButton>button:hover{background:#2471a3 !important}
</style>
""", unsafe_allow_html=True)

def ib(m):  st.markdown(f'<div class="ib">ℹ️ {m}</div>',   unsafe_allow_html=True)
def sb(m):  st.markdown(f'<div class="sb">✅ {m}</div>',   unsafe_allow_html=True)
def wbox(m):st.markdown(f'<div class="wbox">⚠️ {m}</div>', unsafe_allow_html=True)
def ebox(m):st.markdown(f'<div class="ebox">❌ {m}</div>', unsafe_allow_html=True)
def sh(t):  st.markdown(f'<div class="sh">{t}</div>',      unsafe_allow_html=True)

ALL_DAYS  = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
DAY_SHORT = {d:d[:3] for d in ALL_DAYS}

# ─────────────────────────────────────────────────────────────────────
# CORE UTILITIES
# ─────────────────────────────────────────────────────────────────────

def parse_date(raw) -> date | None:
    if raw is None: return None
    try:
        if pd.isna(raw): return None
    except: pass
    if isinstance(raw, datetime): return raw.date()
    if isinstance(raw, date):     return raw
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


def to_24h(h: int, m: str, ampm: str) -> str:
    hh = int(h)
    if ampm == "PM" and hh != 12: hh += 12
    if ampm == "AM" and hh == 12: hh = 0
    return f"{hh:02d}:{m}"


def best_sheet(fb: bytes, hints=None) -> pd.DataFrame:
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


# ─────────────────────────────────────────────────────────────────────
# PARSE ATTENDANCE
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
        raise ValueError("Dates not found in attendance sheet.")

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
        rs  = " ".join(str(v).lower() for v in row if pd.notna(v))
        if any(k in rs for k in ["name","enrol","roll","sr."]):
            for j,v in enumerate(row):
                s = str(v).lower().strip()
                if "name" in s and name_col is None:      name_col   = j
                elif "enrol" in s and enroll_col is None: enroll_col = j
            break

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

    return {"dates":dates_ordered,"students":students}


def group_dates_by_month(dates: list) -> dict:
    """Returns OrderedDict: { (year,month): [date,...] }"""
    result = defaultdict(list)
    for d in dates:
        if d: result[(d.year, d.month)].append(d)
    return dict(sorted(result.items()))


# ─────────────────────────────────────────────────────────────────────
# PARSE SYLLABUS
# ─────────────────────────────────────────────────────────────────────

def parse_syllabus(fb: bytes, filename: str) -> dict:
    """Returns {unit_name: [topic, ...]}"""
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
                if re.search(r"\sv\.\s|\bAIR\b|\bILR\b|\bSCC\b|https?://|\(\d{4}\)\s+\d|\bLL\s+\(", t): continue
                if re.match(r"^[A-Z][a-zA-Z\s]+\s+v\.\s+[A-Z]", t): continue
                if len(t) > 5: result[cur].append(t)

    elif ext in ("xlsx","xls"):
        df_raw = best_sheet(fb)
        cur = "Module 1"; result[cur] = []
        for _, row in df_raw.iterrows():
            for val in row:
                s = str(val).strip()
                if not s or s.lower() in ("nan","none",""): continue
                if re.match(r"^(UNIT|MODULE|CHAPTER)\s+\d+", s, re.IGNORECASE):
                    cur = s; result[cur] = []; break
                elif len(s) > 5: result[cur].append(s); break

    elif ext == "txt":
        cur = "Module 1"; result[cur] = []
        for line in fb.decode("utf-8","ignore").splitlines():
            t = line.strip()
            if not t: continue
            if re.match(r"^(UNIT|MODULE)\s+\d+", t, re.IGNORECASE):
                cur = t; result[cur] = []; continue
            if len(t) > 5: result[cur].append(t)

    return {k:v for k,v in result.items() if v}


# ─────────────────────────────────────────────────────────────────────
# AUTO-DESCRIPTION
# ─────────────────────────────────────────────────────────────────────

def auto_desc(title: str, module: str) -> str:
    t   = title.strip()
    mod = re.sub(r"^UNIT\s+\d+:\s*","",module,flags=re.IGNORECASE).strip()
    tl  = t.lower()
    if any(w in tl for w in ["definition","define","meaning of"]):
        core = re.sub(r"^(definition\s+(of\s+)?|define\s+)","",t,flags=re.IGNORECASE).strip()
        return (f"This session covers the statutory definition and essential elements of {core}. "
                f"Students will examine the legal provisions, judicial interpretations, and practical significance within the framework of {mod}.")
    if any(w in tl for w in ["rights","duties","liability","liabilities","obligation"]):
        return (f"This session examines the rights, duties, and liabilities arising under {t}. "
                f"Students will analyse relevant statutory provisions, landmark judgments, and the legal consequences for parties involved in {mod}.")
    if any(w in tl for w in ["distinction","difference","compare","vs","versus"]):
        return (f"This session provides a comparative analysis of {t}. "
                f"Students will identify key distinctions through statutory provisions and case-law to apply differential reasoning in {mod}.")
    if any(w in tl for w in ["type","kind","classif","categor"]):
        return (f"This session classifies the different types and categories under {t}. "
                f"Students will study the legal significance of each category and their application within {mod}.")
    if any(w in tl for w in ["termination","discharge","revocation","dissolution"]):
        return (f"This session covers the modes of {t} and their legal consequences. "
                f"Students will study statutory provisions, conditions, and judicial precedents governing this aspect of {mod}.")
    if any(w in tl for w in ["creation","formation","essential","element"]):
        return (f"This session discusses the process of {t} and the requisite legal elements. "
                f"Students will examine statutory requirements, judicial interpretations, and practical illustrations relevant to {mod}.")
    if any(w in tl for w in ["remedy","remedies","damages"]):
        return (f"This session discusses available remedies in cases involving {t}. "
                f"Students will examine statutory and equitable remedies, judicial approaches, and computation of relief in {mod}.")
    if any(w in tl for w in ["nature","scope","concept"]):
        return (f"This session explores the nature, scope, and conceptual framework of {t}. "
                f"Students will critically analyse the theoretical underpinnings and legislative intent within {mod}.")
    if any(w in tl for w in ["case","judgment","v."]):
        return (f"This session analyses {t} as a landmark judicial decision. "
                f"Students will examine the facts, legal issues, reasoning of the court, and the precedential value of this ruling in {mod}.")
    return (f"This session provides a comprehensive study of {t} within the domain of {mod}. "
            f"Students will analyse relevant statutory provisions, judicial precedents, and practical applications through structured discussion and case-based learning.")


# ─────────────────────────────────────────────────────────────────────
# BALANCE SELECTED TOPICS across N sessions
# ─────────────────────────────────────────────────────────────────────

def balance_topics(unit_topics: list, total_sessions: int) -> list:
    """
    unit_topics: [{"unit": str, "tlo": str, "topics": [str,...]}, ...]
    Returns list of {"module","tlo","title"} length = total_sessions
    """
    # Flatten all topics respecting unit order
    all_items = []
    for ut in unit_topics:
        for t in ut["topics"]:
            all_items.append({"module": ut["unit"], "tlo": ut["tlo"], "title": t})

    n_avail = len(all_items)
    result  = []

    if n_avail == 0:
        return [{"module":"Session","tlo":"TLO1","title":f"Session {i+1}"}
                for i in range(total_sessions)]

    if n_avail <= total_sessions:
        # Use all, then fill remaining with proportional repeats
        result = list(all_items)
        extra  = total_sessions - n_avail
        for i in range(extra):
            base = all_items[i % n_avail]
            result.append({"module":base["module"],"tlo":base["tlo"],
                           "title":base["title"] + f" (Revision {i//n_avail+2})"})
    else:
        # Trim proportionally per unit
        total_avail = n_avail
        for ut in unit_topics:
            n_take = round(total_sessions * len(ut["topics"]) / total_avail)
            n_take = max(1, min(n_take, len(ut["topics"])))
            for t in ut["topics"][:n_take]:
                result.append({"module": ut["unit"], "tlo": ut["tlo"], "title": t})
        # Adjust to exactly total_sessions
        while len(result) < total_sessions:
            result.append(result[-1] | {"title": result[-1]["title"]+" (Cont.)"})
        result = result[:total_sessions]

    return result


# ─────────────────────────────────────────────────────────────────────
# GENERATE EXCEL FILES
# ─────────────────────────────────────────────────────────────────────

HF = PatternFill("solid", start_color="1F4E79")
HT = Font(bold=True, color="FFFFFF", size=11)
DA = PatternFill("solid", start_color="D6EAF8")
DB = PatternFill("solid", start_color="EBF5FB")
TN = Side(style="thin", color="BBBBBB")
BD = Border(left=TN,right=TN,top=TN,bottom=TN)
CT = Alignment(horizontal="center",vertical="center",wrap_text=True)
LF = Alignment(horizontal="left",  vertical="center",wrap_text=True)

HEADERS=["Module Name*","Start Date Time*","End Date Time*","Title*",
         "Description*","Attendance Mandatory*","TLO","Teaching Faculty Registration ID*"]
WIDTHS =[30,22,22,42,58,22,18,32]


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
        if "email" in hv: ec = c
        elif "registration" in hv or ("reg" in hv and "id" in hv): rc = c
        elif "attendance" in hv: ac = c
    if not all([ec,rc,ac]): ec,rc,ac = 1,2,3

    enroll_att = {str(st["enrollment"]).strip(): (1 if st["att"].get(session_date)==1 else 0)
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
        for fn,fb in daywise.items(): zf.writestr(f"attendance/{fn}", fb)
    buf.seek(0); return buf.getvalue()


# ═════════════════════════════════════════════════════════════════════
#                          MAIN UI
# ═════════════════════════════════════════════════════════════════════

def main():
    st.markdown("""
    <div class="ttl">
      <h1>📋 DG Session & Attendance Sheet Generator</h1>
      <p>Months from attendance → Unit topics (full/custom) → Auto-balance → Generate</p>
    </div>""", unsafe_allow_html=True)

    for k in ["att_data","syl_data","preview_df","results","_att_nm","_syl_nm"]:
        if k not in st.session_state: st.session_state[k] = None

    # ═══════════════════════════════════════════════════════════
    # STEP 1 — FILES
    # ═══════════════════════════════════════════════════════════
    sh("📂 Step 1 — Files Upload Karo")
    c1,c2,c3 = st.columns(3)
    with c1:
        att_file = st.file_uploader("📅 Master Attendance Sheet (.xlsx)", type=["xlsx"], key="att")
        if att_file: st.markdown(f'<span class="bk">✅ {att_file.name}</span>', unsafe_allow_html=True)
    with c2:
        syl_file = st.file_uploader("📚 Syllabus (.docx / .xlsx / .txt)", type=["docx","xlsx","txt"], key="syl")
        if syl_file: st.markdown(f'<span class="bk">✅ {syl_file.name}</span>', unsafe_allow_html=True)
    with c3:
        att_tpl = st.file_uploader("🗂️ DG Attendance Template (.xlsx)", type=["xlsx"], key="atpl")
        if att_tpl: st.markdown(f'<span class="bk">✅ {att_tpl.name}</span>', unsafe_allow_html=True)

    # Auto-parse
    if att_file and att_file.name != st.session_state._att_nm:
        try:
            st.session_state.att_data = parse_attendance(att_file.read())
            st.session_state._att_nm  = att_file.name
            st.session_state.preview_df = None; st.session_state.results = None
        except Exception as ex: ebox(f"Attendance error: {ex}")

    if syl_file and syl_file.name != st.session_state._syl_nm:
        try:
            st.session_state.syl_data = parse_syllabus(syl_file.read(), syl_file.name)
            st.session_state._syl_nm  = syl_file.name
        except Exception as ex: wbox(f"Syllabus warning: {ex}")

    att_data   = st.session_state.att_data
    syl_data   = st.session_state.syl_data or {}
    all_dates  = att_data["dates"] if att_data else []
    month_data = group_dates_by_month(all_dates)  # {(y,m): [date,...]}

    if all_dates:
        sb(f"Attendance: <b>{len(all_dates)} total session dates</b> across "
           f"<b>{len(month_data)} months</b> · <b>{len(att_data['students'])} students</b>")
    if syl_data:
        total_t = sum(len(v) for v in syl_data.values())
        sb(f"Syllabus: <b>{len(syl_data)} units</b> · <b>{total_t} topics</b> extracted")

    st.divider()

    # ═══════════════════════════════════════════════════════════
    # STEP 2 — SELECT MONTHS (from attendance)
    # ═══════════════════════════════════════════════════════════
    sh("📅 Step 2 — Kaun se Months ka Session Banana Hai?")

    if not month_data:
        ib("Pehle attendance sheet upload karo — months yahan dikhenge.")
        selected_dates = []
    else:
        ib("Jinhe include karna hai unhe tick karo. Har month mein attendance ke actual dates dikhenge.")

        # Show months as columns (max 6 per row)
        month_keys  = list(month_data.keys())
        month_cols  = st.columns(min(len(month_keys), 4))
        selected_dates = []

        for i,(ym,ds) in enumerate(month_data.items()):
            with month_cols[i % 4]:
                label = datetime(ym[0],ym[1],1).strftime("%B %Y")
                checked = st.checkbox(
                    f"**{label}**",
                    value=True,
                    key=f"month_{ym[0]}_{ym[1]}",
                )
                # Show dates in this month
                date_pills = " ".join(
                    f'<span class="date-pill">{d.strftime("%d")} {d.strftime("%a")}</span>'
                    for d in ds
                )
                st.markdown(
                    f'<div style="font-size:.73rem;color:#555;margin-top:.2rem;">'
                    f'<b>{len(ds)} sessions</b><br>{date_pills}</div>',
                    unsafe_allow_html=True
                )
                if checked:
                    selected_dates.extend(ds)

        # Sort selected dates
        selected_dates = sorted(set(selected_dates))

        if selected_dates:
            sb(f"Selected: <b>{len(selected_dates)} dates</b> "
               f"({selected_dates[0].strftime('%d %b %Y')} → "
               f"{selected_dates[-1].strftime('%d %b %Y')})")
        else:
            wbox("Kam se kam ek month select karo.")

    st.divider()

    # ═══════════════════════════════════════════════════════════
    # STEP 3 — BASIC SETTINGS
    # ═══════════════════════════════════════════════════════════
    sh("⚙️ Step 3 — Basic Settings")

    g1,g2,g3,g4 = st.columns(4)
    with g1:
        # Sessions = selected dates OR manual override
        auto_count = len(selected_dates)
        manual_override = st.checkbox("✏️ Override session count", value=False, key="override")
        if manual_override:
            total_sessions = st.number_input("Total Sessions",
                min_value=1, max_value=500, value=auto_count or 50)
        else:
            total_sessions = auto_count
            st.metric("📊 Total Sessions", total_sessions,
                      help="= number of selected attendance dates")

    with g2:
        faculty_id = st.text_input("🧑‍🏫 Faculty Registration ID *",
                                   placeholder="e.g. IILMGG006412025")
    with g3:
        mandatory = st.selectbox("✅ Attendance Mandatory?", ["TRUE","FALSE"])
    with g4:
        tlo_max = st.number_input("📊 Max TLO Number", min_value=1, max_value=100, value=5)

    all_tlos = [f"TLO{i}" for i in range(1, tlo_max+1)]

    st.divider()

    # ═══════════════════════════════════════════════════════════
    # STEP 4 — TIMING (7 days, AM/PM)
    # ═══════════════════════════════════════════════════════════
    sh("⏰ Step 4 — Lecture Timing (Din ke Hisab se)")
    ib("Jo din class hoti hai unhe enable karo. Baaki din ki dates automatically timing inherit karengi.")

    # Detect which days are in selected_dates
    detected_days = {}
    for d in selected_dates:
        dn = d.strftime("%A")
        detected_days[dn] = detected_days.get(dn,0) + 1

    day_timing = {}  # {day_name: "HH:MM"} for start and end

    cols7 = st.columns(7)
    for i,day in enumerate(ALL_DAYS):
        count = detected_days.get(day,0)
        with cols7[i]:
            enabled = st.checkbox(f"**{DAY_SHORT[day]}**",
                                  value=(count>0), key=f"de_{day}",
                                  help=f"{count} sessions" if count else "No sessions in selection")
            if count:
                st.caption(f"🔵 {count}x")
            else:
                st.caption("—")

            if enabled:
                # Start
                h1,m1,a1 = st.columns([3,3,3])
                with h1: sh_v = st.selectbox("",list(range(1,13)),index=9,key=f"sh_{day}",label_visibility="collapsed")
                with m1: sm_v = st.selectbox("",["00","10","15","20","30","45"],index=0,key=f"sm_{day}",label_visibility="collapsed")
                with a1: sa_v = st.selectbox("",["AM","PM"],index=1,key=f"sa_{day}",label_visibility="collapsed")
                start_24 = to_24h(sh_v, sm_v, sa_v)

                # End
                h2,m2,a2 = st.columns([3,3,3])
                with h2: eh_v = st.selectbox("",list(range(1,13)),index=10,key=f"eh_{day}",label_visibility="collapsed")
                with m2: em_v = st.selectbox("",["00","10","15","20","30","45"],index=0,key=f"em_{day}",label_visibility="collapsed")
                with a2: ea_v = st.selectbox("",["AM","PM"],index=1,key=f"ea_{day}",label_visibility="collapsed")
                end_24 = to_24h(eh_v, em_v, ea_v)

                st.caption(f"⏱ {start_24}–{end_24}")
                day_timing[day] = (start_24, end_24)

    st.divider()

    # ═══════════════════════════════════════════════════════════
    # STEP 5 — UNIT & TOPIC SELECTION
    # ═══════════════════════════════════════════════════════════
    sh("📚 Step 5 — Unit aur Topics Select Karo")

    syl_units = list(syl_data.keys()) if syl_data else []
    unit_topics_final = []  # [{"unit":str,"tlo":str,"topics":[str,...]}]

    if not syl_units:
        wbox("Syllabus upload nahi hua — topics manually preview table mein bharne padenge.")
    else:
        ib("Har unit ke liye: <b>Full Unit</b> (pura syllabus) ya <b>Custom</b> (specific topics choose karo).")

        num_mods = st.selectbox("Kitne Modules?", list(range(1,len(syl_units)+1)),
                                index=len(syl_units)-1,
                                format_func=lambda x:f"{x} Module{'s' if x>1 else ''}")

        for i in range(num_mods):
            with st.expander(
                f"📦 Module {i+1} — "
                f"{syl_units[i][:50] if i<len(syl_units) else f'Module {i+1}'}",
                expanded=True
            ):
                mc1,mc2 = st.columns([2,1])
                with mc1:
                    def_name = syl_units[i] if i<len(syl_units) else f"Module {i+1}"
                    unit_name = st.text_input("Unit Name *", value=def_name, key=f"un_{i}")
                with mc2:
                    def_tlo = [all_tlos[i % len(all_tlos)]] if all_tlos else ["TLO1"]
                    sel_tlos = st.multiselect("TLOs", all_tlos, default=def_tlo, key=f"tl_{i}")
                    tlo_str  = " | ".join(sel_tlos) if sel_tlos else all_tlos[i%len(all_tlos)]

                # Get available topics from syllabus
                unit_key     = syl_units[i] if i < len(syl_units) else ""
                avail_topics = syl_data.get(unit_key, [])

                coverage_mode = st.radio(
                    f"Topics coverage for Module {i+1}:",
                    ["📖 Full Unit (pura syllabus is unit ka)",
                     "✂️ Custom (main sirf kuch topics chahiye)"],
                    key=f"cov_{i}",
                    horizontal=True
                )

                if "Custom" in coverage_mode:
                    if avail_topics:
                        chosen = st.multiselect(
                            f"Topics select karo (Module {i+1}):",
                            options=avail_topics,
                            default=avail_topics,   # all pre-selected, user can deselect
                            key=f"tp_{i}"
                        )
                        st.caption(f"✅ {len(chosen)} of {len(avail_topics)} topics selected")
                        final_topics = chosen
                    else:
                        st.caption("Syllabus mein is unit ke topics nahi mile.")
                        final_topics = []
                else:
                    final_topics = avail_topics
                    if avail_topics:
                        st.caption(f"📖 All {len(avail_topics)} topics included")
                        # Show them collapsed
                        with st.expander(f"See all {len(avail_topics)} topics"):
                            for t in avail_topics:
                                st.caption(f"• {t}")

                unit_topics_final.append({
                    "unit":   unit_name,
                    "tlo":    tlo_str,
                    "topics": final_topics,
                })

        # Summary
        total_selected_topics = sum(len(u["topics"]) for u in unit_topics_final)
        if total_selected_topics > 0 and total_sessions > 0:
            if total_sessions == total_selected_topics:
                sb(f"<b>Perfect match!</b> {total_sessions} sessions = {total_selected_topics} selected topics.")
            elif total_sessions > total_selected_topics:
                wbox(f"{total_sessions} sessions > {total_selected_topics} topics. "
                     f"Last {total_sessions - total_selected_topics} sessions = revision entries.")
            else:
                wbox(f"{total_sessions} sessions < {total_selected_topics} topics. "
                     f"Topics will be trimmed proportionally per unit.")

    st.divider()

    # ═══════════════════════════════════════════════════════════
    # STEP 6 — BUILD PREVIEW TABLE
    # ═══════════════════════════════════════════════════════════
    sh("👀 Step 6 — Preview Table Banao & Edit Karo")
    ib("'Build Preview' dabao → dates + timing + balanced topics + auto-descriptions dikhenge. "
       "Table mein directly edit kar sakte ho before generating.")

    prev_btn = st.button("🔄 Build Preview Table", type="secondary", use_container_width=True)

    if prev_btn:
        if not selected_dates:
            ebox("Pehle months select karo (Step 2).")
        elif not faculty_id.strip():
            ebox("Faculty Registration ID daalo (Step 3).")
        else:
            # Balance topics
            if unit_topics_final:
                flat = balance_topics(unit_topics_final, total_sessions)
            else:
                flat = [{"module":f"Module {(i%4)+1}","tlo":"TLO1","title":f"Session {i+1}"}
                        for i in range(total_sessions)]

            use_dates = selected_dates[:total_sessions]
            rows = []
            for idx in range(total_sessions):
                if idx < len(use_dates):
                    d   = use_dates[idx]
                    ds  = d.strftime("%Y-%m-%d")
                    dn  = d.strftime("%A")
                    s_t,e_t = day_timing.get(dn, ("10:00","11:00"))
                    start_dt = f"{ds} {s_t}:00"
                    end_dt   = f"{ds} {e_t}:00"
                else:
                    start_dt = end_dt = "TBD"

                sess = flat[idx] if idx < len(flat) else {
                    "module":"Extra","tlo":"TLO1","title":f"Extra Session {idx+1}"}

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

            st.session_state.preview_df = pd.DataFrame(rows)
            st.session_state.results    = None

    if st.session_state.preview_df is not None:
        df = st.session_state.preview_df
        sb(f"Preview ready: <b>{len(df)} rows</b> — neeche table mein directly edit karo.")

        edited = st.data_editor(
            df, use_container_width=True, num_rows="fixed",
            hide_index=True, key="prev_ed",
            column_config={
                "Sr":               st.column_config.NumberColumn("Sr",width=50,disabled=True),
                "Module Name*":     st.column_config.TextColumn("Module Name* ✏️",    width=220),
                "Start Date Time*": st.column_config.TextColumn("Start DateTime* ✏️", width=165),
                "End Date Time*":   st.column_config.TextColumn("End DateTime* ✏️",   width=165),
                "Title*":           st.column_config.TextColumn("Title* ✏️",          width=260),
                "Description*":     st.column_config.TextColumn("Description* ✏️",    width=360),
                "Mandatory*":       st.column_config.SelectboxColumn("Mandatory*",
                                        options=["TRUE","FALSE"],width=100),
                "TLO":              st.column_config.TextColumn("TLO ✏️",             width=120),
                "Faculty Reg ID*":  st.column_config.TextColumn("Faculty Reg ID* ✏️", width=165),
            },
            height=460,
        )
        st.session_state.preview_df = edited

    st.divider()

    # ═══════════════════════════════════════════════════════════
    # STEP 7 — GENERATE
    # ═══════════════════════════════════════════════════════════
    sh("🚀 Step 7 — Generate Karo")

    can_gen = bool(
        st.session_state.preview_df is not None and
        att_data and faculty_id.strip()
    )
    if not can_gen:
        miss = []
        if not att_data:                          miss.append("Attendance Sheet")
        if not faculty_id.strip():                miss.append("Faculty Registration ID")
        if st.session_state.preview_df is None:  miss.append("Step 6 mein Preview Table banao")
        if miss: ib(f"Pehle karo: <b>{', '.join(miss)}</b>")

    gen_btn = st.button("⚡ Generate Session Sheet + All Attendance Files",
                        type="primary", disabled=not can_gen, use_container_width=True)

    if gen_btn and can_gen:
        errors=[]; prog=st.progress(0); log=st.empty()
        try:
            df = st.session_state.preview_df
            log.markdown('<div class="ib">📋 Session rows preparing…</div>', unsafe_allow_html=True)
            session_rows = [{
                "module":      str(r.get("Module Name*","")),
                "start_dt":    str(r.get("Start Date Time*","")),
                "end_dt":      str(r.get("End Date Time*","")),
                "title":       str(r.get("Title*","")),
                "description": str(r.get("Description*","")),
                "mandatory":   str(r.get("Mandatory*","TRUE")),
                "tlo":         str(r.get("TLO","TLO1")),
                "faculty_id":  str(r.get("Faculty Reg ID*", faculty_id.strip())),
            } for _,r in df.iterrows()]
            prog.progress(20)

            log.markdown('<div class="ib">📊 Session sheet generating…</div>', unsafe_allow_html=True)
            session_xlsx = gen_session_sheet(session_rows)
            sb(f"Session sheet: <b>{len(session_rows)} rows</b> in DL.xlsx exact format.")
            prog.progress(50)

            daywise = {}
            if att_tpl:
                log.markdown('<div class="ib">🗂️ Day-wise attendance generating…</div>', unsafe_allow_html=True)
                att_tpl_bytes = att_tpl.read()
                bar = st.progress(0)
                for i,d in enumerate(selected_dates[:total_sessions]):
                    try:
                        fname = f"attendance_{d.strftime('%Y-%m-%d')}.xlsx"
                        daywise[fname] = gen_daywise_att(att_tpl_bytes, att_data, d)
                    except Exception as ex: errors.append(f"{d}: {ex}")
                    bar.progress((i+1)/max(total_sessions,1))
                sb(f"Day-wise: <b>{len(daywise)} attendance files</b> ready.")
            else:
                wbox("Attendance template nahi diya — sirf session sheet download hogi.")

            prog.progress(85)
            zip_bytes = build_zip(session_xlsx, daywise)
            prog.progress(100)
            log.markdown('<div class="sb">🎉 Sab ready! Neeche se download karo.</div>', unsafe_allow_html=True)

            st.session_state.results = {
                "session":session_xlsx,"daywise":daywise,
                "zip":zip_bytes,"errors":errors,"n":len(session_rows)
            }
        except Exception as ex:
            ebox(f"Error: {ex}")
            import traceback; st.code(traceback.format_exc())

    # ═══════════════════════════════════════════════════════════
    # STEP 8 — DOWNLOADS
    # ═══════════════════════════════════════════════════════════
    if st.session_state.results:
        res = st.session_state.results
        st.divider()
        sh("📥 Step 8 — Download Karo")
        for e in res.get("errors",[]): wbox(f"Skipped: {e}")

        d1,d2 = st.columns(2)
        with d1:
            st.download_button("📦 ⬇ Download ALL as ZIP",
                data=res["zip"],
                file_name=f"DG_Output_{datetime.now():%Y%m%d_%H%M}.zip",
                mime="application/zip", use_container_width=True)
        with d2:
            st.download_button("📋 ⬇ Download session_sheet.xlsx",
                data=res["session"], file_name="session_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        if res["daywise"]:
            st.markdown(f"#### 📅 Day-wise Attendance ({len(res['daywise'])} files)")
            files = list(res["daywise"].items())
            for row_i in range(0,len(files),4):
                chunk = files[row_i:row_i+4]; cols = st.columns(4)
                for ci,(fname,fb) in enumerate(chunk):
                    with cols[ci]:
                        dp = fname.replace("attendance_","").replace(".xlsx","")
                        try:    label = datetime.strptime(dp,"%Y-%m-%d").strftime("%d %b %Y")
                        except: label = dp
                        st.download_button(f"📅 {label}",data=fb,file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,key=f"dl_{fname}")

        st.divider()
        if st.button("🔄 Reset"):
            for k in ["att_data","syl_data","preview_df","results","_att_nm","_syl_nm"]:
                st.session_state[k] = None
            st.rerun()

    st.markdown("<center style='color:#aaa;font-size:.7rem;margin-top:.8rem'>"
                "🔒 In-memory processing — no data stored on server</center>",
                unsafe_allow_html=True)


if __name__ == "__main__":
    main()
