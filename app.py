"""
DG Sheet Generator v6.0
========================
Fixes:
  1. Timing: 7 days checkboxes, AM/PM format, enable/disable per day
  2. Date range: From Month → To Month filter on attendance dates
  3. Unit coverage: Full syllabus OR partial (user picks which units)
  4. Sessions = auto-count from filtered dates OR user override
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

st.set_page_config(page_title="DG Sheet Generator", page_icon="📋",
                   layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
.main{padding:.4rem 1.2rem}
.ttl{background:linear-gradient(135deg,#1a2f5a,#2471a3);color:#fff;
     padding:1rem 2rem;border-radius:12px;text-align:center;margin-bottom:.8rem}
.ttl h1{margin:0;font-size:1.55rem;font-weight:800}
.ttl p{margin:.2rem 0 0;opacity:.88;font-size:.83rem}
.sh{font-size:.93rem;font-weight:700;color:#1a2f5a;
    border-bottom:2px solid #2471a3;padding-bottom:.3rem;margin:.8rem 0 .5rem}
.ib{background:#eaf4fb;border-left:4px solid #2471a3;padding:.4rem .8rem;
    border-radius:0 6px 6px 0;margin:.25rem 0;font-size:.81rem}
.sb{background:#d4edda;border-left:4px solid #28a745;padding:.4rem .8rem;
    border-radius:0 6px 6px 0;margin:.25rem 0;font-size:.81rem}
.wb{background:#fff8e1;border-left:4px solid #f39c12;padding:.4rem .8rem;
    border-radius:0 6px 6px 0;margin:.25rem 0;font-size:.81rem}
.eb{background:#fde8e8;border-left:4px solid #e74c3c;padding:.4rem .8rem;
    border-radius:0 6px 6px 0;margin:.25rem 0;font-size:.81rem}
.day-card{background:#f8fafd;border:1px solid #d0dff0;border-radius:8px;
          padding:.6rem .8rem;margin-bottom:.4rem}
.day-on{background:#e8f4fd;border:1px solid #2471a3;border-radius:8px;
        padding:.6rem .8rem;margin-bottom:.4rem}
.bk{background:#d4edda;color:#155724;padding:.1rem .5rem;
    border-radius:9px;font-size:.73rem;font-weight:700}
.stDownloadButton>button{background:#1a2f5a !important;color:#fff !important;
    border-radius:7px !important;font-weight:600 !important;width:100% !important}
.stDownloadButton>button:hover{background:#2471a3 !important}
</style>
""", unsafe_allow_html=True)

def ib(m): st.markdown(f'<div class="ib">ℹ️ {m}</div>', unsafe_allow_html=True)
def sb(m): st.markdown(f'<div class="sb">✅ {m}</div>', unsafe_allow_html=True)
def wb_msg(m): st.markdown(f'<div class="wb">⚠️ {m}</div>', unsafe_allow_html=True)
def eb(m): st.markdown(f'<div class="eb">❌ {m}</div>', unsafe_allow_html=True)
def sh(t): st.markdown(f'<div class="sh">{t}</div>', unsafe_allow_html=True)

ALL_DAYS   = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
DAY_SHORT  = {"Monday":"Mon","Tuesday":"Tue","Wednesday":"Wed","Thursday":"Thu",
              "Friday":"Fri","Saturday":"Sat","Sunday":"Sun"}
DAY_IDX    = {d:i for i,d in enumerate(ALL_DAYS)}
MONTHS_ALL = ["January","February","March","April","May","June",
              "July","August","September","October","November","December"]

# ─────────────────────────────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────────────────────────────

def parse_date(raw) -> date | None:
    if raw is None: return None
    try:
        if pd.isna(raw): return None
    except: pass
    if isinstance(raw, datetime): return raw.date()
    if isinstance(raw, date): return raw
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


def to_24h(h_str: str, m_str: str, ampm: str) -> str:
    """Convert h, m, AM/PM → HH:MM (24hr)"""
    try:
        h = int(h_str); m = int(m_str)
        if ampm == "PM" and h != 12: h += 12
        if ampm == "AM" and h == 12: h = 0
        return f"{h:02d}:{m:02d}"
    except: return "10:00"


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
        n = int(df.notna().sum().sum())
        if n>most: most,best=n,s
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
            date_row_idx=i; date_col_start=first; break
    if date_row_idx is None:
        raise ValueError("Attendance sheet mein dates nahi mili.")

    date_row = df_raw.iloc[date_row_idx]
    col_to_date, dates_ordered, seen = {}, [], set()
    for j in range(date_col_start, df_raw.shape[1]):
        d = parse_date(date_row.iloc[j])
        if d:
            col_to_date[j]=d
            if d not in seen: dates_ordered.append(d); seen.add(d)

    name_col=enroll_col=None
    for i in range(date_row_idx):
        row=df_raw.iloc[i]
        rs=" ".join(str(v).lower() for v in row if pd.notna(v))
        if any(k in rs for k in ["name","enrol","roll","sr."]):
            for j,v in enumerate(row):
                s=str(v).lower().strip()
                if "name" in s and name_col is None:      name_col=j
                elif "enrol" in s and enroll_col is None: enroll_col=j
            break

    meta={}
    for i in range(date_row_idx):
        row=df_raw.iloc[i]
        for j in range(len(row)-1):
            k,v=str(row.iloc[j]).strip(),str(row.iloc[j+1]).strip()
            if k not in ("nan","") and v not in ("nan",""): meta[k]=v

    students=[]
    for i in range(date_row_idx+1, df_raw.shape[0]):
        row=df_raw.iloc[i]
        if row.notna().sum()<3: continue
        name=str(row.iloc[name_col]).strip() if name_col is not None and pd.notna(row.iloc[name_col]) else ""
        if not name or name.lower() in ("nan","none",""): continue
        enroll=str(row.iloc[enroll_col]).strip() if enroll_col is not None and pd.notna(row.iloc[enroll_col]) else ""
        att={}
        for col_j,d in col_to_date.items():
            val=row.iloc[col_j]
            try:    att[d]=int(float(val)) if not pd.isna(val) else None
            except: att[d]=None
        students.append({"name":name,"enrollment":enroll,"att":att})

    return {"dates":dates_ordered,"students":students,"meta":meta}


# ─────────────────────────────────────────────────────────────────────
# PARSE SYLLABUS
# ─────────────────────────────────────────────────────────────────────

def parse_syllabus(fb: bytes, filename: str) -> dict:
    ext=filename.lower().rsplit(".",1)[-1]
    result={}
    if ext=="docx":
        import docx as _d
        doc=_d.Document(io.BytesIO(fb))
        cur=None
        for p in doc.paragraphs:
            t=p.text.strip()
            if not t: continue
            if re.match(r"^UNIT\s+\d+",t,re.IGNORECASE):
                cur=t; result[cur]=[]; continue
            if cur:
                if re.match(r"^Case.?law",t,re.IGNORECASE): continue
                if re.search(r"\sv\.\s|\bAIR\b|\bILR\b|\bSCC\b|https?://|\(\d{4}\)\s+\d|\bLL\s+\(",t): continue
                if re.match(r"^[A-Z][a-zA-Z\s]+\s+v\.\s+[A-Z]",t): continue
                if len(t)>5: result[cur].append(t)
    elif ext in ("xlsx","xls"):
        df_raw=best_sheet(fb)
        cur="Module 1"; result[cur]=[]
        for i,row in df_raw.iterrows():
            for _,val in enumerate(row):
                s=str(val).strip()
                if not s or s.lower() in ("nan","none",""): continue
                if re.match(r"^(UNIT|MODULE|CHAPTER)\s+\d+",s,re.IGNORECASE):
                    cur=s; result[cur]=[]; break
                elif len(s)>5: result[cur].append(s); break
    elif ext=="txt":
        cur="Module 1"; result[cur]=[]
        for line in fb.decode("utf-8","ignore").splitlines():
            t=line.strip()
            if not t: continue
            if re.match(r"^(UNIT|MODULE)\s+\d+",t,re.IGNORECASE):
                cur=t; result[cur]=[]; continue
            if len(t)>5: result[cur].append(t)
    return {k:v for k,v in result.items() if v}


# ─────────────────────────────────────────────────────────────────────
# AUTO-BALANCE TITLES
# ─────────────────────────────────────────────────────────────────────

def balance_titles(syl_data: dict, modules: list, total_sessions: int) -> list:
    mod_data=[]
    for m in modules:
        raw=syl_data.get(m.get("unit_key",""),[])
        clean=[t for t in raw if len(t)>5]
        mod_data.append({"module":m["name"],"tlo":m["tlo"],"titles":clean})

    n_mods=len(mod_data)
    if not n_mods: return []

    total_avail=sum(len(m["titles"]) for m in mod_data)
    alloc=[]
    for m in mod_data:
        share=round(total_sessions*len(m["titles"])/total_avail) if total_avail>0 else total_sessions//n_mods
        alloc.append(max(1,share))

    while sum(alloc)<total_sessions: alloc[alloc.index(min(alloc))]+=1
    while sum(alloc)>total_sessions: alloc[alloc.index(max(alloc))]-=1

    result=[]
    for m,n in zip(mod_data,alloc):
        titles=m["titles"] or [f"Session – {m['module']}"]
        for i in range(n):
            t=titles[i] if i<len(titles) else f"{titles[-1]} (Part {i-len(titles)+2})"
            result.append({"module":m["module"],"tlo":m["tlo"],"title":t})
    return result[:total_sessions]


# ─────────────────────────────────────────────────────────────────────
# AUTO-GENERATE DESCRIPTION
# ─────────────────────────────────────────────────────────────────────

def auto_desc(title: str, module: str) -> str:
    t=title.strip()
    mod=re.sub(r"^UNIT\s+\d+:\s*","",module,flags=re.IGNORECASE).strip()
    tl=t.lower()
    if any(w in tl for w in ["definition","define","meaning of"]):
        core=re.sub(r"^(definition\s+(of\s+)?|define\s+)","",t,flags=re.IGNORECASE).strip()
        return (f"This session covers the statutory definition and essential elements of {core}. "
                f"Students will examine the legal provisions, judicial interpretations, and practical significance within the framework of {mod}.")
    if any(w in tl for w in ["rights","duties","liability","liabilities","obligation"]):
        return (f"This session examines the rights, duties, and liabilities arising under {t}. "
                f"Students will analyse relevant statutory provisions, landmark judgments, and the legal consequences for parties involved in {mod}.")
    if any(w in tl for w in ["distinction","difference","compare","vs","versus"]):
        return (f"This session provides a comparative analysis of {t}. "
                f"Students will identify key distinctions through statutory provisions and case-law, enabling them to apply differential reasoning in {mod}.")
    if any(w in tl for w in ["type","kind","classif","categor","form"]):
        return (f"This session classifies the different types and categories under {t}. "
                f"Students will study the legal significance of each category and their practical application within {mod}.")
    if any(w in tl for w in ["termination","discharge","revocation","dissolution"]):
        return (f"This session covers the modes of {t} and their legal consequences. "
                f"Students will study statutory provisions, conditions, and judicial precedents governing this aspect of {mod}.")
    if any(w in tl for w in ["creation","formation","essential","element","requisite"]):
        return (f"This session discusses the process of {t} and the requisite legal elements. "
                f"Students will examine statutory requirements, judicial interpretations, and practical illustrations relevant to {mod}.")
    if any(w in tl for w in ["remedy","remedies","compensation","damages"]):
        return (f"This session discusses available remedies in cases involving {t}. "
                f"Students will examine statutory and equitable remedies, judicial approaches, and computation of relief in {mod}.")
    if any(w in tl for w in ["nature","scope","extent","concept"]):
        return (f"This session explores the nature, scope, and conceptual framework of {t}. "
                f"Students will critically analyse the theoretical underpinnings and legislative intent governing this concept within {mod}.")
    if any(w in tl for w in ["case","judgment","v."]):
        return (f"This session analyses {t} as a landmark judicial decision. "
                f"Students will examine the facts, legal issues, reasoning of the court, and the precedential value of this ruling in {mod}.")
    return (f"This session provides a comprehensive study of {t} within the domain of {mod}. "
            f"Students will analyse relevant statutory provisions, judicial precedents, and practical applications through structured discussion and case-based learning.")


# ─────────────────────────────────────────────────────────────────────
# GENERATE SESSION SHEET (exact DL.xlsx format)
# ─────────────────────────────────────────────────────────────────────

HF = PatternFill("solid", start_color="1F4E79")
HT = Font(bold=True, color="FFFFFF", size=11)
DA = PatternFill("solid", start_color="D6EAF8")
DB = PatternFill("solid", start_color="EBF5FB")
TN = Side(style="thin", color="BBBBBB")
BD = Border(left=TN,right=TN,top=TN,bottom=TN)
CT = Alignment(horizontal="center", vertical="center", wrap_text=True)
LF = Alignment(horizontal="left",   vertical="center", wrap_text=True)
HEADERS=["Module Name*","Start Date Time*","End Date Time*","Title*",
         "Description*","Attendance Mandatory*","TLO","Teaching Faculty Registration ID*"]
WIDTHS =[30,22,22,42,58,22,18,32]


def gen_session_sheet(rows: list) -> bytes:
    wb=Workbook(); ws=wb.active; ws.title="Session Sheet"
    aligns=[LF,CT,CT,LF,LF,CT,CT,CT]
    for c,(h,w) in enumerate(zip(HEADERS,WIDTHS),1):
        cell=ws.cell(1,c,h)
        cell.fill=HF; cell.font=HT; cell.border=BD; cell.alignment=CT
        ws.column_dimensions[get_column_letter(c)].width=w
    ws.row_dimensions[1].height=24
    for i,row in enumerate(rows,2):
        fill=DA if i%2==0 else DB
        vals=[row.get("module",""),row.get("start_dt",""),row.get("end_dt",""),
              row.get("title",""),row.get("description",""),row.get("mandatory","TRUE"),
              row.get("tlo","TLO1"),row.get("faculty_id","")]
        for c,(v,al) in enumerate(zip(vals,aligns),1):
            cell=ws.cell(i,c,v)
            cell.fill=fill; cell.border=BD; cell.alignment=al; cell.font=Font(size=10)
        ws.row_dimensions[i].height=30
    ws.freeze_panes="A2"
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────
# GENERATE DAY-WISE ATTENDANCE
# ─────────────────────────────────────────────────────────────────────

def gen_daywise_att(att_tpl: bytes, att_data: dict, session_date: date) -> bytes:
    wb=load_workbook(io.BytesIO(att_tpl)); ws=wb.active
    email_col=regid_col=att_col=None
    for c in range(1,ws.max_column+1):
        hv=str(ws.cell(1,c).value or "").lower().strip()
        if "email" in hv:                                   email_col=c
        elif "registration" in hv or ("reg" in hv and "id" in hv): regid_col=c
        elif "attendance" in hv:                            att_col=c
    if not all([email_col,regid_col,att_col]): email_col,regid_col,att_col=1,2,3

    enroll_att={str(st["enrollment"]).strip():(1 if st["att"].get(session_date)==1 else 0)
                for st in att_data["students"]}

    GF=PatternFill("solid",start_color="C6EFCE"); RF=PatternFill("solid",start_color="FFC7CE")
    GT=Font(color="006100",bold=True,size=10);    RT=Font(color="9C0006",bold=True,size=10)
    CA=Alignment(horizontal="center",vertical="center")

    for r in range(2,ws.max_row+1):
        reg=str(ws.cell(r,regid_col).value or "").strip()
        if not reg: continue
        status="PRESENT" if enroll_att.get(reg,0)==1 else "ABSENT"
        cell=ws.cell(r,att_col)
        cell.value=status; cell.fill=GF if status=="PRESENT" else RF
        cell.font=GT if status=="PRESENT" else RT; cell.alignment=CA

    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()


def build_zip(session_bytes: bytes, daywise: dict) -> bytes:
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("session_sheet.xlsx",session_bytes)
        for fname,fb in daywise.items(): zf.writestr(f"attendance/{fname}",fb)
    buf.seek(0); return buf.getvalue()


# ═════════════════════════════════════════════════════════════════════
#                          MAIN UI
# ═════════════════════════════════════════════════════════════════════

def main():
    st.markdown("""
    <div class="ttl">
      <h1>📋 DG Session & Attendance Sheet Generator</h1>
      <p>Date range → Sessions → Timing (AM/PM) → Syllabus balance → Auto descriptions → Download</p>
    </div>""", unsafe_allow_html=True)

    for k in ["att_data","syl_data","preview_df","results","_att_nm","_syl_nm"]:
        if k not in st.session_state: st.session_state[k]=None

    # ═══════════════════════════════════════════════════════════════
    # STEP 1 — FILES
    # ═══════════════════════════════════════════════════════════════
    sh("📂 Step 1 — Files Upload Karo")
    f1,f2,f3=st.columns(3)
    with f1:
        att_file=st.file_uploader("📅 Master Attendance Sheet (.xlsx)",type=["xlsx"],key="att")
        if att_file: st.markdown(f'<span class="bk">✅ {att_file.name}</span>',unsafe_allow_html=True)
    with f2:
        syl_file=st.file_uploader("📚 Syllabus File (.docx/.xlsx/.txt)",type=["docx","xlsx","txt"],key="syl")
        if syl_file: st.markdown(f'<span class="bk">✅ {syl_file.name}</span>',unsafe_allow_html=True)
    with f3:
        att_tpl=st.file_uploader("🗂️ DG Attendance Template (.xlsx)",type=["xlsx"],key="atpl")
        if att_tpl: st.markdown(f'<span class="bk">✅ {att_tpl.name}</span>',unsafe_allow_html=True)

    if att_file and att_file.name!=st.session_state._att_nm:
        try:
            st.session_state.att_data=parse_attendance(att_file.read())
            st.session_state._att_nm=att_file.name
            st.session_state.preview_df=None; st.session_state.results=None
        except Exception as ex: eb(f"Attendance error: {ex}")

    if syl_file and syl_file.name!=st.session_state._syl_nm:
        try:
            st.session_state.syl_data=parse_syllabus(syl_file.read(),syl_file.name)
            st.session_state._syl_nm=syl_file.name
        except Exception as ex: wb_msg(f"Syllabus warning: {ex}")

    att_data=st.session_state.att_data
    syl_data=st.session_state.syl_data or {}
    all_dates=att_data["dates"] if att_data else []

    if all_dates:
        sb(f"Attendance: <b>{len(all_dates)} dates</b> ({all_dates[0].strftime('%d %b %Y')} → "
           f"{all_dates[-1].strftime('%d %b %Y')}) · <b>{len(att_data['students'])} students</b>")
    if syl_data:
        total_t=sum(len(v) for v in syl_data.values())
        sb(f"Syllabus: <b>{len(syl_data)} units</b> · <b>{total_t} topics</b> extracted")

    st.divider()

    # ═══════════════════════════════════════════════════════════════
    # STEP 2 — DATE RANGE + SESSION COUNT
    # ═══════════════════════════════════════════════════════════════
    sh("📅 Step 2 — Session Date Range & Count")
    ib("Kaun se month se kaun se month tak session banana hai? "
       "Attendance sheet mein jo dates hain unhe filter karega.")

    # Get available months from attendance
    avail_months=[]
    seen_ym=set()
    for d in all_dates:
        if d:
            ym=(d.year, d.month)
            if ym not in seen_ym:
                avail_months.append(d); seen_ym.add(ym)

    if avail_months:
        month_labels=[d.strftime("%B %Y") for d in avail_months]

        dr1,dr2,dr3=st.columns(3)
        with dr1:
            start_month_label=st.selectbox("📅 From Month",
                options=month_labels, index=0, key="sm")
        with dr2:
            end_month_label=st.selectbox("📅 To Month",
                options=month_labels, index=len(month_labels)-1, key="em")
        with dr3:
            # Parse selected months
            sm_date=next(d for d in avail_months if d.strftime("%B %Y")==start_month_label)
            em_date=next(d for d in avail_months if d.strftime("%B %Y")==end_month_label)

            # Filter dates in range
            filtered_dates=[
                d for d in all_dates if d and
                (d.year>sm_date.year or (d.year==sm_date.year and d.month>=sm_date.month)) and
                (d.year<em_date.year or (d.year==em_date.year and d.month<=em_date.month))
            ]
            st.metric("📊 Dates in Range", len(filtered_dates))

        if filtered_dates:
            ib(f"Range mein <b>{len(filtered_dates)} dates</b> hain "
               f"({filtered_dates[0].strftime('%d %b %Y')} → {filtered_dates[-1].strftime('%d %b %Y')})")
    else:
        filtered_dates=[]
        wb_msg("Pehle attendance sheet upload karo.")

    # Session count
    sc1,sc2,sc3=st.columns(3)
    with sc1:
        session_mode=st.radio("Sessions count kaise?",
            ["📊 Sare dates use karo (auto)", "✏️ Manually enter karo"],
            horizontal=True)
    with sc2:
        if "auto" in session_mode.lower():
            total_sessions=len(filtered_dates)
            st.metric("Total Sessions", total_sessions)
        else:
            total_sessions=st.number_input("Total Sessions (rows in sheet)",
                min_value=1, max_value=500,
                value=len(filtered_dates) if filtered_dates else 50)
    with sc3:
        faculty_id=st.text_input("🧑‍🏫 Faculty Registration ID *",
                                  placeholder="e.g. IILMGG006412025")

    sc4,sc5=st.columns(2)
    with sc4:
        mandatory=st.selectbox("✅ Attendance Mandatory?",["TRUE","FALSE"])
    with sc5:
        tlo_max=st.number_input("📊 Max TLO Number (1–100)",
                                 min_value=1, max_value=100, value=5)
    all_tlos=[f"TLO{i}" for i in range(1,tlo_max+1)]

    st.divider()

    # ═══════════════════════════════════════════════════════════════
    # STEP 3 — DAY-WISE TIMING (7 days, AM/PM)
    # ═══════════════════════════════════════════════════════════════
    sh("⏰ Step 3 — Lecture Days & Timing (7 din ke options)")
    ib("Jinke din class hoti hai unhe enable karo — baaki off rahenge. AM/PM format supported.")

    # Detect which days appear in filtered dates
    detected={}
    for d in filtered_dates:
        if d: detected[d.strftime("%A")]=detected.get(d.strftime("%A"),0)+1

    day_timing={}   # {day_name: (start_24h, end_24h)}

    cols7=st.columns(7)
    for i,day in enumerate(ALL_DAYS):
        with cols7[i]:
            count=detected.get(day,0)
            enabled=st.checkbox(
                f"**{DAY_SHORT[day]}**",
                value=(count>0),
                key=f"day_en_{day}",
                help=f"{count} classes in selected range" if count else "Not in attendance"
            )
            if count>0:
                st.caption(f"🔵 {count} classes")
            else:
                st.caption("—")

            if enabled:
                # Start time
                sh1,sm1,sa1=st.columns([2,2,2])
                with sh1: st_h=st.selectbox("H",list(range(1,13)),index=9,key=f"sh_{day}",label_visibility="collapsed")
                with sm1: st_m=st.selectbox("M",["00","10","15","20","30","45"],index=0,key=f"sm_{day}",label_visibility="collapsed")
                with sa1: st_a=st.selectbox("",["AM","PM"],index=1,key=f"sa_{day}",label_visibility="collapsed")
                start24=to_24h(str(st_h),st_m,st_a)
                st.caption(f"Start: {start24}")

                # End time
                eh1,em1,ea1=st.columns([2,2,2])
                with eh1: et_h=st.selectbox("H",list(range(1,13)),index=10,key=f"eh_{day}",label_visibility="collapsed")
                with em1: et_m=st.selectbox("M",["00","10","15","20","30","45"],index=0,key=f"em_{day}",label_visibility="collapsed")
                with ea1: et_a=st.selectbox("",["AM","PM"],index=1,key=f"ea_{day}",label_visibility="collapsed")
                end24=to_24h(str(et_h),et_m,et_a)
                st.caption(f"End: {end24}")

                day_timing[day]=(start24,end24)

    # Timing summary
    if day_timing:
        st.markdown("**📋 Timing Summary:**")
        cols_sum=st.columns(len(day_timing))
        for i,(day,(s,e)) in enumerate(day_timing.items()):
            with cols_sum[i]:
                st.info(f"**{DAY_SHORT[day]}**\n{s} – {e}")

    st.divider()

    # ═══════════════════════════════════════════════════════════════
    # STEP 4 — UNIT COVERAGE
    # ═══════════════════════════════════════════════════════════════
    sh("📚 Step 4 — Unit Coverage & Module Config")

    syl_units=list(syl_data.keys()) if syl_data else []

    if syl_units:
        cov_mode=st.radio(
            "Kitna syllabus cover karna hai?",
            ["📖 Pura syllabus (sab units)", "✂️ Partial (kuch units hi)"],
            horizontal=True
        )
        if "partial" in cov_mode.lower() or "Partial" in cov_mode:
            selected_units=st.multiselect(
                "Kaun se units include karne hain?",
                options=syl_units,
                default=syl_units,
                help="Sirf inhi units se titles extract honge"
            )
        else:
            selected_units=syl_units
    else:
        selected_units=[]
        wb_msg("Syllabus upload karo — ya modules manually configure karo.")

    # Module configuration
    n_mods_default=min(len(selected_units),10) if selected_units else 4
    num_modules=st.selectbox("Kitne Modules?",list(range(1,11)),
                              index=n_mods_default-1,
                              format_func=lambda x:f"{x} Module{'s' if x>1 else ''}")

    modules_cfg=[]
    m_cols=st.columns(min(num_modules,4))
    for i in range(num_modules):
        with m_cols[i%min(num_modules,4)]:
            with st.expander(f"📦 Module {i+1}",expanded=(num_modules<=4)):
                def_name=selected_units[i] if i<len(selected_units) else f"Module {i+1}"
                mname=st.text_input("Name*",value=def_name,key=f"mn_{i}")

                def_tlo=[all_tlos[i%len(all_tlos)]] if all_tlos else ["TLO1"]
                mtlos=st.multiselect("TLOs",all_tlos,default=def_tlo,key=f"mt_{i}")
                tlo_str=" | ".join(mtlos) if mtlos else all_tlos[i%len(all_tlos)]

                unit_key=selected_units[i] if i<len(selected_units) else ""
                avail=len(syl_data.get(unit_key,[]))
                if avail: st.caption(f"📖 {avail} topics in syllabus")

                modules_cfg.append({"name":mname,"tlo":tlo_str,"unit_key":unit_key})

    if modules_cfg and syl_data:
        total_avail=sum(len(syl_data.get(m["unit_key"],[])) for m in modules_cfg)
        if total_sessions<=total_avail:
            sb(f"<b>{total_sessions}</b> sessions will be created from <b>{total_avail}</b> available topics — balanced distribution.")
        else:
            wb_msg(f"Sessions ({total_sessions}) > Available topics ({total_avail}). "
                   f"Extra sessions will repeat last topics.")

    st.divider()

    # ═══════════════════════════════════════════════════════════════
    # STEP 5 — BUILD PREVIEW
    # ═══════════════════════════════════════════════════════════════
    sh("👀 Step 5 — Preview Table Banao & Edit Karo")
    ib("'Build Preview' dabao → auto-balanced titles + auto-generated descriptions. Table directly edit kar sakte ho.")

    prev_btn=st.button("🔄 Build Preview Table",type="secondary",use_container_width=True)

    if prev_btn:
        if not filtered_dates:
            eb("Attendance sheet + date range set karo pehle.")
        else:
            # Balance titles
            if syl_data and modules_cfg:
                flat=balance_titles(syl_data,modules_cfg,total_sessions)
            else:
                flat=[{"module":modules_cfg[i%len(modules_cfg)]["name"] if modules_cfg else "Module",
                       "tlo":modules_cfg[i%len(modules_cfg)]["tlo"] if modules_cfg else "TLO1",
                       "title":f"Session {i+1}"} for i in range(total_sessions)]

            # Build rows
            use_dates=filtered_dates[:total_sessions]
            rows=[]
            for idx in range(total_sessions):
                # Date + timing
                if idx<len(use_dates):
                    d=use_dates[idx]
                    ds=d.strftime("%Y-%m-%d")
                    day_name=d.strftime("%A")
                    s_t,e_t=day_timing.get(day_name,("10:00","11:00"))
                    start_dt=f"{ds} {s_t}:00"
                    end_dt  =f"{ds} {e_t}:00"
                else:
                    ds=start_dt=end_dt="TBD"

                sess=flat[idx] if idx<len(flat) else {"module":"Extra","tlo":"TLO1","title":f"Extra Session {idx+1}"}
                desc=auto_desc(sess["title"],sess["module"])

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

            st.session_state.preview_df=pd.DataFrame(rows)
            st.session_state.results=None

    if st.session_state.preview_df is not None:
        df=st.session_state.preview_df
        sb(f"Preview ready: <b>{len(df)} rows</b> — edit karo table mein directly.")

        edited_df=st.data_editor(
            df, use_container_width=True, num_rows="fixed",
            hide_index=True, key="prev_ed",
            column_config={
                "Sr":                st.column_config.NumberColumn("Sr",width=50,disabled=True),
                "Module Name*":      st.column_config.TextColumn("Module Name* ✏️",    width=230),
                "Start Date Time*":  st.column_config.TextColumn("Start DateTime* ✏️", width=165),
                "End Date Time*":    st.column_config.TextColumn("End DateTime* ✏️",   width=165),
                "Title*":            st.column_config.TextColumn("Title* ✏️",          width=260),
                "Description*":      st.column_config.TextColumn("Description* ✏️",    width=370),
                "Mandatory*":        st.column_config.SelectboxColumn("Mandatory*",
                                         options=["TRUE","FALSE"],width=100),
                "TLO":               st.column_config.TextColumn("TLO ✏️",             width=130),
                "Faculty Reg ID*":   st.column_config.TextColumn("Faculty Reg ID* ✏️", width=165),
            },
            height=460,
        )
        st.session_state.preview_df=edited_df

    st.divider()

    # ═══════════════════════════════════════════════════════════════
    # STEP 6 — GENERATE
    # ═══════════════════════════════════════════════════════════════
    sh("🚀 Step 6 — Generate Karo")

    can_gen=bool(
        st.session_state.preview_df is not None and
        att_data and faculty_id.strip()
    )
    if not can_gen:
        miss=[]
        if not att_data: miss.append("Attendance Sheet")
        if not faculty_id.strip(): miss.append("Faculty Registration ID")
        if st.session_state.preview_df is None: miss.append("Step 5 mein Preview Table banao")
        if miss: ib(f"Pehle karo: <b>{', '.join(miss)}</b>")

    gen_btn=st.button("⚡ Generate Session Sheet + All Attendance Files",
                       type="primary",disabled=not can_gen,use_container_width=True)

    if gen_btn and can_gen:
        errors=[]; prog=st.progress(0); log=st.empty()
        try:
            df=st.session_state.preview_df

            # Build final session rows
            log.markdown('<div class="ib">📋 Session rows preparing…</div>',unsafe_allow_html=True)
            session_rows=[{
                "module":    str(r.get("Module Name*","")),
                "start_dt":  str(r.get("Start Date Time*","")),
                "end_dt":    str(r.get("End Date Time*","")),
                "title":     str(r.get("Title*","")),
                "description":str(r.get("Description*","")),
                "mandatory": str(r.get("Mandatory*","TRUE")),
                "tlo":       str(r.get("TLO","TLO1")),
                "faculty_id":str(r.get("Faculty Reg ID*",faculty_id.strip())),
            } for _,r in df.iterrows()]
            prog.progress(20)

            log.markdown('<div class="ib">📊 Session sheet generating…</div>',unsafe_allow_html=True)
            session_xlsx=gen_session_sheet(session_rows)
            sb(f"Session sheet: <b>{len(session_rows)} rows</b> written in DL.xlsx exact format.")
            prog.progress(50)

            daywise={}
            if att_tpl:
                log.markdown('<div class="ib">🗂️ Day-wise attendance generating…</div>',unsafe_allow_html=True)
                att_tpl_bytes=att_tpl.read()
                use_dates=[d for d in filtered_dates if d][:total_sessions]
                bar=st.progress(0)
                for i,d in enumerate(use_dates):
                    try:
                        fname=f"attendance_{d.strftime('%Y-%m-%d')}.xlsx"
                        daywise[fname]=gen_daywise_att(att_tpl_bytes,att_data,d)
                    except Exception as ex: errors.append(f"{d}: {ex}")
                    bar.progress((i+1)/max(len(use_dates),1))
                sb(f"Day-wise attendance: <b>{len(daywise)} files</b> ready.")
            else:
                wb_msg("Attendance template nahi diya — sirf session sheet download hogi.")

            prog.progress(85)
            zip_bytes=build_zip(session_xlsx,daywise)
            prog.progress(100)
            log.markdown('<div class="sb">🎉 Sab ready! Download karo neeche se.</div>',unsafe_allow_html=True)

            st.session_state.results={
                "session":session_xlsx,"daywise":daywise,
                "zip":zip_bytes,"errors":errors,"n":len(session_rows)
            }
        except Exception as ex:
            eb(f"Error: {ex}")
            import traceback; st.code(traceback.format_exc())

    # ═══════════════════════════════════════════════════════════════
    # STEP 7 — DOWNLOADS
    # ═══════════════════════════════════════════════════════════════
    if st.session_state.results:
        res=st.session_state.results
        st.divider()
        sh("📥 Step 7 — Download Karo")
        for e in res.get("errors",[]): wb_msg(f"Skipped: {e}")

        d1,d2=st.columns(2)
        with d1:
            st.download_button("📦 ⬇ Download ALL as ZIP",
                data=res["zip"],
                file_name=f"DG_Output_{datetime.now():%Y%m%d_%H%M}.zip",
                mime="application/zip",use_container_width=True)
        with d2:
            st.download_button("📋 ⬇ Download session_sheet.xlsx",
                data=res["session"],file_name="session_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        if res["daywise"]:
            st.markdown(f"#### 📅 Day-wise Attendance ({len(res['daywise'])} files)")
            files=list(res["daywise"].items())
            for row_i in range(0,len(files),4):
                chunk=files[row_i:row_i+4]; cols=st.columns(4)
                for ci,(fname,fb) in enumerate(chunk):
                    with cols[ci]:
                        dp=fname.replace("attendance_","").replace(".xlsx","")
                        try: label=datetime.strptime(dp,"%Y-%m-%d").strftime("%d %b %Y")
                        except: label=dp
                        st.download_button(f"📅 {label}",data=fb,file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,key=f"dl_{fname}")

        st.divider()
        if st.button("🔄 Reset"):
            for k in ["att_data","syl_data","preview_df","results","_att_nm","_syl_nm"]:
                st.session_state[k]=None
            st.rerun()

    st.markdown("<center style='color:#aaa;font-size:.7rem;margin-top:.8rem'>"
                "🔒 In-memory processing — no data stored on server</center>",
                unsafe_allow_html=True)


if __name__=="__main__":
    main()
