"""
╔══════════════════════════════════════════════════════════════════════╗
║        DG Session & Attendance Sheet Generator  v2.0                ║
║  Streamlit App  |  pandas · openpyxl · zipfile · python-docx        ║
╚══════════════════════════════════════════════════════════════════════╝

HOW TO RUN:
  Local  : streamlit run app.py
  Deploy : Push to GitHub → https://streamlit.io/cloud  (free shareable URL)
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
.main{padding:1rem 2rem}
.app-title{
    background:linear-gradient(135deg,#1a2f5a 0%,#2471a3 100%);
    color:white;padding:1.8rem 2rem;border-radius:14px;
    margin-bottom:1.5rem;text-align:center;
    box-shadow:0 4px 15px rgba(36,113,163,.35)
}
.app-title h1{margin:0;font-size:2rem;font-weight:800;letter-spacing:-0.5px}
.app-title p{margin:.4rem 0 0;opacity:.88;font-size:1rem}
.step-header{
    font-size:1.05rem;font-weight:700;color:#1a2f5a;
    padding:.5rem 0 .3rem;border-bottom:2px solid #2471a3;
    margin:1.2rem 0 .8rem
}
.info-box{background:#eaf4fb;border-left:4px solid #2471a3;
    padding:.7rem 1rem;border-radius:6px;margin:.4rem 0;font-size:.88rem}
.success-box{background:#d4edda;border-left:4px solid #28a745;
    padding:.7rem 1rem;border-radius:6px;margin:.4rem 0;font-size:.88rem}
.warn-box{background:#fff8e1;border-left:4px solid #f39c12;
    padding:.7rem 1rem;border-radius:6px;margin:.4rem 0;font-size:.88rem}
.error-box{background:#fde8e8;border-left:4px solid #e74c3c;
    padding:.7rem 1rem;border-radius:6px;margin:.4rem 0;font-size:.88rem}
.badge-ok{background:#d4edda;color:#155724;padding:.2rem .6rem;
    border-radius:10px;font-size:.78rem;font-weight:700}
.badge-warn{background:#fff3cd;color:#856404;padding:.2rem .6rem;
    border-radius:10px;font-size:.78rem;font-weight:700}
.preview-table{width:100%;border-collapse:collapse;font-size:.82rem}
.preview-table th{background:#1a2f5a;color:white;padding:.4rem .8rem;text-align:left}
.preview-table td{padding:.4rem .8rem;border-bottom:1px solid #dee2e6}
.preview-table tr:nth-child(even) td{background:#f4f8fd}
.stDownloadButton>button{
    background:#1a2f5a !important;color:white !important;
    border-radius:8px !important;font-weight:600 !important;width:100% !important
}
.stDownloadButton>button:hover{background:#2471a3 !important}
.deploy-box{
    background:#f0f9ff;border:1px dashed #2471a3;
    border-radius:10px;padding:1rem 1.2rem;margin-top:.8rem;font-size:.88rem
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# ─── UTILITY FUNCTIONS ───────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════

def parse_date(raw) -> date | None:
    """Convert any date-like value to a Python date object."""
    if raw is None:
        return None
    try:
        if isinstance(raw, float) and pd.isna(raw):
            return None
    except Exception:
        pass
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw

    s = str(raw).strip()
    s = re.sub(r"\([^)]*\)", "", s)          # remove (Monday) etc.
    s = re.sub(r"\b(sub|Sub|SUB)\b", "", s)  # remove Sub
    s = re.sub(r"\s+", " ", s).strip()

    for fmt in ("%d %b %Y", "%d %B %Y", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y",
                "%d %b %y", "%d-%b-%Y", "%d-%b-%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    try:
        return pd.to_datetime(s, dayfirst=True).date()
    except Exception:
        return None


def _copy_style(src, dst):
    """Copy openpyxl cell style."""
    if src.has_style:
        dst.font       = copy(src.font)
        dst.border     = copy(src.border)
        dst.fill       = copy(src.fill)
        dst.alignment  = copy(src.alignment)
        dst.protection = copy(src.protection)
        dst.number_format = src.number_format


def best_sheet_df(file_bytes: bytes, hints: list[str] = None) -> pd.DataFrame:
    """Return most relevant sheet as a raw (no-header) DataFrame."""
    xf    = pd.ExcelFile(io.BytesIO(file_bytes))
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
        n  = int(df.notna().sum().sum())
        if n > most:
            most, best = n, s
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=best, header=None)


def best_sheet_name(file_bytes: bytes, hints: list[str] = None) -> str:
    """Return name of the most relevant sheet."""
    xf    = pd.ExcelFile(io.BytesIO(file_bytes))
    names = xf.sheet_names
    if len(names) == 1:
        return names[0]
    if hints:
        for kw in hints:
            for s in names:
                if kw.lower() in s.lower():
                    return s
    best, most = names[0], -1
    for s in names:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, header=None)
        n  = int(df.notna().sum().sum())
        if n > most:
            most, best = n, s
    return best


# ══════════════════════════════════════════════════════════════════════
# ─── TIMING RESOLVER ─────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════

class TimingResolver:
    """
    Resolves (start_time, end_time) for any session date.

    Priority order per date:
      1. Exact date match in timing sheet   → "Exact match"
      2. Weekday pattern from timing sheet  → "Weekday pattern (Mon/Tue/…)"
      3. Sequential slot list (txt mode)    → "Slot #N"
      4. User default                       → "Default"
    """

    def __init__(self):
        self.date_map: dict[date, tuple[str, str]]    = {}   # exact date → times
        self.weekday_map: dict[int, tuple[str, str]]  = {}   # weekday(0=Mon) → times
        self._seq_slots: list[tuple[str, str]]        = []   # from txt
        self.default: tuple[str, str]                 = ("10:00", "11:00")

    # ── loaders ──────────────────────────────────────────────────────

    def load_excel(self, file_bytes: bytes) -> int:
        """
        Load from Excel.  Handles two formats:
          A) Columns 'Start Date Time' + 'End Date Time'  (DL.xlsx style)
          B) Generic time/slot column
        """
        sname = best_sheet_name(file_bytes, hints=["timing","time","schedule","session"])
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sname, header=None)

        # ── Locate header row ─────────────────────────────────────────
        header_row = 0
        for i, row in df_raw.iterrows():
            row_str = " ".join(str(v).lower() for v in row if pd.notna(v))
            if any(kw in row_str for kw in ["start","end","time","date","timing","slot"]):
                header_row = i
                break

        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sname, header=header_row)
        df.columns = [str(c).strip() for c in df.columns]

        # ── Find start / end columns ──────────────────────────────────
        start_col = end_col = None
        for c in df.columns:
            cl = c.lower()
            if ("start" in cl) or ("date" in cl and "end" not in cl and start_col is None):
                start_col = c
            if "end" in cl and end_col is None:
                end_col = c

        if start_col:
            loaded = 0
            for _, row in df.iterrows():
                rv = row.get(start_col)
                if pd.isna(rv):
                    continue
                try:
                    dt_s  = pd.to_datetime(rv)
                    s_t   = dt_s.strftime("%H:%M")
                    s_d   = dt_s.date()
                    rv_e  = row.get(end_col) if end_col else None
                    e_t   = (pd.to_datetime(rv_e).strftime("%H:%M")
                             if rv_e is not None and not pd.isna(rv_e)
                             else (dt_s + timedelta(hours=1)).strftime("%H:%M"))
                    self.date_map[s_d] = (s_t, e_t)
                    loaded += 1
                except Exception:
                    pass
            if loaded:
                self._derive_weekday()
                return loaded

        # ── Fallback: scan for time-like strings ──────────────────────
        return self._load_generic_df(df_raw)

    def load_txt(self, content: str) -> int:
        """Load timing slots from plain text (one slot per line)."""
        slots = []
        for line in content.splitlines():
            line = line.strip()
            if not line:
                continue
            times = re.findall(r"\d{1,2}:\d{2}(?:\s*[APap][Mm])?", line)
            if len(times) >= 2:
                slots.append((self._to24(times[0]), self._to24(times[1])))
            elif times:
                slots.append((self._to24(times[0]), "TBD"))
        self._seq_slots = slots
        return len(slots)

    def set_default(self, text: str):
        """Parse and store the user-supplied default timing string."""
        times = re.findall(r"\d{1,2}:\d{2}(?:\s*[APap][Mm])?", str(text))
        if len(times) >= 2:
            self.default = (self._to24(times[0]), self._to24(times[1]))
        elif times:
            self.default = (self._to24(times[0]), "TBD")

    # ── resolver ─────────────────────────────────────────────────────

    def resolve(self, session_date: date, idx: int = 0) -> tuple[str, str, str]:
        """Return (start, end, source_label)."""
        if session_date in self.date_map:
            return (*self.date_map[session_date], "Exact match")
        wd = session_date.weekday()
        if wd in self.weekday_map:
            return (*self.weekday_map[wd], f"Weekday pattern ({session_date.strftime('%A')})")
        if self._seq_slots:
            t = self._seq_slots[idx % len(self._seq_slots)]
            return (*t, f"Slot #{(idx % len(self._seq_slots)) + 1}")
        return (*self.default, "Default")

    def preview(self, dates: list[date]) -> pd.DataFrame:
        rows = []
        for i, d in enumerate(dates):
            s, e, src = self.resolve(d, i)
            rows.append({"Date": d.strftime("%d %b %Y"),
                         "Day": d.strftime("%A"),
                         "Start": s, "End": e, "Source": src})
        return pd.DataFrame(rows)

    # ── internals ────────────────────────────────────────────────────

    def _derive_weekday(self):
        wd_times: dict[int, list] = {}
        for d, times in self.date_map.items():
            wd_times.setdefault(d.weekday(), []).append(times)
        for wd, lst in wd_times.items():
            self.weekday_map[wd] = Counter(lst).most_common(1)[0][0]

    def _load_generic_df(self, df_raw: pd.DataFrame) -> int:
        for col_idx in range(df_raw.shape[1]):
            col = df_raw.iloc[:, col_idx].dropna().astype(str).tolist()
            hits = [v for v in col if re.search(r"\d{1,2}:\d{2}", v)]
            if hits:
                for h in hits:
                    times = re.findall(r"\d{1,2}:\d{2}", h)
                    if len(times) >= 2:
                        self._seq_slots.append((times[0], times[1]))
                    elif times:
                        self._seq_slots.append((times[0], "TBD"))
                return len(self._seq_slots)
        return 0

    @staticmethod
    def _to24(t: str) -> str:
        t = t.strip()
        for fmt in ("%I:%M %p", "%I:%M%p", "%H:%M"):
            try:
                return datetime.strptime(t.upper(), fmt).strftime("%H:%M")
            except ValueError:
                pass
        return t


# ══════════════════════════════════════════════════════════════════════
# ─── PARSER: MASTER ATTENDANCE ───────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════

def parse_attendance(file_bytes: bytes) -> dict:
    """
    Returns:
      dates    : [date, ...]            ordered unique session dates
      students : [{name,enrollment,roll,attendance:{date:0/1/None}}]
      meta     : {key:value}            header metadata (teacher, class, etc.)
    """
    df_raw = best_sheet_df(file_bytes, hints=["attendance","student"])

    # ── Find date row ─────────────────────────────────────────────────
    date_row_idx = date_col_start = None
    for i, row in df_raw.iterrows():
        hits, first = [], None
        for j, v in enumerate(row):
            d = parse_date(v)
            if d:
                hits.append(d)
                first = first if first is not None else j
        if len(hits) >= 2:
            date_row_idx  = i
            date_col_start = first
            break

    if date_row_idx is None:
        raise ValueError("❌ Cannot find session dates in the attendance sheet.")

    # ── col → date mapping ────────────────────────────────────────────
    date_row = df_raw.iloc[date_row_idx]
    col_to_date: dict[int, date] = {}
    dates_ordered: list[date]    = []
    seen: set[date]              = set()
    for j in range(date_col_start, df_raw.shape[1]):
        d = parse_date(date_row.iloc[j])
        if d:
            col_to_date[j] = d
            if d not in seen:
                dates_ordered.append(d)
                seen.add(d)

    # ── Find header row ───────────────────────────────────────────────
    name_col = enroll_col = roll_col = None
    header_row_idx = None
    for i in range(date_row_idx):
        row = df_raw.iloc[i]
        row_str = " ".join(str(v).lower() for v in row if pd.notna(v))
        if any(kw in row_str for kw in ["name","enrol","roll","sr."]):
            header_row_idx = i
            for j, v in enumerate(row):
                s = str(v).lower().strip()
                if "name" in s and name_col is None:       name_col   = j
                elif "enrol" in s and enroll_col is None:  enroll_col = j
                elif "roll" in s and roll_col is None:     roll_col   = j
            break

    # ── Meta ─────────────────────────────────────────────────────────
    meta: dict[str, str] = {}
    for i in range(header_row_idx or 0):
        row = df_raw.iloc[i]
        for j in range(len(row) - 1):
            k, v = str(row.iloc[j]).strip(), str(row.iloc[j+1]).strip()
            if k not in ("nan","") and v not in ("nan",""):
                meta[k] = v

    # ── Students ──────────────────────────────────────────────────────
    students: list[dict] = []
    for i in range(date_row_idx + 1, df_raw.shape[0]):
        row = df_raw.iloc[i]
        if row.notna().sum() < 3:
            continue
        name = str(row.iloc[name_col]).strip() if name_col is not None and pd.notna(row.iloc[name_col]) else ""
        if not name or name.lower() in ("nan","none",""):
            continue
        enroll = str(row.iloc[enroll_col]).strip() if enroll_col is not None and pd.notna(row.iloc[enroll_col]) else ""
        roll   = str(row.iloc[roll_col]).strip()   if roll_col   is not None and pd.notna(row.iloc[roll_col])   else ""
        att: dict[date, int | None] = {}
        for col_j, d in col_to_date.items():
            val = row.iloc[col_j]
            att[d] = int(float(val)) if not pd.isna(val) else None
        students.append({"name": name, "enrollment": enroll, "roll": roll, "attendance": att})

    return {"dates": dates_ordered, "students": students, "meta": meta}


# ══════════════════════════════════════════════════════════════════════
# ─── PARSER: SYLLABUS ────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════

def parse_syllabus(file_bytes: bytes, filename: str) -> list[str]:
    ext = filename.lower().rsplit(".",1)[-1]
    topics: list[str] = []

    if ext in ("xlsx","xls"):
        df_raw = best_sheet_df(file_bytes, hints=["syllabus","topic","session"])
        for i, row in df_raw.iterrows():
            if any(kw in str(v).lower() for v in row if pd.notna(v) for kw in ["topic","title","lecture","unit","session"]):
                df = pd.read_excel(io.BytesIO(file_bytes), header=i)
                df.columns = [str(c).strip() for c in df.columns]
                for c in df.columns:
                    if any(kw in c.lower() for kw in ["topic","title","session","lecture"]):
                        topics = df[c].dropna().astype(str).str.strip().tolist()
                        break
                break
        if not topics:
            for ci in range(df_raw.shape[1]):
                col = df_raw.iloc[:,ci].dropna().astype(str).str.strip().tolist()
                good = [t for t in col if len(t)>3 and "nan" not in t.lower()]
                if len(good) > 3:
                    topics = good; break

    elif ext == "docx":
        import docx as _docx
        doc = _docx.Document(io.BytesIO(file_bytes))
        in_u = False
        for p in doc.paragraphs:
            t = p.text.strip()
            if not t: continue
            if re.match(r"^UNIT\s+\d+", t, re.IGNORECASE):
                in_u = True; topics.append(t); continue
            if in_u:
                if re.match(r"^Case.?law", t, re.IGNORECASE): continue
                if re.search(r"\bv\.\b|\bAIR\b|\bILR\b|\bSCC\b|https?://|LL \(", t): continue
                if len(t) > 4: topics.append(t)
        if not topics:
            topics = [p.text.strip() for p in doc.paragraphs if p.text.strip() and len(p.text.strip()) > 4]

    elif ext == "txt":
        topics = [l.strip() for l in file_bytes.decode("utf-8", errors="ignore").splitlines()
                  if l.strip() and len(l.strip()) > 2]

    elif ext == "csv":
        df = pd.read_csv(io.BytesIO(file_bytes))
        for c in df.columns:
            if any(kw in c.lower() for kw in ["topic","title","session","lecture"]):
                topics = df[c].dropna().astype(str).str.strip().tolist(); break
        if not topics:
            topics = df.iloc[:,0].dropna().astype(str).str.strip().tolist()

    return [t for t in topics if t and t.lower() not in ("nan","none","")]


# ══════════════════════════════════════════════════════════════════════
# ─── GENERATOR: SESSION SHEET ────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════

def generate_session_sheet(
    template_bytes: bytes,
    dates: list[date],
    topics: list[str],
    timing: TimingResolver,
    meta: dict,
    subject_name: str = "",
    faculty_reg_id: str = "",
) -> bytes:
    """
    Fill DG Session Sheet template.
    Each attendance date → one row with correct timing from TimingResolver.
    """
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # ── Detect header row & column positions ─────────────────────────
    COL_HINTS = {
        "module":      ["module"],
        "start":       ["start"],
        "end":         ["end"],
        "title":       ["title"],
        "description": ["description","desc"],
        "mandatory":   ["mandatory"],
        "tlo":         ["tlo"],
        "faculty":     ["faculty","teaching","registration"],
    }
    col_map: dict[str, int] = {}
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
    style: dict[int, object] = {}
    if ws.max_row >= data_start:
        for c in range(1, ws.max_column + 1):
            style[c] = ws.cell(data_start, c)

    # Clear old data rows
    for r in range(ws.max_row, header_row, -1):
        ws.delete_rows(r)

    # ── Write rows ────────────────────────────────────────────────────
    module = subject_name or meta.get("Subject Title:", "") or "Course"
    fac_id = faculty_reg_id or "IILMGG006412025"

    def dt_str(d: date, t: str) -> str:
        return f"{d} {t}:00" if t != "TBD" else str(d)

    for idx, session_date in enumerate(dates):
        row_num   = data_start + idx
        topic     = topics[idx] if idx < len(topics) else "Extra Session"
        s_t, e_t, _ = timing.resolve(session_date, idx)

        row_vals = {
            col_map.get("module"):      module,
            col_map.get("start"):       dt_str(session_date, s_t),
            col_map.get("end"):         dt_str(session_date, e_t),
            col_map.get("title"):       topic,
            col_map.get("description"): topic,
            col_map.get("mandatory"):   "TRUE",
            col_map.get("tlo"):         f"TLO{(idx % 5) + 1}",
            col_map.get("faculty"):     fac_id,
        }

        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row_num, c)
            if c in style:
                _copy_style(style[c], cell)
            if c in row_vals and row_vals[c] is not None:
                cell.value = row_vals[c]

    # Auto-fit columns
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        max_w  = max((len(str(cell.value or "")) for cell in col), default=10)
        ws.column_dimensions[letter].width = min(max_w + 4, 55)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════
# ─── GENERATOR: DAY-WISE ATTENDANCE ──────────────────────────────────
# ══════════════════════════════════════════════════════════════════════

def generate_daywise_attendance(
    template_bytes: bytes,
    att_data: dict,
    session_date: date,
) -> bytes:
    """One attendance file per date — PRESENT (green) / ABSENT (red)."""
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # Detect header row
    header_row = 1
    email_col = regid_col = attend_col = None
    for r in range(1, min(5, ws.max_row + 1)):
        vals = [str(ws.cell(r, c).value or "").lower() for c in range(1, ws.max_column + 1)]
        if any(kw in " ".join(vals) for kw in ["email","registration","attendance"]):
            header_row = r
            for ci, v in enumerate(vals, 1):
                if "email" in v:                          email_col  = ci
                elif "registration" in v or "reg" in v:   regid_col  = ci
                elif "attendance" in v:                   attend_col = ci
            break

    # Email lookup from template
    reg_to_email: dict[str, str] = {}
    try:
        sname = best_sheet_name(template_bytes)
        tdf   = pd.read_excel(io.BytesIO(template_bytes), sheet_name=sname, header=header_row-1)
        tdf.columns = [str(c).strip() for c in tdf.columns]
        ec = next((c for c in tdf.columns if "email" in c.lower()), None)
        rc = next((c for c in tdf.columns if "reg"   in c.lower()), None)
        if ec and rc:
            for _, row in tdf.iterrows():
                em = str(row[ec]).strip(); rg = str(row[rc]).strip()
                if em and em.lower() not in ("nan","none"):
                    reg_to_email[rg] = em
    except Exception:
        pass

    # Capture style
    style: dict[int, object] = {}
    if ws.max_row >= header_row + 1:
        for c in range(1, ws.max_column + 1):
            style[c] = ws.cell(header_row + 1, c)

    # Clear old rows
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
                _copy_style(style[c], cell)
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

def build_zip(session_bytes: bytes, daywise: dict[str, bytes]) -> bytes:
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
        <p>Upload 5 files → timings auto-matched by date/weekday → download ready Excel sheets</p>
    </div>
    """, unsafe_allow_html=True)

    if "results" not in st.session_state:
        st.session_state.results = None

    # ══════════════════════════════════════════════════════════════════
    # STEP 1 — UPLOADS
    # ══════════════════════════════════════════════════════════════════
    st.markdown('<div class="step-header">📂 Step 1 — Upload Your Files</div>', unsafe_allow_html=True)

    L, R = st.columns(2)

    with L:
        st.markdown("**📊 Data Files**")

        att_file = st.file_uploader("📅 Master Attendance Sheet (.xlsx)", type=["xlsx"], key="att")
        if att_file:
            st.markdown(f'<span class="badge-ok">✅ {att_file.name}</span>', unsafe_allow_html=True)

        syl_file = st.file_uploader("📚 Syllabus File (.xlsx / .docx / .txt)",
                                     type=["xlsx","docx","txt","csv"], key="syl")
        if syl_file:
            st.markdown(f'<span class="badge-ok">✅ {syl_file.name}</span>', unsafe_allow_html=True)

        tim_file = st.file_uploader("⏰ Timing Details Sheet — e.g. DL.xlsx (.xlsx / .txt)",
                                     type=["xlsx","txt","csv"], key="tim")
        if tim_file:
            st.markdown(f'<span class="badge-ok">✅ {tim_file.name}</span>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="warn-box">⚠️ Upload your timing sheet (DL.xlsx). '
                        'App will match timings to attendance dates automatically.</div>',
                        unsafe_allow_html=True)

    with R:
        st.markdown("**📄 DG Format Templates**")

        ses_tpl = st.file_uploader("📋 DG Session Format Template (.xlsx)", type=["xlsx"], key="stpl")
        if ses_tpl:
            st.markdown(f'<span class="badge-ok">✅ {ses_tpl.name}</span>', unsafe_allow_html=True)

        att_tpl = st.file_uploader("🗂️ DG Attendance Format Template (.xlsx)", type=["xlsx"], key="atpl")
        if att_tpl:
            st.markdown(f'<span class="badge-ok">✅ {att_tpl.name}</span>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════
    # STEP 2 — SETTINGS
    # ══════════════════════════════════════════════════════════════════
    st.markdown('<div class="step-header">⚙️ Step 2 — Settings (Optional)</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        subject_name = st.text_input("📖 Subject / Course Name",
                                     placeholder="e.g. Law of Contract II")
    with c2:
        faculty_reg  = st.text_input("🧑‍🏫 Faculty Registration ID",
                                     placeholder="e.g. IILMGG006412025")
    with c3:
        default_time = st.text_input("⏱️ Fallback Timing (if date not in timing sheet)",
                                     value="10:00 - 11:00")

    # ══════════════════════════════════════════════════════════════════
    # STEP 3 — GENERATE
    # ══════════════════════════════════════════════════════════════════
    st.markdown('<div class="step-header">🚀 Step 3 — Generate</div>', unsafe_allow_html=True)

    ready = bool(att_file and ses_tpl and att_tpl)
    if not ready:
        st.markdown('<div class="info-box">ℹ️ Upload at least: Attendance Sheet + both DG Templates.</div>',
                    unsafe_allow_html=True)

    go = st.button("⚡ Generate All Files", type="primary",
                   disabled=not ready, use_container_width=True)

    if go and ready:
        errors: list[str] = []
        progress = st.progress(0)
        log      = st.empty()

        try:
            # 1 ── Attendance ─────────────────────────────────────────
            log.markdown('<div class="info-box">📅 Parsing attendance sheet…</div>',
                         unsafe_allow_html=True)
            att_data = parse_attendance(att_file.read())
            dates    = att_data["dates"]
            meta     = att_data["meta"]
            progress.progress(15)
            st.markdown(
                f'<div class="success-box">✅ Attendance: '
                f'<b>{len(dates)} session dates</b> &nbsp;·&nbsp; '
                f'<b>{len(att_data["students"])} students</b></div>',
                unsafe_allow_html=True)

            # 2 ── Timing ─────────────────────────────────────────────
            log.markdown('<div class="info-box">⏰ Building date→timing map…</div>',
                         unsafe_allow_html=True)
            timing = TimingResolver()
            timing.set_default(default_time)

            if tim_file:
                ext = tim_file.name.lower().rsplit(".",1)[-1]
                tb  = tim_file.read()
                if ext == "txt":
                    n = timing.load_txt(tb.decode("utf-8", errors="ignore"))
                else:
                    n = timing.load_excel(tb)

                st.markdown(
                    f'<div class="success-box">✅ Timing sheet loaded: '
                    f'<b>{n} total entries</b> &nbsp;·&nbsp; '
                    f'<b>{len(timing.date_map)} exact dates</b> &nbsp;·&nbsp; '
                    f'<b>{len(timing.weekday_map)} weekday patterns derived</b> '
                    f'(Mon/Tue/Wed/Thu/Fri)</div>',
                    unsafe_allow_html=True)
            else:
                st.markdown('<div class="warn-box">⚠️ No timing file — using default for all.</div>',
                            unsafe_allow_html=True)

            # Show timing preview
            with st.expander("🔍 Preview: How timings are mapped to your attendance dates", expanded=True):
                prev_df = timing.preview(dates)
                # Colour rows by source
                def colour_source(row):
                    if row["Source"] == "Exact match":
                        return ["background-color:#d4edda"]*len(row)
                    elif "Weekday" in row["Source"]:
                        return ["background-color:#fff8e1"]*len(row)
                    else:
                        return ["background-color:#fde8e8"]*len(row)
                st.dataframe(
                    prev_df.style.apply(colour_source, axis=1),
                    use_container_width=True,
                    hide_index=True,
                )
                st.markdown(
                    "🟢 Green = exact date match &nbsp; 🟡 Yellow = weekday pattern &nbsp; 🔴 Red = default fallback",
                    unsafe_allow_html=True)
            progress.progress(35)

            # 3 ── Syllabus ───────────────────────────────────────────
            topics: list[str] = []
            if syl_file:
                log.markdown('<div class="info-box">📚 Parsing syllabus…</div>',
                             unsafe_allow_html=True)
                topics = parse_syllabus(syl_file.read(), syl_file.name)
                st.markdown(f'<div class="success-box">✅ Syllabus: <b>{len(topics)} topics</b> extracted.</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div class="warn-box">⚠️ No syllabus — all sessions → "Extra Session".</div>',
                            unsafe_allow_html=True)

            if topics and len(topics) < len(dates):
                st.markdown(
                    f'<div class="warn-box">⚠️ Topics ({len(topics)}) < Dates ({len(dates)}). '
                    f'Last {len(dates)-len(topics)} sessions will use "Extra Session".</div>',
                    unsafe_allow_html=True)
            progress.progress(50)

            # 4 ── Session Sheet ──────────────────────────────────────
            log.markdown('<div class="info-box">📋 Generating session sheet…</div>',
                         unsafe_allow_html=True)
            session_xlsx = generate_session_sheet(
                template_bytes=ses_tpl.read(),
                dates=dates,
                topics=topics,
                timing=timing,
                meta=meta,
                subject_name=subject_name or meta.get("Subject Title:", ""),
                faculty_reg_id=faculty_reg or "",
            )
            st.markdown(f'<div class="success-box">✅ Session sheet: <b>{len(dates)} rows</b> written with correct timings.</div>',
                        unsafe_allow_html=True)
            progress.progress(65)

            # 5 ── Day-wise Attendance ────────────────────────────────
            log.markdown('<div class="info-box">🗂️ Generating day-wise attendance files…</div>',
                         unsafe_allow_html=True)
            att_tpl_bytes = att_tpl.read()
            daywise: dict[str, bytes] = {}
            bar = st.progress(0)

            for i, d in enumerate(dates):
                fname = f"attendance_{d.strftime('%Y-%m-%d')}.xlsx"
                try:
                    daywise[fname] = generate_daywise_attendance(att_tpl_bytes, att_data, d)
                except Exception as ex:
                    errors.append(f"{d}: {ex}")
                bar.progress((i+1) / len(dates))

            st.markdown(f'<div class="success-box">✅ Day-wise: <b>{len(daywise)} files</b> created.</div>',
                        unsafe_allow_html=True)
            progress.progress(90)

            # 6 ── ZIP ────────────────────────────────────────────────
            log.markdown('<div class="info-box">📦 Packing ZIP archive…</div>',
                         unsafe_allow_html=True)
            zip_bytes = build_zip(session_xlsx, daywise)
            progress.progress(100)
            log.markdown('<div class="success-box">🎉 Done! Download your files below.</div>',
                         unsafe_allow_html=True)

            st.session_state.results = {
                "session": session_xlsx,
                "daywise": daywise,
                "zip":     zip_bytes,
                "dates":   dates,
                "errors":  errors,
            }

        except ValueError as ve:
            st.markdown(f'<div class="error-box">{ve}</div>', unsafe_allow_html=True)
        except Exception as ex:
            st.markdown(f'<div class="error-box">Unexpected error: {ex}</div>',
                        unsafe_allow_html=True)
            import traceback
            st.code(traceback.format_exc())

    # ══════════════════════════════════════════════════════════════════
    # STEP 4 — DOWNLOADS
    # ══════════════════════════════════════════════════════════════════
    if st.session_state.results:
        res = st.session_state.results
        st.markdown('<div class="step-header">📥 Step 4 — Download Files</div>',
                    unsafe_allow_html=True)

        for e in res.get("errors",[]):
            st.markdown(f'<div class="warn-box">⚠️ Skipped — {e}</div>', unsafe_allow_html=True)

        total = 1 + len(res["daywise"])
        st.markdown(f"""
        <table class="preview-table">
          <tr><th>Output File</th><th>Count</th><th>Status</th></tr>
          <tr><td>📋 session_sheet.xlsx</td><td>1</td>
              <td><span class="badge-ok">✅ Ready</span></td></tr>
          <tr><td>📅 attendance_YYYY-MM-DD.xlsx</td><td>{len(res["daywise"])}</td>
              <td><span class="badge-ok">✅ Ready</span></td></tr>
          <tr><td>📦 ZIP (all {total} files)</td><td>1</td>
              <td><span class="badge-ok">✅ Ready</span></td></tr>
        </table><br>
        """, unsafe_allow_html=True)

        # Primary downloads
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "📦 ⬇ Download ALL as ZIP",
                data=res["zip"],
                file_name=f"DG_Output_{datetime.now():%Y%m%d_%H%M}.zip",
                mime="application/zip",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                "📋 ⬇ Download session_sheet.xlsx",
                data=res["session"],
                file_name="session_sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        # Day-wise attendance
        st.markdown("#### 📅 Individual Day-wise Attendance Files")
        files  = list(res["daywise"].items())
        N_COLS = 4
        for row_i in range(0, len(files), N_COLS):
            chunk = files[row_i: row_i + N_COLS]
            cols  = st.columns(N_COLS)
            for ci, (fname, fb) in enumerate(chunk):
                with cols[ci]:
                    dp = fname.replace("attendance_","").replace(".xlsx","")
                    try:   label = datetime.strptime(dp, "%Y-%m-%d").strftime("%d %b %Y")
                    except: label = dp
                    st.download_button(
                        f"📅 {label}",
                        data=fb, file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"dl_{fname}",
                    )

        st.markdown("---")
        if st.button("🔄 Reset & Start Over"):
            st.session_state.results = None
            st.rerun()

    # ══════════════════════════════════════════════════════════════════
    # FACULTY DISTRIBUTION GUIDE
    # ══════════════════════════════════════════════════════════════════
    with st.expander("🌐 100 Faculty ko kaise share karein? (Deployment Guide)"):
        st.markdown("""
        <div class="deploy-box">
        <b>✅ Option 1 — Streamlit Cloud (FREE) — Best for sharing with 100 faculty</b><br>
        1. GitHub pe free account banayein → <code>app.py</code> + <code>requirements.txt</code> upload karein<br>
        2. <a href="https://streamlit.io/cloud" target="_blank">streamlit.io/cloud</a> → "New App" → apna repo select karein<br>
        3. Ek permanent URL milega jaise: <code>https://yourname-dg-generator.streamlit.app</code><br>
        4. Yeh URL WhatsApp/email se 100 faculty ko bhej dein — kisi ko kuch install nahi karna 🎉<br><br>

        <b>✅ Option 2 — Local Network (LAN) sharing</b><br>
        <code>streamlit run app.py --server.address 0.0.0.0 --server.port 8501</code><br>
        Faculty browser mein kholein: <code>http://&lt;AAPKA_IP&gt;:8501</code><br><br>

        <b>✅ Option 3 — Faculty khud run karein</b><br>
        <code>pip install -r requirements.txt</code> → <code>streamlit run app.py</code>
        </div>
        """, unsafe_allow_html=True)

    st.markdown(
        "<br><center style='color:#aaa;font-size:.78rem;'>"
        "🔒 All processing in-memory — no data stored &nbsp;|&nbsp;"
        "Streamlit · pandas · openpyxl · python-docx"
        "</center>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
