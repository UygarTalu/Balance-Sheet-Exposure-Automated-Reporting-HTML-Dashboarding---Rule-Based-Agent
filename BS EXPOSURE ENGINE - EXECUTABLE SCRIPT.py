# -*- coding: utf-8 -*-
"""
Merged Exposure Automation (WORKING-ready enhancements)
Base: EXPOSURE AUTOMATION CODE 2 (final not working)  + EXE-ready patterns from CODE 1 (working).

What changed (high level):
 - Added robust startup logging + log file path under %LOCALAPPDATA% (so remote user can send logs)
 - Top-level exception catcher writes exposure_last_error.txt and shows a Windows MessageBox
 - Keeps all new features (IS_IN_ANALYSIS flows) intact
 - Keeps interactive flows; doesn't remove functionality
 - Keeps xlwings abort behavior (clear message if xlwings missing)
 - Multiprocessing freeze_support preserved

Additional revision (2025-11-17):
 - Introduced read_excel_sheet_auto() for flexible master sheet detection
 - Exported MASTERDATA now uses the same sheet name (REVAL_MASTER_SHEET_IN_SOURCE)
   so that Python-generated files can be fed back as masterdata input next month

Additional revision (2025-11-18):
 - align_to_schema() now uses std_header-based matching so that columns like
   "Account", "ACCOUNT NO", etc. still map correctly to the BASE schema
   column "ACCOUNT" (and friends). This fixes the issue where COMPANY/DATE
   were filled but all other columns were blank for new monthly files.
"""

import os, sys, re, time, warnings, pandas as pd, traceback
from pathlib import Path

warnings.simplefilter("ignore", FutureWarning)
warnings.filterwarnings("ignore", message="Parsing dates.*dayfirst=True.*", category=UserWarning)

# --- Preload commonly problematic Windows COM lib (helps PyInstaller detect) ---
try:
    import win32com.client  # noqa: F401
except Exception:
    # don't fail here — xlwings import will indicate missing Excel runtime later
    pass

# ---------------------------
# Robust logging / startup wrapper
# ---------------------------
try:
    if getattr(sys, 'frozen', False):
        # When packaged by PyInstaller, sys._MEIPASS points to temp bundle (read-only)
        base_dir = Path(sys._MEIPASS)
        run_dir = Path(os.path.dirname(sys.executable))
    else:
        base_dir = Path(__file__).parent
        run_dir = Path.cwd()

    log_dir = Path(os.getenv('LOCALAPPDATA', run_dir)) / "ExposureAutomationLogs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "exposure_automation_run.log"

    import logging
    # keep a simple logger writing to file and stdout
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(levelname)s: %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.info("=== Exposure Automation STARTUP ===")
    logging.info(f"Base dir: {base_dir}; Run dir: {run_dir}; frozen={getattr(sys,'frozen',False)}")
except Exception:
    # fallback print if logging init fails
    print("Failed to initialize logging:", traceback.format_exc())
    log_file = Path.cwd() / "exposure_automation_run.log"

def write_error_and_notify(exc: Exception):
    try:
        txt = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        # write to last_error and to log (if logging available)
        try:
            with open(Path(run_dir) / "exposure_last_error.txt", "w", encoding="utf-8") as f:
                f.write(txt)
        except Exception:
            pass
        try:
            logging.error("Unhandled exception:\n%s", txt)
        except Exception:
            print("ERROR:", txt)
        # Windows message box (safe fallback)
        try:
            import ctypes
            MB_OK = 0x0
            ctypes.windll.user32.MessageBoxW(
                0,
                f"Exposure Automation encountered an error.\nCheck: {log_file}\n\nError: {str(exc)[:200]}",
                "Exposure Automation - Error",
                MB_OK
            )
        except Exception:
            pass
    except Exception:
        pass

# ---------------------------
# (Below is the original combined application logic)
# - Mostly preserved from the "EXPOSURE AUTOMATION CODE 2" (new)
# - Kept important functions & flows from CODE 1 where relevant
# ---------------------------

# === CONFIG ===
REVAL_MASTER_SHEET_IN_SOURCE = "MASTERDATA_REVAL"
DEST_SHEET_REVAL = "REVAL_MASTER_DATA_1"
DEST_TABLE_NAME = "MASTERDATA1"
COA_DEST_SHEET = "CHART_OF_ACCOUNTS"
COA_TABLE_BASENAME = "COA_DATA"
STATS_SHEET = "DATA_PROCESS_STATS"
GATE_SHEET = "BS_EXPOSURE_REPORT"
GATE_CELL = "H5"
VALID_INPUT_EXTS = {".xlsx", ".xlsm", ".xls"}
VALID_REPORT_EXTS = {".xlsx", ".xlsm"}

COMPANY_PATTERNS = [("YPNLBV","YPNL BV"), ("YPNL BV","YPNL BV"), ("YPOPL","YPOPL"), ("YPRODAS","YPRODAS")]

MONTH_NAME_TO_NUM = {
    "JAN":1,"JANUARY":1,"FEB":2,"FEBRUARY":2,"MAR":3,"MARCH":3,"APR":4,"APRIL":4,"MAY":5,
    "JUN":6,"JUNE":6,"JUL":7,"JULY":7,"AUG":8,"AUGUST":8,"SEP":9,"SEPT":9,"SEPTEMBER":9,
    "OCT":10,"OCTOBER":10,"NOV":11,"NOVEMBER":11,"DEC":12,"DECEMBER":12
}
MMM = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

# New file's DICT_IS_IN_ANALYSIS (kept)
DICT_IS_IN_ANALYSIS = {
    '116010000': 'CONSIDERED','204010000': 'CONSIDERED','205020000': 'CONSIDERED','208010600': 'CONSIDERED',
    '208039900': 'NOT_CONSIDERED','299020700': 'CONSIDERED','299030100': 'NOT_CONSIDERED','301020200': 'CONSIDERED',
    '301030400': 'CONSIDERED','301030900': 'NOT_CONSIDERED','301039900': 'CONSIDERED','303020000': 'CONSIDERED',
    '310010000': 'NOT_CONSIDERED','202010000': 'CONSIDERED','204020000': 'CONSIDERED','207000000': 'CONSIDERED',
    '208020500': 'NOT_CONSIDERED','299020200': 'CONSIDERED','299020800': 'CONSIDERED','301010100': 'CONSIDERED',
    '301020400': 'NOT_CONSIDERED','301030500': 'CONSIDERED','301031000': 'CONSIDERED','301050000': 'NOT_CONSIDERED',
    '304020000': 'CONSIDERED','312010000': 'CONSIDERED','203020000': 'CONSIDERED','205010000': 'CONSIDERED',
    '208010300': 'CONSIDERED','208029900': 'NOT_CONSIDERED','299020600': 'CONSIDERED','299020900': 'CONSIDERED',
    '301020100': 'CONSIDERED','301030100': 'CONSIDERED','301030600': 'CONSIDERED','301031500': 'CONSIDERED',
    '302020000': 'CONSIDERED','308060000': 'CONSIDERED','408060000': 'CONSIDERED',
}
DICT_ACCOUNT_DESC_FALLBACK = {}

STEP = 0
def slog(msg):
    global STEP
    STEP += 1
    print(f"[Step {STEP}] {msg}")
    try:
        logging.info(msg)
    except Exception:
        pass

def abort(msg):
    print(f"[ERROR] {msg}")
    try:
        logging.error(msg)
    except Exception:
        pass
    print("PROCEDURE DISTORTED — PLEASE RESTART WITH CORRECT INPUTS")
    sys.exit(1)

def safe_input(prompt: str) -> str:
    try:
        return input(prompt)
    except EOFError:
        abort("NO INPUT WAS PROVIDED (EOF).")
    except KeyboardInterrupt:
        abort("OPERATION CANCELLED BY USER (KeyboardInterrupt)")

# === Prompts / IO ===
def ensure_ext(path, allowed):
    ext = os.path.splitext(path)[1].lower()
    if ext not in allowed:
        abort(f"UNSUPPORTED FILE TYPE: {ext}. Allowed: {', '.join(sorted(allowed))}")

def prompt_path(q, must_exist=True, allowed_exts=None):
    raw = safe_input(q).strip().strip('"').strip("'")
    if not raw:
        abort(f"NO INPUT FOR: {q}")
    p = os.path.abspath(os.path.expanduser(raw))
    if must_exist and not os.path.exists(p):
        abort(f"FILE NOT FOUND: {p}")
    if allowed_exts:
        ensure_ext(p, allowed_exts)
    return p

def prompt_yesno(q):
    raw = safe_input(q).strip().upper()
    while raw not in {"YES","NO"}:
        raw = safe_input("Please type YES or NO: ").strip().upper()
    return raw

def prompt_int(q):
    raw = safe_input(q).strip()
    if not raw:
        abort(f"NO INPUT FOR: {q}")
    try:
        n = int(raw)
    except:
        abort(f"INVALID INTEGER FOR: {q} (got '{raw}')")
    if n < 0:
        abort(f"INTEGER MUST BE >= 0 FOR: {q} (got {n})")
    return n

def prompt_year(q):
    raw = safe_input(q).strip()
    if len(raw) != 4 or not raw.isdigit():
        abort("YEAR MUST BE 4 DIGITS (e.g., 2025)")
    yr = int(raw)
    if yr < 1900 or yr > 2100:
        abort(f"YEAR OUT OF RANGE: {yr}")
    return yr

def prompt_month(q):
    raw = safe_input(q).strip()
    if not raw:
        abort(f"NO INPUT FOR: {q}")
    u = raw.upper().replace(".","").strip()
    if u.isdigit():
        m = int(u)
        if 1 <= m <= 12:
            return m
        abort(f"MONTH NUMBER OUT OF RANGE: {m}")
    if u in MONTH_NAME_TO_NUM:
        return MONTH_NAME_TO_NUM[u]
    short = u[:3]
    if short in MONTH_NAME_TO_NUM:
        return MONTH_NAME_TO_NUM[short]
    abort(f"COULD NOT RECOGNIZE MONTH: '{raw}'")

def detect_company_from_filename(path):
    fname = os.path.basename(path).upper()
    for pat,label in COMPANY_PATTERNS:
        if pat.upper() in fname:
            return label
    manual = safe_input(f'No company tag in "{os.path.basename(path)}". Type COMPANY label: ').strip()
    if not manual:
        abort("COMPANY LABEL WAS NOT PROVIDED")
    return manual

def unique_copy_path_same_folder(reporting_path):
    folder = os.path.dirname(reporting_path)
    name,ext = os.path.splitext(os.path.basename(reporting_path))
    candidate = os.path.join(folder, f"{name}_PYTHON_GENERATED{ext}")
    i = 2
    while os.path.exists(candidate):
        candidate = os.path.join(folder, f"{name}_PYTHON_GENERATED ({i}){ext}")
        i += 1
    return candidate

def read_excel_first_sheet(path):
    try:
        return pd.read_excel(path, sheet_name=0, dtype=object)
    except Exception as e:
        abort(f"FAILED TO READ (first sheet): {e}")

def read_excel_named_sheet(path, sheet):
    try:
        return pd.read_excel(path, sheet_name=sheet, dtype=object)
    except Exception as e:
        abort(f"FAILED TO READ SHEET '{sheet}': {e}")

# NEW: flexible sheet reader for masterdata
def read_excel_sheet_auto(path, preferred=REVAL_MASTER_SHEET_IN_SOURCE):
    """
    Try to read the preferred sheet first, then fall back to several candidates,
    finally to the first sheet (index 0). If none works, abort with clear message.
    This allows feeding back Python-generated MASTERDATA files that might use a
    different sheet name.
    """
    # 1) Preferred
    try:
        return pd.read_excel(path, sheet_name=preferred, dtype=object)
    except Exception:
        pass

    # 2) Fallback candidates
    candidates = [
        REVAL_MASTER_SHEET_IN_SOURCE,   # "MASTERDATA_REVAL"
        "MASTERDATA",
        DEST_SHEET_REVAL,               # "REVAL_MASTER_DATA_1"
        "REVAL_MASTER_DATA",
        "MASTERDATA1",
        0                               # last resort: the very first sheet
    ]
    tried = set([preferred])
    for s in candidates:
        if s in tried:
            continue
        tried.add(s)
        try:
            return pd.read_excel(path, sheet_name=s, dtype=object)
        except Exception:
            continue

    # 3) Hard fail with info
    abort(
        f"FAILED TO READ a valid sheet in file: {os.path.basename(path)}. "
        f"Tried sheets (by name or index): {list(tried)}"
    )

def normalize_headers(df):
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out

def drop_trailing_total_row(df):
    if df is None or df.shape[0] == 0:
        return df
    cleaned = df.copy().replace(r'^\s*$', pd.NA, regex=True)
    mask = cleaned.notna().any(axis=1)
    if not mask.any():
        return df.iloc[0:0]
    nonempty = df.loc[mask]
    if len(nonempty) <= 1:
        return df.iloc[0:0]
    return nonempty.iloc[:-1].reset_index(drop=True)

# === Date helpers ===
def excel_serial_to_datetime(val):
    try:
        if isinstance(val,(int,float)) and not pd.isna(val):
            return pd.to_datetime(val, unit="D", origin="1899-12-30", errors="coerce")
    except:
        pass
    return pd.NaT

def _fmt_ddmmyyyy(ts) -> str:
    return f"{ts.day:02d}/{ts.month:02d}/{ts.year:04d}"

def _format_to_ddmmyyyy_series(s: pd.Series) -> pd.Series:
    out = []
    for v in s:
        if pd.isna(v):
            out.append("")
            continue
        if isinstance(v, pd.Timestamp):
            out.append(_fmt_ddmmyyyy(v))
            continue
        if isinstance(v,(int,float)):
            dt = excel_serial_to_datetime(v)
            if pd.notna(dt):
                out.append(_fmt_ddmmyyyy(dt))
                continue
        txt = str(v).strip()
        if txt == "":
            out.append("")
            continue
        dt = pd.to_datetime(txt, errors="coerce", dayfirst=True)
        out.append(_fmt_ddmmyyyy(dt) if pd.notna(dt) else txt)
    return pd.Series(out, dtype=object)

def last_day_of_month(year:int, month:int) -> int:
    import calendar
    return calendar.monthrange(year, month)[1]

def build_lastday_ddmmyyyy(year:int, month:int) -> str:
    dd = last_day_of_month(year, month)
    return f"{dd:02d}/{month:02d}/{year:04d}"

def _detect_date_col_name(df: pd.DataFrame):
    up = {c.upper():c for c in df.columns}
    if "DATE" in up:
        return up["DATE"]
    for c in df.columns:
        u = c.upper()
        if "DATE" in u or "PERIOD" in u or "MONTH" in u:
            return c
    return None

def prepare_base_master_ddmmyyyy(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_headers(df)
    comp_col = {c.upper():c for c in df.columns}.get("COMPANY")
    date_col = _detect_date_col_name(df)
    comp = df[comp_col].astype(object) if comp_col is not None else pd.Series([""] * len(df), dtype=object)
    date_txt = _format_to_ddmmyyyy_series(df[date_col]).astype(str) if date_col is not None else pd.Series([""] * len(df), dtype=object)
    keep_cols = [c for c in df.columns if c.upper() not in ("COMPANY","DATE")]
    body = df.loc[:, keep_cols]
    return pd.concat([pd.DataFrame({"COMPANY":comp,"DATE":date_txt}), body], axis=1)

def mark_date_as_text(df: pd.DataFrame) -> pd.DataFrame:
    if "DATE" in df.columns:
        s = df["DATE"].fillna("").astype(str)
        df = df.copy()
        df["DATE"] = s.apply(lambda x: x if x == "" else ("'"+x))
    return df

def align_to_schema(df: pd.DataFrame, schema_cols: list) -> pd.DataFrame:
    """
    Align df to schema_cols using std_header-based matching.
    This allows IFS exports and Python-generated masterdata to be merged
    even if their headers differ in case/spacing/punctuation
    (e.g. 'Account', 'ACCOUNT NO', 'Account No.' -> 'ACCOUNT').
    """
    n = len(df)
    if n == 0:
        # boş dataframe ama kolonlar schema ile aynı
        return pd.DataFrame({c: pd.Series([], dtype=object) for c in schema_cols})

    # df tarafında std_header -> orijinal kolon adı map'i
    df_cols = list(df.columns)
    df_std_map = {}
    for c in df_cols:
        key = std_header(c)
        if key not in df_std_map:  # aynı key'e sahip birden fazla kolon varsa ilkini al
            df_std_map[key] = c

    aligned = {}
    for sc in schema_cols:
        key = std_header(sc)
        if key in df_std_map:
            aligned[sc] = df[df_std_map[key]]
        else:
            aligned[sc] = pd.Series([""] * n, dtype=object)

    return pd.DataFrame(aligned)

# === Header std ===
def std_header(txt: str) -> str:
    t = re.sub(r'\s+',' ',str(txt).strip()).upper()
    t = re.sub(r'[^A-Z0-9]+','_',t)
    t = re.sub(r'_+','_',t).strip('_')
    return t or "COL"

def uniquify(names):
    seen = {}
    out = []
    for n in names:
        base = n
        cnt = seen.get(base,0)
        if cnt == 0 and base not in out:
            out.append(base)
            seen[base] = 1
        else:
            k = cnt + 1
            cand = f"{base}_{k}"
            while cand in out:
                k += 1
                cand = f"{base}_{k}"
            out.append(cand)
            seen[base] = k
    return out

# === xlwings ===
def import_xlwings():
    try:
        import xlwings as xw
        return xw
    except Exception as e:
        abort(f"xlwings is required. Install with: pip install xlwings pywin32\nDETAILS: {e}")

def set_perf_mode(app, enable_fast: bool):
    try:
        if enable_fast:
            app.api.Calculation = -4135  # xlCalculationManual
            app.api.EnableEvents = False
            app.screen_updating = False
            app.display_alerts = False
            try:
                app.api.AutoCorrect.AutoFillFormulasInLists = False
            except:
                pass
        else:
            app.api.Calculation = -4105  # xlCalculationAutomatic
            app.api.EnableEvents = True
            app.screen_updating = True
            app.display_alerts = True
            try:
                app.api.AutoCorrect.AutoFillFormulasInLists = True
            except:
                pass
            try:
                app.calculate()
            except:
                pass
    except:
        pass

def set_gate_value(wb, value: str):
    try:
        wb.sheets[GATE_SHEET].range(GATE_CELL).value = value
    except:
        pass

def get_or_create_sheet(wb, name: str):
    for s in wb.sheets:
        if s.name.strip().upper() == name.strip().upper():
            return s
    return wb.sheets.add(name=name, after=wb.sheets[-1])

def find_listobject_by_name(sheet, table_name: str):
    try:
        los = sheet.api.ListObjects
        for i in range(1, int(los.Count)+1):
            lo = los.Item(i)
            if str(lo.Name).strip().upper() == table_name.strip().upper():
                return lo
    except:
        pass
    return None

def create_table_on_used_range(sheet, table_name: str):
    used = sheet.range("A1").expand()
    lo = sheet.api.ListObjects.Add(1, used.api, None, 1)
    lo.Name = table_name
    return lo

def ensure_table_masterdata1(wb, sheet_name: str, table_name: str):
    sht = get_or_create_sheet(wb, sheet_name)
    lo = find_listobject_by_name(sht, table_name)
    if lo is None:
        sht.range("A1").value = ["COMPANY","DATE"]
        lo = create_table_on_used_range(sht, table_name)
    return sht, lo

def read_table_headers(sheet, lo):
    hdr = lo.HeaderRowRange
    tl_row = int(hdr.Row)
    tl_col = int(hdr.Column)
    ncols = int(lo.ListColumns.Count)
    vals = sheet.range((tl_row, tl_col),(tl_row, tl_col+ncols-1)).value
    if not isinstance(vals, list):
        vals = [vals]
    headers = [(v if v is not None else "") for v in vals]
    return tl_row, tl_col, headers

def table_size(lo):
    try:
        nrows = int(lo.ListRows.Count)
    except:
        nrows = 0
    try:
        ncols = int(lo.ListColumns.Count)
    except:
        ncols = 0
    return nrows, ncols

def _find_col_index(headers, name):
    up = [str(h).strip().upper() for h in headers]
    key = name.strip().upper()
    try:
        return up.index(key)
    except ValueError:
        return None

def _enforce_date_text_format(sheet, tl_row, tl_col, headers, total_rows):
    try:
        idx = _find_col_index(headers,"DATE")
        if idx is None:
            return
        col_ix = tl_col + idx
        rng = sheet.range((tl_row, col_ix),(tl_row+total_rows-1, col_ix))
        rng.number_format = "@"
    except:
        pass

def _to_writeable_values(df: pd.DataFrame):
    out = df.copy()
    if "DATE" in out.columns:
        s = out["DATE"].astype(str)
        out["DATE"] = s.apply(lambda x: x if x == "" else ("'"+x))
    return out.values.tolist()

def resize_table_and_write(sheet, lo, new_headers, df_values):
    tl_row, tl_col, _ = read_table_headers(sheet, lo)
    n_rows = len(df_values)
    n_cols = len(new_headers)
    total_rows = max(1,n_rows) + 1
    target = sheet.range((tl_row, tl_col),(tl_row+total_rows-1, tl_col+n_cols-1))
    lo.ShowHeaders = True
    lo.ShowTotals = False
    lo.Resize(target.api)
    sheet.range((tl_row, tl_col),(tl_row, tl_col+n_cols-1)).value = new_headers
    _enforce_date_text_format(sheet, tl_row, tl_col, new_headers, total_rows)
    if n_rows > 0:
        sheet.range((tl_row+1, tl_col)).options(index=False, header=False).value = df_values
    else:
        try:
            if lo.DataBodyRange is not None:
                lo.DataBodyRange.ClearContents()
        except:
            pass

# === Export helpers ===
def _resolve_xlsx_engine():
    try:
        import xlsxwriter
        return "xlsxwriter"
    except Exception:
        try:
            import openpyxl
            return "openpyxl"
        except Exception:
            return None

XLSX_ENGINE = _resolve_xlsx_engine()

def excel_writer(path: str):
    if XLSX_ENGINE is None:
        abort("No Excel writer engine found. Please install 'XlsxWriter' or 'openpyxl'.")
    return pd.ExcelWriter(path, engine=XLSX_ENGINE)

def export_masterdata_table(sheet, lo, export_path: str):
    hdr = lo.HeaderRowRange
    tl_row, tl_col = int(hdr.Row), int(hdr.Column)
    ncols = int(lo.ListColumns.Count)
    headers = sheet.range((tl_row, tl_col),(tl_row, tl_col+ncols-1)).value
    if not isinstance(headers, list):
        headers = [headers]
    try:
        nrows = int(lo.ListRows.Count)
    except:
        nrows = 0
    if nrows > 0:
        body = sheet.range((tl_row+1, tl_col),(tl_row+nrows, tl_col+ncols-1)).value
        if nrows == 1 and not isinstance(body, list):
            body = [body]
        df = pd.DataFrame(body, columns=headers)
    else:
        df = pd.DataFrame(columns=headers)
    if "DATE" in df.columns:
        df["DATE"] = df["DATE"].apply(
            lambda x: str(x)[1:] if isinstance(x,str) and x.startswith("'") else x
        )
    with excel_writer(export_path) as w:
        # IMPORTANT CHANGE: write with the same sheet name as REVAL_MASTER_SHEET_IN_SOURCE
        # so that Python-generated MASTERDATA can be used as next month's input.
        df.to_excel(w, index=False, sheet_name=REVAL_MASTER_SHEET_IN_SOURCE)

# === Formatting helpers (kept) ===
def apply_number_format_for_headers(sheet, lo, header_to_fmt: dict):
    tl_row, tl_col, headers = read_table_headers(sheet, lo)
    nrows,_ = table_size(lo)
    total_rows = (nrows if nrows > 0 else 1) + 1
    up = [std_header(h) for h in headers]
    for name,fmt in header_to_fmt.items():
        key = std_header(name)
        if key in up:
            idx = up.index(key)
            col_ix = tl_col + idx
            try:
                sheet.range((tl_row+1, col_ix),(tl_row+total_rows-1, col_ix)).number_format = fmt
            except:
                pass

def uppercase_text_columns(sheet, lo, columns: list, skip: set=None):
    skip = skip or set()
    tl_row, tl_col, headers = read_table_headers(sheet, lo)
    nrows,_ = table_size(lo)
    if nrows <= 0:
        return
    up = [std_header(h) for h in headers]
    for name in columns:
        key = std_header(name)
        if key in skip:
            continue
        if key in up:
            idx = up.index(key)
            col_ix = tl_col + idx
            rng = sheet.range((tl_row+1, col_ix),(tl_row+nrows, col_ix))
            vals = rng.value
            if nrows == 1 and not isinstance(vals, list):
                vals = [vals]
            if isinstance(vals, list) and (len(vals) == nrows) and (not isinstance(vals[0], list)):
                vals = [[v] for v in vals]
            for i in range(len(vals)):
                v = vals[i][0]
                if v is None:
                    continue
                vals[i][0] = str(v).upper()
            rng.value = vals
            try:
                sheet.range((tl_row, col_ix),(tl_row+nrows, col_ix)).number_format = "@"
            except:
                pass

# === COA helpers ===
def standardize_headers_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [std_header(c) for c in out.columns]
    return out

def standardize_coa_dataframe(df_coa: pd.DataFrame) -> pd.DataFrame:
    df = standardize_headers_df(df_coa)
    for col in df.columns:
        if col in {"ACCOUNT","ACCOUNT_CODE","GL_ACCOUNT","GL_ACCOUNT_CODE"}:
            continue
        df[col] = df[col].apply(lambda x: (str(x).upper() if pd.notna(x) else x))
    return df

def to_digit_key(val):
    if pd.isna(val):
        return ""
    try:
        f = float(val)
        if f.is_integer():
            return str(int(f))
        return str(int(round(f)))
    except Exception:
        s = str(val).strip()
        return re.sub(r'\D+', '', s)

def find_best_column(df: pd.DataFrame, priority: list, contains_any: list = None):
    cols = list(df.columns)
    for c in priority:
        if c in cols:
            return c
    if contains_any:
        for c in cols:
            u = c.upper()
            if all(tok.upper() in u for tok in contains_any):
                return c
        for c in cols:
            u = c.upper()
            if any(tok.upper() in u for tok in contains_any):
                return c
    return None

def build_coa_maps(df_coa_std: pd.DataFrame):
    acc_col = find_best_column(
        df_coa_std,
        ["ACCOUNT","ACCOUNT_CODE","GL_ACCOUNT","GL_ACCOUNT_CODE"],
        ["ACCOUNT"]
    )
    g_col   = find_best_column(
        df_coa_std,
        ["ACCOUNT_GROUP_DESCRIPTION","GROUP_DESCRIPTION","ACCOUNT_GROUP"],
        ["GROUP"]
    )
    t_col   = find_best_column(
        df_coa_std,
        ["ACCOUNT_TYPE","TYPE"],
        ["TYPE"]
    )
    print(f"         COA columns detected -> ACCOUNT: {acc_col}, GROUP: {g_col}, TYPE: {t_col}")
    m_group, m_type = {}, {}
    if not acc_col:
        return m_group, m_type
    for _, row in df_coa_std.iterrows():
        key = to_digit_key(row.get(acc_col, None))
        if not key:
            continue
        if g_col and pd.notna(row.get(g_col, None)):
            m_group[key] = str(row.get(g_col))
        if t_col and pd.notna(row.get(t_col, None)):
            m_type[key] = str(row.get(t_col))
    return m_group, m_type

def normalize_account_key(x):
    return to_digit_key(x)

# === Table std & ensure cols ===
def standardize_table_headers_in_place(sheet, lo):
    tl_row, tl_col, headers = read_table_headers(sheet, lo)
    std = [std_header(h) for h in headers]
    std = uniquify(std)
    if std != headers:
        sheet.range((tl_row, tl_col),(tl_row, tl_col+len(std)-1)).value = std
        print(f"         Headers standardized -> {std}")
    else:
        print("         Headers already standardized")
    return std

def ensure_columns_exist_at_end(sht, lo, names: list):
    tl_row, tl_col, headers = read_table_headers(sht, lo)
    up = [std_header(h) for h in headers]
    missing = [std_header(n) for n in names if std_header(n) not in up]
    if not missing:
        return headers
    try:
        for name in missing:
            lo.ListColumns.Add()
            lo.ListColumns(int(lo.ListColumns.Count)).Name = name
        _,_,h2 = read_table_headers(sht, lo)
        return h2
    except Exception as e:
        print(f"         WARN: ListColumns.Add failed ({e}). Fallback to Resize.")
    try:
        cur_rows, cur_cols = table_size(lo)
        total_rows = (cur_rows if cur_rows > 0 else 1) + 1
        k = len(missing)
        cur_len = len(headers)
        target = sht.range((tl_row, tl_col),(tl_row+total_rows-1, tl_col+cur_len+k-1))
        lo.Resize(target.api)
        for j,name in enumerate(missing, start=0):
            sht.range((tl_row, tl_col+cur_len+j)).value = name
        _,_,h3 = read_table_headers(sht, lo)
        return h3
    except Exception as e2:
        abort(f"FAILED TO ADD COLUMNS BY RESIZE: {e2}")

def read_table_as_dataframe(sht, lo) -> pd.DataFrame:
    tl_row, tl_col, headers = read_table_headers(sht, lo)
    nrows, ncols = table_size(lo)
    if ncols == 0:
        return pd.DataFrame(columns=[])
    if nrows > 0:
        body = sht.range((tl_row+1, tl_col),(tl_row+nrows, tl_col+ncols-1)).value
        if nrows == 1 and not isinstance(body, list):
            body = [body]
        df = pd.DataFrame(body, columns=headers)
    else:
        df = pd.DataFrame(columns=headers)
    df.columns = [std_header(c) for c in df.columns]
    return df

def find_account_header_idx(up_headers: list) -> int:
    for i,h in enumerate(up_headers):
        if h == "ACCOUNT":
            return i
    for i,h in enumerate(up_headers):
        if re.fullmatch(r"ACCOUNT(_\d+)?", h) and h != "ACCOUNT_DESCRIPTION":
            return i
    return -1

# === IS_IN_ANALYSIS fill (new) ===
def fill_and_write_analysis_cols(sht, lo, map_acc_to_group, map_acc_to_type, override_is_in: dict):
    df = read_table_as_dataframe(sht, lo)
    if df.shape[0] == 0:
        return
    up_cols = list(df.columns)
    acc_idx = find_account_header_idx(up_cols)
    if acc_idx < 0:
        return
    acc_keys = df.iloc[:, acc_idx].apply(to_digit_key)
    group_vals = acc_keys.map(map_acc_to_group).where(pd.notna(acc_keys), None)
    type_vals  = acc_keys.map(map_acc_to_type).where(pd.notna(acc_keys), None)

    def map_isin(k):
        if not k:
            return None
        if k in (override_is_in or {}):
            return override_is_in[k]
        if k in DICT_IS_IN_ANALYSIS:
            return DICT_IS_IN_ANALYSIS[k]
        return None

    isin_vals = acc_keys.apply(map_isin)
    ensure_columns_exist_at_end(sht, lo, ["ACCOUNT_GROUP_DESCRIPTION","ACCOUNT_TYPE","IS_IN_ANALYSIS"])
    tl_row, tl_col, hdrs_now = read_table_headers(sht, lo)
    hdrs_now_up = [std_header(h) for h in hdrs_now]
    nrows, _ = table_size(lo)

    def col_ix(col_name):
        key = std_header(col_name)
        if key not in hdrs_now_up:
            return None
        return tl_col + hdrs_now_up.index(key)

    g_ix = col_ix("ACCOUNT_GROUP_DESCRIPTION")
    t_ix = col_ix("ACCOUNT_TYPE")
    i_ix = col_ix("IS_IN_ANALYSIS")

    def write_col(ix, series):
        if ix is None:
            return
        rng = sht.range((tl_row+1, ix),(tl_row+nrows, ix))
        rng.value = [[(None if (pd.isna(v) or v == "") else v)] for v in series.tolist()]
        try:
            sht.range((tl_row, ix),(tl_row+nrows, ix)).number_format = "@"
        except:
            pass

    write_col(g_ix, group_vals)
    write_col(t_ix, type_vals)
    write_col(i_ix, isin_vals)

    m_coa = int((group_vals.notna() & (group_vals != "")).sum())
    m_is  = int((isin_vals.notna() & (isin_vals != "")).sum())
    print(f"         COA enrichment done: {m_coa}/{len(df)} rows filled (GROUP/TYPE).")
    print(f"         IS_IN_ANALYSIS fill: {m_is}/{len(df)} rows assigned.")
    try:
        ex = pd.DataFrame({
            "ACCOUNT_KEY": acc_keys.head(5),
            "GROUP": group_vals.head(5),
            "TYPE": type_vals.head(5),
            "IS_IN": isin_vals.head(5)
        })
        print("         Sample (first 5):")
        for _,r in ex.iterrows():
            print(f"             {r['ACCOUNT_KEY']} -> G={r['GROUP']} | T={r['TYPE']} | IS_IN={r['IS_IN']}")
    except:
        pass

def prompt_show_current_setup(df_coa_std, df_master_preview_accounts):
    accs = set()
    accs.update([a for a in DICT_IS_IN_ANALYSIS.keys()])
    if df_master_preview_accounts is not None:
        for a in df_master_preview_accounts:
            if a:
                accs.add(a)
    desc_map = {}
    if df_coa_std is not None and not df_coa_std.empty:
        acc_col = find_best_column(
            df_coa_std,
            ["ACCOUNT","ACCOUNT_CODE","GL_ACCOUNT","GL_ACCOUNT_CODE"],
            ["ACCOUNT"]
        )
        desc_col = find_best_column(
            df_coa_std,
            ["ACCOUNT_DESCRIPTION","DESCRIPTION","NAME"],
            ["DESC","NAME"]
        )
        if acc_col and desc_col:
            for _, row in df_coa_std.iterrows():
                k = to_digit_key(row.get(acc_col, None))
                if k:
                    d = row.get(desc_col, None)
                    if pd.notna(d):
                        desc_map[k] = str(d)
    rows = []
    for a in sorted(accs):
        rows.append({
            "ACCOUNT": a,
            "ACCOUNT_DESCRIPTION": desc_map.get(a, DICT_ACCOUNT_DESC_FALLBACK.get(a,"")),
            "IS_IN_ANALYSIS": DICT_IS_IN_ANALYSIS.get(a, "")
        })
    if rows:
        print("\n--- CURRENT SETUP (ACCOUNT | ACCOUNT_DESCRIPTION | IS_IN_ANALYSIS) ---")
        for r in rows[:200]:
            print(f"{r['ACCOUNT']:>10} | {str(r['ACCOUNT_DESCRIPTION'])[:50]:<50} | {r['IS_IN_ANALYSIS']}")
        if len(rows) > 200:
            print(f"... and {len(rows)-200} more rows")
        print("---------------------------------------------------------------------\n")
    else:
        print("\n(No rows to display for current setup)\n")

def interactive_override_is_in(df_after, df_coa_std):
    override = {}
    use_current = prompt_yesno(
        "WOULD YOU WANNA USE THE CURRENT SET UP TO DECIDED ON WHICH ACCOUNTS WILL BE CONSIDERED IN THE REPORT? (YES/NO): "
    )
    see_setup = prompt_yesno(
        "WOULD YOU WANNA SEE THE CURRENT SET UP TO CHECK? (YES/NO): "
    )
    if see_setup == "YES":
        acc_col_idx = None
        master_accs = []
        if df_after is not None and not df_after.empty:
            cols_up = [std_header(c) for c in df_after.columns]
            try:
                acc_col_idx = cols_up.index("ACCOUNT")
            except ValueError:
                acc_col_idx = -1
            if acc_col_idx >= 0:
                master_accs = [
                    to_digit_key(v)
                    for v in df_after.iloc[:, acc_col_idx].dropna().astype(str).tolist()
                ][:400]
        import time
        time.sleep(5)
        prompt_show_current_setup(df_coa_std, master_accs)
    if use_current == "NO":
        k = prompt_int("HOW MANY ACCOUNT NUMBER YOU WOULD LIKE TO RELABEL? ")
        for i in range(1, k+1):
            acc = to_digit_key(
                safe_input(
                    f"PLEASE TELL ME WHICH ACCOUNT NUMBER YOU WOULD LIKE TO CHANGE THE CATEGORY FOR? ({i}/{k}): "
                )
            )
            if not acc:
                print("   Skipped (empty).")
                continue
            desc = ""
            if df_coa_std is not None and not df_coa_std.empty:
                acc_col = find_best_column(
                    df_coa_std,
                    ["ACCOUNT","ACCOUNT_CODE","GL_ACCOUNT","GL_ACCOUNT_CODE"],
                    ["ACCOUNT"]
                )
                desc_col = find_best_column(
                    df_coa_std,
                    ["ACCOUNT_DESCRIPTION","DESCRIPTION","NAME"],
                    ["DESC","NAME"]
                )
                if acc_col and desc_col:
                    hit = df_coa_std[df_coa_std[acc_col].map(to_digit_key)==acc]
                    if not hit.empty:
                        d = hit.iloc[0][desc_col]
                        if pd.notna(d):
                            desc = str(d)
            cur = DICT_IS_IN_ANALYSIS.get(acc, "")
            if cur == "":
                cur = "(not in current dict)"
            print(f"   ACCOUNT {acc} | CURRENT: {cur} | DESC: {desc[:70]}")
            lab = safe_input("   SHOULD IT BE CONSIDERED OR NOT_CONSIDERED? ").strip().upper()
            while lab not in {"CONSIDERED","NOT_CONSIDERED"}:
                lab = safe_input("   Please type exactly CONSIDERED or NOT_CONSIDERED: ").strip().upper()
            override[acc] = lab
    unmatched = set()
    if df_after is not None and not df_after.empty:
        cols_up = [std_header(c) for c in df_after.columns]
        try:
            acc_col_idx = cols_up.index("ACCOUNT")
        except ValueError:
            acc_col_idx = -1
        if acc_col_idx >= 0:
            accs = [
                to_digit_key(v)
                for v in df_after.iloc[:, acc_col_idx].dropna().astype(str).tolist()
            ]
            for a in set(accs):
                if (a not in DICT_IS_IN_ANALYSIS) and (a not in override):
                    unmatched.add(a)
    if unmatched:
        print(f"Found {len(unmatched)} accounts not in current dictionary.")
        do_manual = prompt_yesno("DO YOU WANT TO MANUALLY LABEL THESE NEW ACCOUNTS? (YES/NO): ")
        if do_manual == "YES":
            desc_map = {}
            if df_coa_std is not None and not df_coa_std.empty:
                acc_col = find_best_column(
                    df_coa_std,
                    ["ACCOUNT","ACCOUNT_CODE","GL_ACCOUNT","GL_ACCOUNT_CODE"],
                    ["ACCOUNT"]
                )
                desc_col = find_best_column(
                    df_coa_std,
                    ["ACCOUNT_DESCRIPTION","DESCRIPTION","NAME"],
                    ["DESC","NAME"]
                )
                if acc_col and desc_col:
                    for _, row in df_coa_std.iterrows():
                        k = to_digit_key(row.get(acc_col, None))
                        if k:
                            d = row.get(desc_col, None)
                            if pd.notna(d):
                                desc_map[k] = str(d)
            um_list = sorted(list(unmatched))
            for idx, a in enumerate(um_list, start=1):
                d = desc_map.get(a, "")
                print(f"   ({idx}/{len(um_list)}) ACCOUNT {a} | DESC: {d[:70]}")
                lab = safe_input("   IS IT CONSIDERED OR NOT_CONSIDERED? ").strip().upper()
                while lab not in {"CONSIDERED","NOT_CONSIDERED"}:
                    lab = safe_input("   Please type exactly CONSIDERED or NOT_CONSIDERED: ").strip().upper()
                override[a] = lab
            unmatched = set()
        else:
            print("Unmatched accounts left unlabeled; they will appear blank in IS_IN_ANALYSIS.")
    return override, unmatched

# === Stats builder (kept new but simplified naming) ===
def _strip_apostrophe(s):
    if isinstance(s, str) and s.startswith("'"):
        return s[1:]
    return s

def _parse_date_series_ddmmyyyy(s: pd.Series):
    return pd.to_datetime(s.map(_strip_apostrophe), dayfirst=True, errors="coerce")

def _classify_dtype(series: pd.Series) -> str:
    if series.dropna().empty:
        return "EMPTY"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "DATE"
    if pd.api.types.is_numeric_dtype(series):
        return "NUMBER"
    return "TEXT"

def _add_df_as_table(sht, top_row: int, left_col: int, df: pd.DataFrame, table_name: str, title: str):
    sht.range((top_row, left_col)).value = title
    start = (top_row+1, left_col)
    sht.range(start).options(index=False, header=True).value = df
    nrows, ncols = df.shape
    lo = sht.api.ListObjects.Add(
        1,
        sht.range(start, (top_row+1+nrows, left_col+ncols-1)).api,
        None,
        1
    )
    lo.Name = table_name
    try:
        sht.used_range.columns.autofit()
    except:
        pass
    return top_row + nrows + 3

def build_and_write_stats_sheet(
    wb,
    df_base_out,
    base_company_counts,
    updates_info_current,
    updates_info_outdated,
    df_after,
    df_coa_std
):
    for s in wb.sheets:
        if s.name.strip().upper() == STATS_SHEET:
            s.delete()
            break
    sht = wb.sheets.add(name=STATS_SHEET, after=wb.sheets[-1])

    base_summary = pd.DataFrame([
        {"Metric":"Base rows (before updates)","Value": int(df_base_out.shape[0])},
        {"Metric":"Base cols (before updates)","Value": int(df_base_out.shape[1])}
    ])
    base_by_company = base_company_counts.rename(columns={"COMPANY":"COMPANY","ROWS":"ROWS"})
    row_ptr = 1
    row_ptr = _add_df_as_table(sht, row_ptr, 1, base_summary, "STATS_BASE_SUMMARY", "BASE SUMMARY")
    row_ptr = _add_df_as_table(sht, row_ptr, 1, base_by_company, "STATS_BASE_BY_COMPANY", "BASE ROWS BY COMPANY")

    if updates_info_outdated:
        upd_df_o = pd.DataFrame(updates_info_outdated)[
            ["FILE_INDEX","COMPANY","ROWS","COLS","DATE_MIN","DATE_MAX","FILENAME"]
        ]
        row_ptr = _add_df_as_table(
            sht,
            row_ptr,
            1,
            upd_df_o,
            "STATS_UPDATES_OUTDATED",
            "OUT-OF-DATE FILES OVERVIEW"
        )

    if updates_info_current:
        upd_df_c = pd.DataFrame(updates_info_current)[
            ["FILE_INDEX","COMPANY","ROWS","COLS","DATE_MIN","DATE_MAX","FILENAME"]
        ]
        row_ptr = _add_df_as_table(
            sht,
            row_ptr,
            1,
            upd_df_c,
            "STATS_UPDATES_CURRENT",
            "CURRENT FILES OVERVIEW"
        )

    cols_up = list(df_after.columns)
    def col_or_none(name):
        return name if name in cols_up else None

    cur_col  = col_or_none("TRANSACTION_CURRENCY_CODE")
    accd_col = col_or_none("ACCOUNT_DESCRIPTION")
    agd_col  = col_or_none("ACCOUNT_GROUP_DESCRIPTION")
    atp_col  = col_or_none("ACCOUNT_TYPE")
    isin_col = col_or_none("IS_IN_ANALYSIS")

    uniq_currency = int(df_after[cur_col].dropna().nunique()) if cur_col else 0
    uniq_accdesc  = int(df_after[accd_col].dropna().nunique()) if accd_col else 0
    uniq_group    = int(df_after[agd_col].dropna().nunique()) if agd_col else 0
    uniq_type     = int(df_after[atp_col].dropna().nunique()) if atp_col else 0
    n_considered  = int((df_after[isin_col]=="CONSIDERED").sum()) if isin_col else 0
    n_not_cons    = int((df_after[isin_col]=="NOT_CONSIDERED").sum()) if isin_col else 0

    final_summary = pd.DataFrame([
        {"Metric":"Final rows","Value": int(df_after.shape[0])},
        {"Metric":"Final cols","Value": int(df_after.shape[1])},
        {"Metric":"Unique currencies","Value": uniq_currency},
        {"Metric":"Unique account descriptions","Value": uniq_accdesc},
        {"Metric":"Unique account group descriptions","Value": uniq_group},
        {"Metric":"Unique account types","Value": uniq_type},
        {"Metric":"CONSIDERED rows","Value": n_considered},
        {"Metric":"NOT_CONSIDERED rows","Value": n_not_cons},
    ])
    row_ptr = _add_df_as_table(
        sht,
        row_ptr,
        1,
        final_summary,
        "STATS_FINAL_SUMMARY",
        "FINAL MASTER SUMMARY"
    )

    date_series = _parse_date_series_ddmmyyyy(df_after["DATE"]) if "DATE" in df_after.columns else pd.Series(dtype="datetime64[ns]")
    min_final = date_series.min()
    max_final = date_series.max()

    date_rows = []
    if updates_info_outdated:
        for item in updates_info_outdated:
            date_rows.append({
                "DATASET": f"OUTDATED FILE {item['FILE_INDEX']} ({item['COMPANY']})",
                "MIN_DATE": item["DATE_MIN"],
                "MAX_DATE": item["DATE_MAX"]
            })
    if updates_info_current:
        for item in updates_info_current:
            date_rows.append({
                "DATASET": f"CURRENT FILE {item['FILE_INDEX']} ({item['COMPANY']})",
                "MIN_DATE": item["DATE_MIN"],
                "MAX_DATE": item["DATE_MAX"]
            })
    date_rows.append({
        "DATASET": "FINAL MASTER",
        "MIN_DATE": (min_final.strftime("%d/%m/%Y") if pd.notna(min_final) else ""),
        "MAX_DATE": (max_final.strftime("%d/%m/%Y") if pd.notna(max_final) else "")
    })
    df_dates = pd.DataFrame(date_rows)
    row_ptr = _add_df_as_table(
        sht,
        row_ptr,
        1,
        df_dates,
        "STATS_DATE_RANGES",
        "DATE RANGES"
    )

    acc_col = find_best_column(
        df_coa_std,
        ["ACCOUNT","ACCOUNT_CODE","GL_ACCOUNT","GL_ACCOUNT_CODE"],
        ["ACCOUNT"]
    )
    uniq_accounts = 0
    if acc_col:
        uniq_accounts = int(
            pd.Series(df_coa_std[acc_col])
            .map(to_digit_key)
            .replace("", pd.NA)
            .dropna()
            .nunique()
        )

    coa_summary = pd.DataFrame([
        {"Metric":"COA rows","Value": int(df_coa_std.shape[0])},
        {"Metric":"COA cols","Value": int(df_coa_std.shape[1])},
        {"Metric":"Unique ACCOUNT (normalized)","Value": uniq_accounts},
    ])
    row_ptr = _add_df_as_table(
        sht,
        row_ptr,
        1,
        coa_summary,
        "STATS_COA_SUMMARY",
        "CHART OF ACCOUNTS SUMMARY"
    )

    dtypes_master = pd.DataFrame({
        "COLUMN": df_after.columns,
        "DTYPE":   [ _classify_dtype(df_after[c]) for c in df_after.columns ]
    })
    row_ptr = _add_df_as_table(
        sht,
        row_ptr,
        1,
        dtypes_master,
        "STATS_MASTER_DTYPES",
        "FINAL MASTER COLUMN TYPES"
    )

    dtypes_coa = pd.DataFrame({
        "COLUMN": df_coa_std.columns,
        "DTYPE":  [ _classify_dtype(df_coa_std[c]) for c in df_coa_std.columns ]
    })
    row_ptr = _add_df_as_table(
        sht,
        row_ptr,
        1,
        dtypes_coa,
        "STATS_COA_DTYPES",
        "COA COLUMN TYPES"
    )
    try:
        sht.used_range.columns.autofit()
    except:
        pass

    return {
        "base_rows": int(df_base_out.shape[0]),
        "base_cols": int(df_base_out.shape[1]),
        "base_by_company": base_company_counts,
        "updates_current": updates_info_current,
        "updates_outdated": updates_info_outdated,
        "final_rows": int(df_after.shape[0]),
        "final_cols": int(df_after.shape[1]),
        "uniq_currency": uniq_currency,
        "uniq_accdesc": uniq_accdesc,
        "uniq_group": uniq_group,
        "uniq_type": uniq_type,
        "n_considered": n_considered,
        "n_not_considered": n_not_cons,
        "final_min_date": (min_final.strftime("%d/%m/%Y") if pd.notna(min_final) else ""),
        "final_max_date": (max_final.strftime("%d/%m/%Y") if pd.notna(max_final) else ""),
        "coa_rows": int(df_coa_std.shape[0]),
        "coa_cols": int(df_coa_std.shape[1]),
        "coa_uniq_accounts": uniq_accounts
    }

# === MAIN (kept structure from EXPOSURE AUTOMATION CODE 2) ===
INTRO_BANNER = r"""
============================================================================
BALANCE SHEET EXPOSURE GENERATOR - PYTHON ENGINE V1.0 
TREASURY REPORTING TEAM
============================================================================
============================================================================
INFORMATION: 

This Python engine is specifically designed to automate the Balance Exposure
reporting procedure to minimize human interaction and to increase the efficiency
and accuracy of our reports.

Please answer the questions below and wait for python to finish the task for you. 
============================================================================
ARE YOU READY? 
"""

def main():
    if os.name != "nt":
        abort("This application currently targets Windows + Excel (COM) environment.")
    print(INTRO_BANNER)
    ready = prompt_yesno("(Type YES to start, NO to exit): ")
    if ready != "YES":
        print("Exit requested. Goodbye.")
        sys.exit(0)

    print("===================================================")
    print("   REVALUATION → BALANCE SHEET EXPOSURE PYTHON ENGINE   ")
    print("===================================================")

    reval_master_path = prompt_path(
        'PLEASE PROVIDE THE PATH OF THE CURRENT MASTERDATA FOR REVALUATION: ',
        True,
        VALID_INPUT_EXTS
    )
    slog("CURRENT MASTER DATA recognized")

    reporting_path    = prompt_path(
        'PLEASE PROVIDE THE PATH OF THE CURRENT REPORTING FILE: ',
        True,
        VALID_REPORT_EXTS
    )
    slog("CURRENT REPORTING FILE recognized")

    coa_path          = prompt_path(
        'PLEASE PROVIDE THE PATH OF THE CHART OF ACCOUNTS: ',
        True,
        VALID_INPUT_EXTS
    )
    slog("CHART OF ACCOUNTS recognized")

    outdated_answer = prompt_yesno('DO YOU HAVE FILES WHICH ARE NOT UP TO DATE? (YES/NO): ')
    outdated_items = []
    if outdated_answer == "YES":
        n_outdated = prompt_int('HOW MANY FILES ARE NOT UP TO DATE? ')
        for i in range(1, n_outdated+1):
            p = prompt_path(
                f'PLEASE PROVIDE THE FULL PATH OF THE FILE {i} TO UPDATE THE REVALUATION MASTER DATA: ',
                True,
                VALID_INPUT_EXTS
            )
            y = prompt_year(f'PLEASE PROVIDE THE YEAR INFORMATION FOR THE FILE {i}? ')
            m = prompt_month(f'PLEASE PROVIDE THE MONTH INFORMATION FOR THE FILE {i}: ')
            lastday_text_i = build_lastday_ddmmyyyy(y, m)
            comp = detect_company_from_filename(p)
            outdated_items.append({
                "idx": i,
                "path": p,
                "year": y,
                "month": m,
                "lastday": lastday_text_i,
                "company": comp
            })
            print(
                f"         Outdated file {i} queued: "
                f"{os.path.basename(p)} | {comp} | {lastday_text_i}"
            )

    n_updates_current = prompt_int(
        'HOW MANY FILES YOU HAVE TO UPDATE MASTER DATA AND GENERATE THE CURRENT BALANCE SHEET EXPOSURE REPORT? '
    )
    update_paths_current = []
    for i in range(1, n_updates_current+1):
        p = prompt_path(
            f'PLEASE PROVIDE THE PATH OF THE FILE {i} TO UPDATE THE REVALUATION MASTER DATA: ',
            True,
            VALID_INPUT_EXTS
        )
        update_paths_current.append(p)
        print(f"         Current file {i} queued")

    year  = prompt_year('WHAT IS THE YEAR OF THE REPORT? ')
    month = prompt_month('WHAT IS THE MONTH THAT YOU ARE CREATING THE REPORT FOR? ')
    month_abbr = MMM[month-1]
    lastday_text_current = build_lastday_ddmmyyyy(year, month)
    date_tag_for_file = f"{year}{month_abbr.upper()}"
    slog(f"CURRENT UPDATE DATE (month-end) will be {lastday_text_current}")

    slog("Reading BASE revaluation sheet")
    # IMPORTANT CHANGE: flexible sheet reader
    df_base = read_excel_sheet_auto(reval_master_path, REVAL_MASTER_SHEET_IN_SOURCE)
    df_base = normalize_headers(df_base)
    print(f"         BASE read: {df_base.shape[0]} rows, {df_base.shape[1]} cols")
    df_base_out = prepare_base_master_ddmmyyyy(df_base)
    base_cols = list(df_base_out.columns)
    print(f"         BASE schema locked ({len(base_cols)} cols)")
    base_company_counts = (
        df_base_out
        .groupby("COMPANY", dropna=False)["DATE"]
        .count()
        .reset_index()
        .rename(columns={"DATE":"ROWS"})
    )

    updates_aligned_outdated = []
    updates_info_outdated = []
    for item in outdated_items:
        p = item["path"]
        comp = item["company"]
        lastday = item["lastday"]
        i = item["idx"]
        slog(f"Outdated {i}: COMPANY -> {comp}")
        df_u = normalize_headers(read_excel_first_sheet(p))
        ru0, cu0 = df_u.shape
        print(f"         Outdated {i} read: {ru0} rows, {cu0} cols")
        df_u = drop_trailing_total_row(df_u)
        if df_u.shape[0] != ru0:
            print(f"         Outdated {i}: TOTAL row dropped -> {df_u.shape[0]} rows")
        front = pd.DataFrame({
            "COMPANY": [comp]*len(df_u),
            "DATE": [lastday]*len(df_u)
        })
        df_u2 = pd.concat([front, df_u.reset_index(drop=True)], axis=1)
        df_u2 = align_to_schema(df_u2, base_cols)
        updates_aligned_outdated.append(df_u2)
        print(f"         Outdated {i}: aligned rows -> {df_u2.shape[0]}")
        updates_info_outdated.append({
            "FILE_INDEX": i,
            "COMPANY": comp,
            "ROWS": int(df_u.shape[0]),
            "COLS": int(df_u.shape[1]),
            "DATE_MIN": lastday,
            "DATE_MAX": lastday,
            "FILENAME": os.path.basename(p)
        })

    updates_aligned_current = []
    updates_info_current = []
    for i,p in enumerate(update_paths_current, start=1):
        comp = detect_company_from_filename(p)
        slog(f"Current {i}: COMPANY -> {comp}")
        df_u = normalize_headers(read_excel_first_sheet(p))
        ru0, cu0 = df_u.shape
        print(f"         Current {i} read: {ru0} rows, {cu0} cols")
        df_u = drop_trailing_total_row(df_u)
        if df_u.shape[0] != ru0:
            print(f"         Current {i}: TOTAL row dropped -> {df_u.shape[0]} rows")
        front = pd.DataFrame({
            "COMPANY": [comp]*len(df_u),
            "DATE": [lastday_text_current]*len(df_u)
        })
        df_u2 = pd.concat([front, df_u.reset_index(drop=True)], axis=1)
        df_u2 = align_to_schema(df_u2, base_cols)
        updates_aligned_current.append(df_u2)
        print(f"         Current {i}: aligned rows -> {df_u2.shape[0]}")
        updates_info_current.append({
            "FILE_INDEX": i,
            "COMPANY": comp,
            "ROWS": int(df_u.shape[0]),
            "COLS": int(df_u.shape[1]),
            "DATE_MIN": lastday_text_current,
            "DATE_MAX": lastday_text_current,
            "FILENAME": os.path.basename(p)
        })

    pieces = [df_base_out]
    if updates_aligned_outdated:
        pieces += updates_aligned_outdated
    if updates_aligned_current:
        pieces += updates_aligned_current
    df_final = pd.concat(pieces, ignore_index=True) if len(pieces) > 1 else df_base_out.copy()
    df_final = mark_date_as_text(df_final)

    import shutil
    copy_path = unique_copy_path_same_folder(reporting_path)
    try:
        shutil.copy2(reporting_path, copy_path)
    except Exception as e:
        abort(f"FAILED TO CREATE WORKING COPY: {e}")
    slog(f"Working copy created -> {copy_path}")

    xw = import_xlwings()
    try:
        app = xw.App(visible=True, add_book=False)
        wb = app.books.open(copy_path)
    except Exception as e:
        abort(f"FAILED TO OPEN WORKING COPY IN EXCEL: {e}")

    try:
        set_gate_value(wb, "IN PROGRESS")
        set_perf_mode(app, True)
        slog("Gate=IN PROGRESS; fast mode ON")

        df_coa_std = standardize_coa_dataframe(read_excel_first_sheet(coa_path))
        for s in wb.sheets:
            if s.name.strip().upper() == COA_DEST_SHEET:
                s.delete()
                break
        sht_coa = wb.sheets.add(name=COA_DEST_SHEET, after=wb.sheets[-1])
        sht_coa.range("A1").options(index=False, header=True).value = df_coa_std
        used = sht_coa.range("A1").expand()
        lo_coa = sht_coa.api.ListObjects.Add(1, used.api, None, 1)
        lo_coa.Name = COA_TABLE_BASENAME
        print(
            f"         COA standardized & written: "
            f"{df_coa_std.shape[0]} rows, {df_coa_std.shape[1]} cols"
        )
        map_acc_to_group, map_acc_to_type = build_coa_maps(df_coa_std)
        slog(f"COA maps ready (group={len(map_acc_to_group)}, type={len(map_acc_to_type)})")

        sht, lo = ensure_table_masterdata1(wb, DEST_SHEET_REVAL, DEST_TABLE_NAME)
        slog(f"Target sheet/table ready: {DEST_SHEET_REVAL}/{DEST_TABLE_NAME}")
        cols = list(df_final.columns)
        vals = _to_writeable_values(df_final)
        resize_table_and_write(sht, lo, cols, vals)
        print(
            f"         MASTERDATA1 updated -> "
            f"Rows: {df_final.shape[0]}, Cols: {df_final.shape[1]}"
        )
        std_hdrs = standardize_table_headers_in_place(sht, lo)
        slog(f"Header standardization complete ({len(std_hdrs)} columns)")

        df_after_preview = read_table_as_dataframe(sht, lo)
        override_is_in, unmatched = interactive_override_is_in(df_after_preview, df_coa_std)

        ensure_columns_exist_at_end(
            sht,
            lo,
            ["ACCOUNT_GROUP_DESCRIPTION","ACCOUNT_TYPE","IS_IN_ANALYSIS"]
        )
        slog("Additional columns ensured (GROUP/TYPE/IS_IN_ANALYSIS)")
        fill_and_write_analysis_cols(sht, lo, map_acc_to_group, map_acc_to_type, override_is_in)

        text_cols_upper = [
            "COMPANY",
            "ACCOUNT_DESCRIPTION",
            "CURRENCY_RATE_TYPE_DESCRIPTION",
            "NOT_REVERSED",
            "REVALUATION_LEVEL",
            "REVALUATION_METHOD",
            "USED_IN_LEDGER",
            "ACCOUNT_GROUP_DESCRIPTION",
            "ACCOUNT_TYPE",
            "IS_IN_ANALYSIS"
        ]
        uppercase_text_columns(sht, lo, text_cols_upper, skip={"DATE"})
        header_to_fmt = {
            "ACCOUNT": "0",
            "CURRENCY_AMOUNT": "#.##0,00;-#.##0,00",
            "REVALUED_AMOUNT": "#.##0,00;-#.##0,00",
            "AMOUNT_BEFORE_REVALUATION": "#.##0,00;-#.##0,00",
            "UNREALIZED_GAIN": "#.##0,00;-#.##0,00",
            "UNREALIZED_LOSS": "#.##0,00;-#.##0,00",
            "NET_CURRENCY_GAIN_LOSS": "#.##0,00;-#.##0,00",
            "CURRENCY_RATE": "0,000000",
            "CONVERSION_FACTOR": "0,000000",
            "CURRENCY_RATE_TYPE": "0"
        }
        apply_number_format_for_headers(sht, lo, header_to_fmt)
        tl_row2, tl_col2, headers2 = read_table_headers(sht, lo)
        _enforce_date_text_format(
            sht,
            tl_row2,
            tl_col2,
            headers2,
            (table_size(lo)[0] if table_size(lo)[0] > 0 else 1) + 1
        )

        df_after = read_table_as_dataframe(sht, lo)
        stats_small = build_and_write_stats_sheet(
            wb=wb,
            df_base_out=df_base_out,
            base_company_counts=base_company_counts,
            updates_info_current=updates_info_current,
            updates_info_outdated=updates_info_outdated,
            df_after=df_after,
            df_coa_std=df_coa_std
        )

        print("--------- DATA PROCESS STATS (SUMMARY) ---------")
        print(f"Base: rows={stats_small['base_rows']}, cols={stats_small['base_cols']}")
        print("Base by company:")
        for _,r in stats_small["base_by_company"].iterrows():
            print(f"   {r['COMPANY']}: {int(r['ROWS'])} rows")
        if stats_small["updates_outdated"]:
            print("Out-of-date files:")
            for u in stats_small["updates_outdated"]:
                print(
                    f"   File {u['FILE_INDEX']} ({u['COMPANY']}): "
                    f"rows={u['ROWS']}, cols={u['COLS']}, "
                    f"date={u['DATE_MIN']}..{u['DATE_MAX']} [{u['FILENAME']}]"
                )
        if stats_small["updates_current"]:
            print("Current files:")
            for u in stats_small["updates_current"]:
                print(
                    f"   File {u['FILE_INDEX']} ({u['COMPANY']}): "
                    f"rows={u['ROWS']}, cols={u['COLS']}, "
                    f"date={u['DATE_MIN']}..{u['DATE_MAX']} [{u['FILENAME']}]"
                )
        print(
            f"Final master: rows={stats_small['final_rows']}, "
            f"cols={stats_small['final_cols']}"
        )
        print(
            f"Unique currencies={stats_small['uniq_currency']}, "
            f"acct desc={stats_small['uniq_accdesc']}, "
            f"group desc={stats_small['uniq_group']}, "
            f"acct type={stats_small['uniq_type']}"
        )
        print(
            f"CONSIDERED={stats_small['n_considered']}, "
            f"NOT_CONSIDERED={stats_small['n_not_considered']}"
        )
        print(
            f"Final date range: "
            f"{stats_small['final_min_date']} .. {stats_small['final_max_date']}"
        )
        print(
            f"COA: rows={stats_small['coa_rows']}, cols={stats_small['coa_cols']}, "
            f"unique ACCOUNT(normalized)={stats_small['coa_uniq_accounts']}"
        )
        print("------------------------------------------------")

        set_gate_value(wb, "DONE")
        set_perf_mode(app, False)
        slog("Gate=DONE; fast mode OFF")
        wb.save()
        slog("Working copy saved")

        engine = _resolve_xlsx_engine()
        if engine is None:
            abort("No Excel writer engine found. Please install 'XlsxWriter' or 'openpyxl'.")
        export_name = f"MASTERDATA_{date_tag_for_file}_PYTHON_GENERATED.xlsx"
        export_path = os.path.join(os.path.dirname(copy_path), export_name)
        export_masterdata_table(sht, lo, export_path)
        slog(f"Exported MASTERDATA -> {export_path}")

    except Exception as e:
        # close excel gracefully but keep error info for debugging
        try:
            wb.close(save_changes=False)
            app.quit()
        finally:
            write_error_and_notify(e)
            abort(f"UNEXPECTED ERROR DURING EXCEL WRITE: {e}")
    finally:
        try:
            wb.close()
        except:
            pass
        try:
            app.quit()
        except:
            pass

    print("===================================================")
    print("                 PROCESS COMPLETED                 ")
    print("===================================================")
    time.sleep(1)
    for _ in range(3):
        print("")
    print("DESIGNED AND PRODUCED BY TREASURY REPORTING TEAM - YINSON PRODUCTION NETHERLANDS BV")
    time.sleep(1)
    print("UYGAR TALU")
    time.sleep(1)


# --------------------------
# Entrypoint wrapper (keeps freeze_support and global exception capture)
# --------------------------
def main_wrapper():
    try:
        # Multiprocessing freeze support for PyInstaller on Windows (both original files had this)
        try:
            import multiprocessing as _mp
            _mp.freeze_support()
        except Exception:
            pass
        main()
    except KeyboardInterrupt:
        abort("OPERATION CANCELLED BY USER (KeyboardInterrupt)")
    except Exception as e:
        write_error_and_notify(e)
        # keep console open so remote user can see prompt if launched by double-click
        try:
            input("An unexpected error occurred. Press Enter to exit and check logs.")
        except Exception:
            pass
        # ensure exit non-zero
        sys.exit(1)

if __name__ == "__main__":
    main_wrapper()

