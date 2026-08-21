"""
api.py — FastAPI service for the Tag Request Form's Workbook tab.

Serves the JSON contract the frontend (workbook.js) expects:
  GET  /workbooks/overview          -> per-mill counts + totals + flags
  GET  /workbooks/{code}/detail     -> that mill's tags grouped by table
  GET  /workbooks/{code}/reconcile  -> flagged suggestions (SQL reconcile)
  POST /workbooks/run               -> run the worker now ("Run program" button)

Reads status from dbo.WorkbookStatus (filled by run_program.py). The 5 AM
Task Scheduler job and the /run button both call the same worker.

Run locally:
  py -3.12 -m uvicorn api:app --host 0.0.0.0 --port 8100

Frontend: set WB_API in workbook.js to this host:port, and USE_MOCK = false.

Env (.env, same folder):
  TRD_SERVER, TRD_DATABASE, TRD_UID, TRD_PWD    (read WorkbookStatus)
  HIST_SERVER, HIST_DATABASE, HIST_UID, HIST_PWD (reconcile against historian)
  WORKBOOK_DIR                                   (for totals + reconcile + run)
"""

import os
import pyodbc
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import workbook_manager as wm
import sql_reconcile as sr
import run_program


app = FastAPI(title="Workbook Manager API")

# The form is served from a different origin, so allow cross-origin reads.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# Full mill display names (code -> name).
MILL_NAMES = {
    "ANG": "Angelina", "ARM": "Armour", "DGM": "Dudley", "HNM": "Henderson",
    "JOY": "Joyce", "MAP": "Maplesville", "NEW": "New Boston",
    "NWB": "Newberry", "OPE": "Opelika", "RUS": "Russellville",
}


# ----------------------------------------------------------------------
#  DB helpers
# ----------------------------------------------------------------------

def trd_conn():
    return pyodbc.connect(
        "Driver={ODBC Driver 18 for SQL Server};"
        f"Server={os.getenv('TRD_SERVER', 'ROC-HIST01')};"
        f"Database={os.getenv('TRD_DATABASE', 'TRD_MSTR')};"
        f"UID={os.getenv('TRD_UID', '')};"
        "PWD={" + os.getenv("TRD_PWD", "") + "};"
        "Encrypt=no;TrustServerCertificate=yes;",
        timeout=15,
    )


def _mill_totals():
    """
    Total real data rows per mill, read from the workbooks. Used for the
    'X of Y complete' figure. Cheap enough to compute per overview call;
    could be cached later if needed.
    """
    totals = {}
    folder = os.getenv("WORKBOOK_DIR", "Workbooks")
    for path in wm.list_workbook_files(folder):
        mill = wm.mill_code_from_filename(path)
        totals[mill] = wm.count_tags(path)
    return totals


# ----------------------------------------------------------------------
#  Endpoints
# ----------------------------------------------------------------------

@app.get("/workbooks/overview")
def overview():
    """Per-mill counts of each status bucket, total tags, and flag count."""
    conn = trd_conn()
    try:
        cur = conn.cursor()
        cur.execute("""
            SELECT Mill, Status, COUNT(*)
            FROM dbo.WorkbookStatus
            GROUP BY Mill, Status
        """)
        raw = cur.fetchall()

        # Last refresh time (all rows share the same value; read one).
        last_run = None
        try:
            cur.execute("SELECT TOP 1 LastRun FROM dbo.WorkbookStatus")
            row = cur.fetchone()
            if row and row[0]:
                # Stored as UTC (SYSUTCDATETIME); append Z so the browser parses
                # it as UTC and converts to its own local time.
                last_run = row[0].isoformat() + "Z"
        except Exception:
            last_run = None
    finally:
        conn.close()

    totals = _mill_totals()

    # Assemble per-mill counts.
    mills = {}
    def bucket(m):
        return mills.setdefault(m, {"new": 0, "inProgress": 0, "faulty": 0, "complete": 0})

    status_key = {
        "new": "new", "in progress": "inProgress",
        "faulty": "faulty", "complete": "complete",
    }
    for mill, status, n in raw:
        key = status_key.get((status or "").lower())
        if key:
            bucket(mill)[key] += n

    # Build the response, one entry per known mill (so mills with no rows
    # still appear with zeros).
    codes = sorted(set(list(totals) + list(mills) + list(MILL_NAMES)))
    out = []
    for code in codes:
        c = mills.get(code, {"new": 0, "inProgress": 0, "faulty": 0, "complete": 0})
        out.append({
            "code": code,
            "name": MILL_NAMES.get(code, code),
            "counts": c,
            "total": totals.get(code, 0),
            "flags": 0,   # populated on demand via /reconcile; kept 0 here for speed
        })
    return {"mills": out, "last_run": last_run}


@app.get("/workbooks/{code}/detail")
def detail(code: str):
    """That mill's tags grouped by table (from WorkbookStatus)."""
    code = code.upper()
    conn = trd_conn()
    try:
        cur = conn.cursor()
        cur.execute("""
            SELECT TableName, Status, TagName, PlcTag, PlcName, PlcIp,
                   DataType, Description
            FROM dbo.WorkbookStatus
            WHERE Mill = ?
            ORDER BY TableName, TagName
        """, code)
        rows = cur.fetchall()
    finally:
        conn.close()

    tables = {}
    for tname, status, tag, plctag, plcname, plcip, dtype, desc in rows:
        tables.setdefault(tname or "(no table)", []).append({
            "status": status or "",
            "tagName": tag or "",
            "plcTag": plctag or "",
            "plcName": plcname or "",
            "plcIp": plcip or "",
            "dataType": dtype or "",
            "description": desc or "",
        })

    return {
        "code": code,
        "name": MILL_NAMES.get(code, code),
        "tables": [{"name": n, "rows": r} for n, r in tables.items()],
    }


@app.get("/workbooks/{code}/reconcile")
def reconcile(code: str):
    """
    Flagged suggestions for one mill, computed live against the historian.
    Uses the existing reconcile logic (suggest-only; nothing is written).
    """
    code = code.upper()
    folder = os.getenv("WORKBOOK_DIR", "Workbooks")

    # Find the workbook file for this mill.
    path = None
    for p in wm.list_workbook_files(folder):
        if wm.mill_code_from_filename(p) == code:
            path = p
            break
    if not path:
        raise HTTPException(404, f"No workbook found for {code}")

    hist = sr_historian_conn()
    try:
        _column, flags = sr.build_status_column(path, hist, code)
    finally:
        hist.close()

    # build_status_column returns human-readable flag strings; wrap them in the
    # shape the frontend expects. (Kind/confidence are coarse here; the richer
    # judgment cases can be added later.)
    suggestions = []
    for f in flags:
        suggestions.append({
            "row": None, "table": "", "tagName": "",
            "kind": "status",
            "current": "", "suggested": "",
            "reason": f, "confidence": "high",
        })

    return {"code": code, "name": MILL_NAMES.get(code, code), "suggestions": suggestions}


@app.post("/workbooks/run")
def run_now():
    """Trigger the worker immediately (the 'Run program' button)."""
    try:
        run_program.run()
    except SystemExit as e:
        raise HTTPException(500, f"Run failed: {e}")
    except Exception as e:
        raise HTTPException(500, f"Run failed: {e}")
    return {"ok": True}


@app.post("/workbooks/{code}/send-new")
def send_new(code: str):
    """
    Open an Outlook draft (on THIS machine — roc-wfm-rv) containing the mill's
    NEW tags in the request-column format, ready for the user to review and
    send to the agent. The draft is displayed, not sent.

    "New" = blank status that meets the new-tag criteria (same rule the worker
    uses). All such tags are included.
    """
    code = code.upper()
    folder = os.getenv("WORKBOOK_DIR", "Workbooks")

    # Find this mill's workbook.
    path = None
    for p in wm.list_workbook_files(folder):
        if wm.mill_code_from_filename(p).upper() == code:
            path = p
            break
    if not path:
        raise HTTPException(404, f"No workbook found for {code}")

    # Collect the New tags (blank status + meets new-tag criteria).
    new_rows = [r for r in wm.read_rows(path)
                if not (r["status"] or "").strip() and wm.is_new_tag(r)]

    if not new_rows:
        raise HTTPException(400, f"No new tags to send for {code}.")

    import outlook_sender
    try:
        count = outlook_sender.open_draft(code, MILL_NAMES.get(code, code), new_rows)
    except ImportError:
        raise HTTPException(500, "pywin32 not installed (needed for Outlook).")
    except Exception as e:
        # Most common: Classic Outlook not available / COM error.
        raise HTTPException(500, f"Couldn't open Outlook draft: {e}")

    return {"ok": True, "opened": count}


@app.post("/workbooks/{code}/llm-suggest")
def llm_suggest_endpoint(code: str):
    """
    Run ONE Claude call over this mill's flagged rows and return the structured
    suggestions. Triggered by the button in the "LLM suggestions" tab — never
    automatic, since it costs money and takes a few seconds.

    Flagging is deterministic (flagger.py); the LLM only suggests fixes.
    """
    code = code.upper()
    folder = os.getenv("WORKBOOK_DIR", "Workbooks")

    path = None
    for p in wm.list_workbook_files(folder):
        if wm.mill_code_from_filename(p).upper() == code:
            path = p
            break
    if not path:
        raise HTTPException(404, f"No workbook found for {code}")

    import flagger
    import llm_suggest

    rows = wm.read_rows(path)
    flagged = flagger.flag_rows(rows)
    if not flagged["row_flags"] and not flagged["shared_groups"]:
        return {"code": code, "name": MILL_NAMES.get(code, code),
                "counts": flagger.summarize(flagged),
                "suggestions": [], "shared_group_comments": []}

    try:
        result = llm_suggest.suggest_for_mill(flagged, rows)
    except RuntimeError as e:      # missing API key
        raise HTTPException(500, str(e))
    except ValueError as e:        # bad JSON from model
        raise HTTPException(502, str(e))
    except Exception as e:
        raise HTTPException(502, f"LLM call failed: {e}")

    # Row suggestions: re-attach the original row (Claude doesn't echo it) and a
    # copyable block, matched by row_index.
    public_cols = llm_suggest.WORKBOOK_COLUMNS
    for s in result.get("suggestions", []):
        idx = s.get("row_index")
        if isinstance(idx, int) and 0 <= idx < len(rows):
            orig = rows[idx]
            s["original_row"] = {k: (orig.get(k) or "") for k in public_cols}
        else:
            s["original_row"] = {}
        s["copyable"] = llm_suggest.copyable_block(s)

    return {
        "code": code,
        "name": MILL_NAMES.get(code, code),
        "counts": flagger.summarize(flagged),
        "suggestions": result.get("suggestions", []),
        "shared_group_comments": result.get("shared_group_comments", []),
    }


# Historian connection for reconcile (reuse run_program's builder).
def sr_historian_conn():
    return run_program.historian_conn()


# ----------------------------------------------------------------------
#  DASHBOARD — replaces the Power BI iframe. Pulls the same SQL the
#  dashboard used: tag requests (TRD_MSTR) + tag coverage (Historian),
#  joined to the DataParcMills dimension. No date filter.
# ----------------------------------------------------------------------

@app.get("/dashboard")
def dashboard():
    """
    One JSON array, one entry per DataParc mill:
      millCode, millName, connected,
      submitted, inProgress, complete,     (tag requests)
      plcTags, optimizerTags               (tags actively recording)
    """
    # --- Query C: mill dimension (names + connected flag) ---
    mills = {}
    trd = trd_conn()
    try:
        cur = trd.cursor()
        cur.execute("""
            SELECT MillCode, MillName, HistorianConnected
            FROM dbo.DataParcMills
        """)
        for code, name, connected in cur.fetchall():
            mills[(code or "").upper()] = {
                "millCode": code,
                "millName": name or code,
                "connected": bool(connected),
                "submitted": 0, "inProgress": 0, "complete": 0,
                "plcTags": 0, "optimizerTags": 0,
            }

        # --- Query A: tag requests per mill (unfiltered) ---
        cur.execute("""
            SELECT
                i.MillCode,
                COUNT(*) AS Submitted,
                SUM(CASE WHEN i.ColumnExists=1 AND i.HasData=1 THEN 1 ELSE 0 END) AS Complete
            FROM dbo.TagRequestItems i
            GROUP BY i.MillCode
        """)
        for code, submitted, complete in cur.fetchall():
            key = (code or "").upper()
            if key in mills:
                submitted = int(submitted or 0)
                complete = int(complete or 0)
                mills[key]["submitted"] = submitted
                mills[key]["complete"] = complete
                mills[key]["inProgress"] = submitted - complete
    finally:
        trd.close()

    # --- Query B: tag coverage (PLC + optimizer recording) ---
    hist = run_program.historian_conn()
    hist.timeout = 300
    try:
        cur = hist.cursor()
        cur.execute("""
            SELECT
                Site AS MillCode,
                SUM(CASE WHEN [Schema]='dbo' AND Status='active' THEN 1 ELSE 0 END) AS PlcTags,
                SUM(CASE WHEN [Schema]='opt' AND Status='active' THEN 1 ELSE 0 END) AS OptimizerTags
            FROM dbo._PlcOptTagStatus_AllMills
            GROUP BY Site
        """)
        for code, plc, opt in cur.fetchall():
            key = (code or "").upper()
            if key in mills:   # drop sites not in DataParcMills (out of scope)
                mills[key]["plcTags"] = int(plc or 0)
                mills[key]["optimizerTags"] = int(opt or 0)
    finally:
        hist.close()

    return {"mills": list(mills.values())}