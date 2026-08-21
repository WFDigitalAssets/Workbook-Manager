"""
run_program.py
--------------
The worker behind the "Run program" button and the 5 AM scheduled job.

For every mill workbook it:
  1. Reads each tag row (reusing workbook_manager.read_rows).
  2. Determines a DISPLAY status: New / In Progress / Faulty / Complete,
     upgrading In Progress -> Complete when SQL shows the tag writing
     (reusing sql_reconcile's writer + reconcile_row logic).
  3. Writes one row per tag into dbo.WorkbookStatus (TRD_MSTR), replacing
     the previous run's rows in a single transaction.

It does NOT modify the workbooks — it only computes status for display.

Run manually:   py -3.12 run_program.py
Scheduled:      Task Scheduler calls the same command at 5 AM.

Requires (same .env as the reconcile):
  HIST_SERVER, HIST_DATABASE, HIST_UID, HIST_PWD      (historian, read)
  TRD_SERVER,  TRD_DATABASE,  TRD_UID,  TRD_PWD       (TRD_MSTR, write)
  WORKBOOK_DIR                                         (folder with .xlsx)
"""

import os
import sys
import pyodbc

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Reuse the code we already built and tested.
import workbook_manager as wm
import sql_reconcile as sr


# ----------------------------------------------------------------------
#  Connections
# ----------------------------------------------------------------------

def _conn(server, database, uid, pwd):
    conn_str = (
        "Driver={ODBC Driver 18 for SQL Server};"
        f"Server={server};"
        f"Database={database};"
        f"UID={uid};"
        "PWD={" + pwd + "};"
        "Encrypt=no;"
        "TrustServerCertificate=yes;"
    )
    return pyodbc.connect(conn_str, timeout=15)


def historian_conn():
    return _conn(
        os.getenv("HIST_SERVER", "ROC-HIST01"),
        os.getenv("HIST_DATABASE", "Historian"),
        os.getenv("HIST_UID", ""),
        os.getenv("HIST_PWD", ""),
    )


def trd_conn():
    return _conn(
        os.getenv("TRD_SERVER", "ROC-HIST01"),
        os.getenv("TRD_DATABASE", "TRD_MSTR"),
        os.getenv("TRD_UID", ""),
        os.getenv("TRD_PWD", ""),
    )


# ----------------------------------------------------------------------
#  Status computation
# ----------------------------------------------------------------------

def display_status(row, writing):
    """
    Map a workbook row to a display status string:
    New / In Progress / Faulty / Complete, or "" if the row is none of
    these (e.g. a filled-in tag we don't classify). Status is taken from
    the workbook; SQL is only used to raise flags, never to change the bucket.
    """
    raw = (row["status"] or "").strip().lower()

    # Blank status: it's "New" only if it meets the new-tag criteria.
    if raw == "":
        return "New" if wm.is_new_tag(row) else ""

    # Non-blank: normalize the workbook status to one of the four buckets.
    new_status, _flag = sr.reconcile_row(row, writing)
    ns = (new_status or "").strip().lower()
    if ns == sr.STATUS_IN_PROGRESS:
        return "In Progress"
    if ns == sr.STATUS_FAULTY:
        return "Faulty"
    if ns == sr.STATUS_COMPLETE:
        return "Complete"
    # Any other status (e.g. "Not Relevant") isn't one of the four display
    # buckets — return "" so collect_rows skips it.
    return ""


def _cap(value, n):
    """Trim a value to n chars so an over-long cell can't fail the insert."""
    if value is None:
        return None
    s = str(value)
    return s[:n] if len(s) > n else s


def collect_rows(folder, hist):
    """
    Walk every workbook, compute each tag's display status, and return a
    flat list of tuples ready for insert into dbo.WorkbookStatus.
    """
    writer = sr.make_writer(hist)          # one cached historian query
    out = []

    for path in wm.list_workbook_files(folder):
        mill = wm.mill_code_from_filename(path)
        connected = mill.upper() in sr.CONNECTED_MILLS

        for row in wm.read_rows(path):
            # System/metadata columns are ignored on both sides.
            if row["tag_name"].upper() in sr.SYSTEM_COLUMNS:
                continue

            # Only ask SQL about connected mills; others can't be checked.
            writing = False
            if connected and sr.needs_sql_check(row):
                writing = writer(hist, mill, row["table"], row["tag_name"])

            status = display_status(row, writing)
            if not status:
                continue   # not one of the four display buckets — skip

            # Cap each field to its column width so a single over-long cell
            # can't abort the whole insert (see widen_workbook_status.sql).
            out.append((
                _cap(mill, 10),
                _cap(row["table"], 400),
                _cap(row["tag_name"], 400),
                _cap(row["plc_tag"], 400),
                _cap(row["plc_name"], 400),
                _cap(row["plc_ip"], 200),
                _cap(row["data_type"], 400),
                _cap(row["tag_purpose"], 1000),
                _cap(status, 30),
            ))
    return out


# ----------------------------------------------------------------------
#  Write to SQL (delete previous run + insert this one, in one transaction)
# ----------------------------------------------------------------------

def write_rows(trd, rows):
    cur = trd.cursor()
    try:
        cur.execute("DELETE FROM dbo.WorkbookStatus;")
        cur.executemany(
            """INSERT INTO dbo.WorkbookStatus
            (Mill, TableName, TagName, PlcTag, PlcName, PlcIp,
                DataType, Description, Status, LastRun)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, SYSUTCDATETIME())""",
            rows,
        )
        trd.commit()
    except Exception:
        trd.rollback()
        raise
    finally:
        cur.close()


# ----------------------------------------------------------------------
#  Entry point
# ----------------------------------------------------------------------

def run():
    folder = os.getenv("WORKBOOK_DIR", "Workbooks")
    if not os.path.isdir(folder):
        print(f"WORKBOOK_DIR not found: {folder}", file=sys.stderr)
        sys.exit(1)

    print(f"Reading workbooks from: {folder}")
    hist = historian_conn()
    try:
        rows = collect_rows(folder, hist)
    finally:
        hist.close()

    print(f"Computed {len(rows)} tag rows. Writing to dbo.WorkbookStatus…")
    trd = trd_conn()
    try:
        write_rows(trd, rows)
    finally:
        trd.close()

    # Quick per-mill / per-status summary to stdout (useful in the scheduler log).
    summary = {}
    for r in rows:
        mill, status = r[0], r[8]
        summary.setdefault(mill, {}).setdefault(status, 0)
        summary[mill][status] += 1
    print("Done. Summary:")
    for mill in sorted(summary):
        parts = ", ".join(f"{s}={n}" for s, n in sorted(summary[mill].items()))
        print(f"  {mill}: {parts}")


if __name__ == "__main__":
    run()