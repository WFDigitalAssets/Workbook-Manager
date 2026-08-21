"""
flagger.py
----------
Deterministic detection of rows that need LLM review. NO LLM here — this is
plain code that finds candidates; the LLM only SUGGESTS fixes for what this
flags.

Three flag types (per the agreed design):

  1. INCOMPLETE — the 3 PLC inputs (PLC Tag, PLC Name, PLC IP) are filled, but
     one or more of {TableName, StandardTagName, DataType, TagPurpose, Notes}
     is blank. DefaultSource / Criticality being blank does NOT count.

  2. MULTI_PATH — the PLC Tag cell appears to contain 2+ PLC tags crammed
     together. Detected LOOSELY here (split on delimiters incl. whitespace,
     newline, comma, slash, semicolon; plus a repeated-structure heuristic).
     The LLM confirms and expands — so false positives here are acceptable;
     the LLM is the second gate.

  3. SHARED_IDENTITY — two or more rows share PLC identity:
       (PLC Tag + PLC Name + PLC IP) all equal, OR
       (PLC Tag + PLC Name) equal, OR
       (PLC Tag + PLC IP) equal.

BIT data type is intentionally NOT a flag.

Each flagged row is returned with its flag type(s) and the context the LLM
needs (the full row, plus any partner rows for SHARED_IDENTITY).
"""

import re

# Columns that must be present for an otherwise-complete tag. Blank here (with
# all 3 PLC inputs filled) => INCOMPLETE.
# Columns that must be present for an otherwise-complete tag. Blank here (with
# all 3 PLC inputs filled) => INCOMPLETE. Notes is NOT required (often blank on
# valid rows); DefaultSource / Criticality are also not required.
REQUIRED_WHEN_PLC_FILLED = ["table", "tag_name", "data_type", "tag_purpose"]

# System/metadata columns are never user tags — skip entirely.
SYSTEM_COLUMNS = {
    "DATETIMESTAMP", "RECORDID", "DATETIMEWRITTEN",
    "TMSTAMP", "DATETIMESTAMPWRITTEN",
}


def _blank(v):
    return v is None or str(v).strip() == ""


def _plc_filled(v):
    """A PLC input counts as filled if non-empty and not the N/A placeholder."""
    return not _blank(v) and str(v).strip().upper() != "N/A"


def _three_plc_inputs_filled(row):
    return (_plc_filled(row.get("plc_tag"))
            and _plc_filled(row.get("plc_name"))
            and _plc_filled(row.get("plc_ip")))


# ----------------------------------------------------------------------
#  Flag 1 — incomplete row
# ----------------------------------------------------------------------
def _incomplete_missing(row):
    """
    If the 3 PLC inputs are filled, return the list of required columns that are
    blank. Empty list => not incomplete. (Notes may not exist as a key in every
    parser output; treat missing key as blank.)
    """
    if not _three_plc_inputs_filled(row):
        return []
    return [c for c in REQUIRED_WHEN_PLC_FILLED if _blank(row.get(c))]


# ----------------------------------------------------------------------
#  Flag 2 — multiple PLC paths in one cell (loose detection)
# ----------------------------------------------------------------------
# Delimiters people might use between crammed tags.
_SPLIT_RE = re.compile(r"[/,;\n\r\t]| {2,}")

def _looks_like_multi_path(plc_tag):
    """
    Loose heuristic: does this one cell look like it holds 2+ PLC tags?
    We deliberately over-detect; the LLM confirms.

    Signals:
      - splits into 2+ non-trivial chunks on common delimiters, OR
      - contains a repeated near-identical structure (e.g. TA_x / TB_x) even
        without a clear delimiter.
    """
    if _blank(plc_tag):
        return False
    s = str(plc_tag).strip()

    # 1) delimiter split
    parts = [p.strip() for p in _SPLIT_RE.split(s) if p.strip()]
    if len(parts) >= 2:
        # require at least two parts that look "tag-like" (contain a letter and
        # are reasonably long) to avoid splitting a single path with a stray comma
        taglike = [p for p in parts if len(p) >= 4 and re.search(r"[A-Za-z]", p)]
        if len(taglike) >= 2:
            return True

    # 2) repeated-structure heuristic (no delimiter case): look for two runs that
    # share a long common substring but differ by a small token (TA vs TB, K1T1
    # vs K1T2). Cheap approximation: if a token of length>=4 appears 2+ times.
    tokens = re.findall(r"[A-Za-z][A-Za-z0-9]{3,}", s)
    seen = {}
    for t in tokens:
        seen[t] = seen.get(t, 0) + 1
        if seen[t] >= 2:
            return True
    return False


# ----------------------------------------------------------------------
#  Flag 3 — shared PLC identity across rows
# ----------------------------------------------------------------------
def _identity_keys(row):
    """The three collision keys for a row (only if the needed fields are filled)."""
    tag = str(row.get("plc_tag") or "").strip().upper()
    name = str(row.get("plc_name") or "").strip().upper()
    ip = str(row.get("plc_ip") or "").strip().upper()
    keys = []
    if tag and name and ip:
        keys.append(("tag+name+ip", (tag, name, ip)))
    if tag and name:
        keys.append(("tag+name", (tag, name)))
    if tag and ip:
        keys.append(("tag+ip", (tag, ip)))
    return keys


def find_shared_identity_groups(rows):
    """
    Return a list of GROUPS, each a sorted list of row indexes that share PLC
    identity (on tag+name+ip, tag+name, or tag+ip). Groups are merged so that if
    row A shares with B and B shares with C, all three are one group. Each group
    has 2+ rows. Every group is reported ONCE (no per-row duplication).
    """
    # Build buckets by each identity key.
    buckets = {}
    for i, r in enumerate(rows):
        if not _plc_filled(r.get("plc_tag")):
            continue
        for _label, kv in _identity_keys(r):
            buckets.setdefault(kv, []).append(i)

    # Union-find to merge overlapping buckets into connected groups.
    parent = {}
    def find(x):
        parent.setdefault(x, x)
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x
    def union(a, b):
        parent[find(a)] = find(b)

    for kv, idxs in buckets.items():
        if len(idxs) < 2:
            continue
        first = idxs[0]
        for j in idxs[1:]:
            union(first, j)

    # Collect members by root.
    groups = {}
    for kv, idxs in buckets.items():
        if len(idxs) < 2:
            continue
        for i in idxs:
            groups.setdefault(find(i), set()).add(i)

    return [sorted(members) for members in groups.values() if len(members) >= 2]


# ----------------------------------------------------------------------
#  Top-level: flag a list of rows
# ----------------------------------------------------------------------
def flag_rows(rows):
    """
    Given parsed workbook rows, return:
      {
        "row_flags": [   # per-row: INCOMPLETE and MULTI_PATH only
           {"row_index": int, "row": <dict>,
            "flags": [{"type":"INCOMPLETE","missing":[...]}, {"type":"MULTI_PATH"}]},
           ...
        ],
        "shared_groups": [   # one entry per PLC-path group (2+ rows)
           {"row_indexes": [i, j, ...]},
           ...
        ],
      }
    Shared-identity is grouped (each group reported ONCE, covering all its tags),
    not repeated per row. System/metadata columns are skipped.
    """
    row_flags = []
    for i, r in enumerate(rows):
        if str(r.get("tag_name") or "").strip().upper() in SYSTEM_COLUMNS:
            continue
        flags = []
        missing = _incomplete_missing(r)
        if missing:
            flags.append({"type": "INCOMPLETE", "missing": missing})
        if _looks_like_multi_path(r.get("plc_tag")):
            flags.append({"type": "MULTI_PATH"})
        if flags:
            row_flags.append({"row_index": i, "row": r, "flags": flags})

    groups = find_shared_identity_groups(rows)
    # Skip any group whose rows are all system columns (rare, but safe).
    shared_groups = []
    for g in groups:
        real = [i for i in g
                if str(rows[i].get("tag_name") or "").strip().upper() not in SYSTEM_COLUMNS]
        if len(real) >= 2:
            shared_groups.append({"row_indexes": real})

    return {"row_flags": row_flags, "shared_groups": shared_groups}


def summarize(flagged):
    """Counts by flag type. flagged is the dict returned by flag_rows."""
    counts = {"INCOMPLETE": 0, "MULTI_PATH": 0, "SHARED_IDENTITY_GROUPS": 0}
    for rec in flagged.get("row_flags", []):
        for f in rec["flags"]:
            counts[f["type"]] += 1
    counts["SHARED_IDENTITY_GROUPS"] = len(flagged.get("shared_groups", []))
    return counts