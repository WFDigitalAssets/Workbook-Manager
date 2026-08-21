"""
llm_suggest.py
--------------
One Claude call per mill. Takes the deterministically-flagged rows (from
flagger.py) and asks Claude to SUGGEST fixes, guided by standards.md.

Design:
  - standards.md is loaded into the SYSTEM prompt, marked cacheable so repeated
    calls in a review session are cheap.
  - The USER message carries only the flagged rows (+ partner rows for
    SHARED_IDENTITY) as structured JSON — Claude never sees raw HTML/Excel.
  - Claude returns STRICT JSON: for each flagged row, the original row and one
    or more suggested rows, plus a confidence level and a short reason.
  - Suggest-only. Nothing is written anywhere. A human reviews and applies.

Requires ANTHROPIC/CLAUDE API key in env as CLAUDE_API_KEY.
"""

import os
import json

# Workbook column order for copyable output (matches the request/workbook layout).
WORKBOOK_COLUMNS = [
    "status", "table", "tag_name", "plc_tag", "plc_name", "plc_ip",
    "data_type", "tag_purpose", "default_source", "notes", "criticality",
]

# Human-friendly headers for those keys, in the same order (for the copyable
# tab-separated block the user pastes into Excel).
COLUMN_HEADERS = [
    "Status", "TableName", "StandardTagName", "PLC Tag", "PLC Name",
    "PLC IP Address", "DataType", "TagPurpose", "DefaultSource", "Notes",
    "Criticality",
]

MODEL = "claude-sonnet-5"      # fast + capable for this structured task
MAX_TOKENS = 8000              # ample for an 8-row batch; under the streaming threshold
BATCH_SIZE = 8                 # smaller batches = smaller responses; run in parallel


def _load_standards(path=None):
    path = path or os.path.join(os.path.dirname(__file__), "standards.md")
    with open(path, encoding="utf-8") as f:
        return f.read()


def _row_public(row):
    """Trim a parsed row to just the workbook fields Claude should see."""
    return {k: (row.get(k) or "") for k in WORKBOOK_COLUMNS}


def build_row_payload(row_flags):
    """Payload for per-row flags (INCOMPLETE, MULTI_PATH)."""
    items = []
    for rec in row_flags:
        items.append({
            "row_index": rec["row_index"],
            "row": _row_public(rec["row"]),
            "flags": [
                ({"type": "INCOMPLETE", "missing": f["missing"]}
                 if f["type"] == "INCOMPLETE" else {"type": "MULTI_PATH"})
                for f in rec["flags"]
            ],
        })
    return items


def build_group_payload(shared_groups, all_rows):
    """Payload for shared-identity groups: each group lists all tags sharing a path.
    Uses each group's OWN 'gid' (stable global id) so comments match back correctly
    even when groups are processed in separate batches."""
    groups = []
    for g in shared_groups:
        gid = g["gid"]
        tags = [dict(_row_public(all_rows[i]), row_index=i) for i in g["row_indexes"]]
        groups.append({"group_index": gid, "tags": tags})
    return groups


ROW_SYSTEM = """You are a data-quality assistant for West Fraser's PLC tag workbooks. \
You review tag rows FLAGGED by deterministic code and SUGGEST corrections. You never \
invent data you cannot reasonably infer, and you never change a tag's Status.

You will receive the West Fraser PLC Tag Standards, then a JSON list of flagged rows. \
For EACH flagged row, produce a suggestion by flag type:

- INCOMPLETE: the 3 PLC inputs are filled but one or more of TableName, StandardTagName, \
DataType, TagPurpose is blank. Suggest the correct value(s) for ONLY the blank field(s), \
using the standards and the row's context. Do not alter fields already filled.

- MULTI_PATH: the PLC Tag cell may contain 2+ PLC tags crammed together. First DECIDE if \
it really is multiple tags. If yes, EXPAND into one complete suggested row per distinct \
PLC tag, each correct per the standards (StandardTagName depends on the table/area). If \
it is actually a single tag, say so and make no expansion.

CONFIDENCE: label each with ONE word — "high", "medium", or "low". MULTI_PATH expansion \
is usually mechanical (often high); INCOMPLETE blank-fills are guesses (usually medium/low). \
Lower confidence when unsure; never present a guess as certain.

REASON FORMAT — every reason must START with a short direct verdict, then the explanation:
- INCOMPLETE: start with "<Column> should be <value>." (e.g. "DataType should be int.") \
If multiple fields are blank, lead with the most important one. Then a sentence of why.
- MULTI_PATH: start with "Expansion needed." or "No expansion needed." Then a sentence of why.
Keep the whole reason to AT MOST two sentences including the verdict.

OUTPUT: STRICT JSON ONLY — no prose, markdown, or code fences. Keep it SMALL: do NOT echo \
the original row back. Schema:
{
  "suggestions": [
    {
      "row_index": <int matching the input>,
      "flag_types": ["INCOMPLETE" | "MULTI_PATH", ...],
      "confidence": "high" | "medium" | "low",
      "reason": "<at most two sentences>",
      "suggested_rows": [ { <one complete suggested row> }, ... ]
    }
  ]
}
Each suggested row uses these keys: status, table, tag_name, plc_tag, plc_name, plc_ip, \
data_type, tag_purpose, default_source, notes, criticality. Return the COMPLETE row so it \
can be pasted into the workbook. ALWAYS copy the original Status into every suggested row \
unchanged."""


GROUP_SYSTEM = """You are a data-quality assistant for West Fraser's PLC tag workbooks. \
You are given GROUPS of tags that each share the same PLC identity (same PLC Tag, and/or \
PLC Name, and/or PLC IP). For each group, judge whether it is sensible for ALL these tags \
to originate from that one PLC path — i.e. whether one physical source legitimately feeds \
these distinct logical tags, or whether it looks like a copy/paste error where some tags \
need their own different PLC paths.

Consider the tag names, their purposes, and their meanings: a single source feeding tags \
whose meanings CONFLICT (e.g. one bit feeding both "Running" and "Downtime") is suspicious; \
a single source feeding closely related values may be fine.

IMPORTANT: Base your comment ONLY on the tag names, purposes, and data types actually shown \
in the group. Do NOT reference any signal, setpoint, or equipment that is not present in the \
provided data. If you are unsure what a tag means, say so rather than inventing a meaning.

COMMENT FORMAT — start with a short direct verdict, then the explanation:
- Start with "Duplicates look correct." or "Duplicates may not be correct."
- Then ONE sentence of why (name/purpose conflict, data-type mismatch, etc.).
Keep the whole comment to AT MOST two sentences including the verdict.

OUTPUT: STRICT JSON ONLY — no prose, markdown, or code fences. For each group give ONE \
comment and a one-word confidence. Do NOT rewrite the rows. Schema:
{
  "group_comments": [
    {
      "group_index": <int matching the input>,
      "confidence": "high" | "medium" | "low",
      "comment": "<verdict first, then 1 sentence why — 2 sentences max>"
    }
  ]
}"""


def _suggest_row_batch(row_flags, client, standards):
    """One Claude call for a batch of per-row flags. Returns list of suggestions."""
    payload = build_row_payload(row_flags)
    resp = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=[
            {"type": "text", "text": ROW_SYSTEM},
            {"type": "text", "text": standards, "cache_control": {"type": "ephemeral"}},
        ],
        messages=[{"role": "user",
                   "content": "Flagged rows as JSON. Return suggestions per the schema.\n\n"
                              + json.dumps(payload, indent=2)}],
    )
    return _parse_json(resp).get("suggestions", [])


def _comment_group_batch(shared_groups, all_rows, client, standards):
    """One Claude call for a batch of shared-identity GROUPS. Returns group_comments."""
    payload = build_group_payload(shared_groups, all_rows)
    resp = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=[
            {"type": "text", "text": GROUP_SYSTEM},
            {"type": "text", "text": standards, "cache_control": {"type": "ephemeral"}},
        ],
        messages=[{"role": "user",
                   "content": "Shared-identity groups as JSON. One comment per group.\n\n"
                              + json.dumps(payload, indent=2)}],
    )
    return _parse_json(resp).get("group_comments", [])


def _parse_json(resp):
    text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text").strip()
    if getattr(resp, "stop_reason", None) == "max_tokens":
        raise ValueError("A batch was cut off (hit max_tokens). Try a smaller BATCH_SIZE.")
    if text.startswith("```"):
        text = text.strip("`")
        if text.lstrip().lower().startswith("json"):
            text = text.lstrip()[4:]
    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        raise ValueError(f"LLM did not return valid JSON: {e}\n---\n{text[:500]}")


def suggest_for_mill(flagged, all_rows, api_key=None, standards_path=None):
    """
    flagged is the dict from flagger.flag_rows: {"row_flags": [...], "shared_groups": [...]}.
    Returns:
      {
        "suggestions": [ per-row INCOMPLETE / MULTI_PATH suggestions ],
        "shared_group_comments": [
           {"group_index", "confidence", "comment", "tags": [ {row fields...}, ... ]},
        ],
      }
    Row batches and group batches all run in parallel; standards are cached.
    """
    import anthropic

    api_key = api_key or os.getenv("CLAUDE_API_KEY") or os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("CLAUDE_API_KEY not set in environment.")

    row_flags = flagged.get("row_flags", [])
    shared_groups = flagged.get("shared_groups", [])
    if not row_flags and not shared_groups:
        return {"suggestions": [], "shared_group_comments": []}

    # Give each group a stable global id so batch-relative indexing can't
    # misattribute comments to the wrong group.
    for gid, g in enumerate(shared_groups):
        g["gid"] = gid

    standards = _load_standards(standards_path)
    client = anthropic.Anthropic(api_key=api_key)

    row_batches = [row_flags[i:i + BATCH_SIZE]
                   for i in range(0, len(row_flags), BATCH_SIZE)]
    group_batches = [shared_groups[i:i + BATCH_SIZE]
                     for i in range(0, len(shared_groups), BATCH_SIZE)]

    from concurrent.futures import ThreadPoolExecutor
    suggestions, group_comments = [], []

    tasks = []
    for b in row_batches:
        tasks.append(("row", b))
    for b in group_batches:
        tasks.append(("group", b))

    def run(task):
        kind, batch = task
        if kind == "row":
            return ("row", _suggest_row_batch(batch, client, standards))
        return ("group", _comment_group_batch(batch, all_rows, client, standards))

    if tasks:
        with ThreadPoolExecutor(max_workers=min(8, len(tasks))) as pool:
            for kind, result in pool.map(run, tasks):
                (suggestions if kind == "row" else group_comments).extend(result)

    # Attach each group's tag rows to its comment, matched by the STABLE gid
    # (the model echoes group_index = gid back).
    by_gid = {c.get("group_index"): c for c in group_comments}
    enriched = []
    for g in shared_groups:
        gid = g["gid"]
        c = dict(by_gid.get(gid, {"group_index": gid, "confidence": "low", "comment": ""}))
        c["tags"] = [dict(_row_public(all_rows[i]), row_index=i)
                     for i in g["row_indexes"]]
        enriched.append(c)

    return {"suggestions": suggestions, "shared_group_comments": enriched}


def to_copyable(suggested_row):
    """
    Turn one suggested row dict into a single tab-separated line in workbook
    column order, ready to paste into Excel. Status preserved as-is.
    """
    return "\t".join(str(suggested_row.get(k, "") or "") for k in WORKBOOK_COLUMNS)


def copyable_block(suggestion):
    """
    A full copyable block for one suggestion: header line + one line per suggested
    row, tab-separated. (Header can be dropped by the frontend if pasting into an
    existing sheet.)
    """
    lines = ["\t".join(COLUMN_HEADERS)]
    for sr in suggestion.get("suggested_rows", []):
        lines.append(to_copyable(sr))
    return "\n".join(lines)