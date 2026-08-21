# Workbook Manager — Architecture

Technical reference for the Workbook Manager. For a high-level overview and usage,
see `README.md`.


## 1. Components & data flow

```
 ┌─────────────────────────┐
 │  Mill workbooks (.xlsx) │   OneDrive-synced on roc-wf-rv
 │  one file per mill      │   path = WORKBOOK_DIR
 └───────────┬─────────────┘
             │  workbook_manager.read_rows()  (copy-first, openpyxl)
             ▼
 ┌─────────────────────────┐        ┌────────────────────────────┐
 │  run_program.py (worker)│──────> │  Historian (ROC-HIST01)    │
 │  parse + reconcile      │ recon  │  _PlcOptTagStatus_AllMills │
 └───────────┬─────────────┘        └────────────────────────────┘
             │  wipe + rewrite one row per tag
             ▼
 ┌─────────────────────────┐
 │  dbo.WorkbookStatus     │  TRD_MSTR (ROC-HIST01)
 │  + LastRun timestamp    │
 └───────────┬─────────────┘
             │  SELECT (fast reads)
             ▼
 ┌─────────────────────────┐
 │  api.py  (FastAPI :8100)│  overview / detail / reconcile / dashboard /
 │                         │  run / send-new / llm-suggest
 └───────────┬─────────────┘
             │  JSON over HTTP (CORS open)
             ▼
 ┌─────────────────────────┐
 │  Web UI (static, :5500) │  workbook.js, dashboard.js, form.js, index.html
 └─────────────────────────┘
```

Two independent processes: **uvicorn** (API, 8100) and a static file server
(**http.server**, 5500). The UI fetches the API cross-origin (CORS is open).


## 2. The workbook parser (`workbook_manager.py`)

Reads a mill `.xlsx` into a list of per-tag dicts.

### Header detection
Column positions are **not** assumed fixed. Each table's header row is detected
and its columns mapped via `_HEADER_MAP` (prefix matching on normalized header
text), so mills with different layouts parse correctly:
- Armour has no Status column → status defaults to `""`.
- Opelika headers read `StandardTagName DataParc` etc. → matched by prefix.

`_HEADER_MAP` maps: status, table, tag_name, plc_tag, plc_name, plc_ip,
data_type, tag_purpose, **default_source, notes, criticality**. (The last three
were being silently dropped before — a fix; Notes in particular is populated on
most rows and is needed by the flagger.)

### First sheet, not active sheet
Uses `wb.worksheets[0]`, **not** `wb.active`. `wb.active` returns whichever sheet
was selected when the file was last saved, which is wrong on multi-sheet
workbooks. This applies to both `read_rows` (here) and `classify_rows`
(`sql_reconcile.py`). The output-builder's `wb_out.active` is left as-is — that's
a freshly created workbook.

### Copy-first read (lock-safe)
`_load_workbook_safe(path)` copies the `.xlsx` to a temp file and parses the
**copy**, then deletes it. A read-copy succeeds even while the original is open
in Excel, so a workbook left open can't lock or hang the parse. Falls back to
opening the original directly if the copy fails.

> **Why this exists:** the hourly scheduled worker was observed hanging
> indefinitely ("Running" forever, freezing the schedule) because it opened a
> workbook that was open in Excel on the same machine. Copy-first eliminates that.

### `count_tags`
Total real data rows = rows where **both** TableName and StandardTagName have
content. Still includes system/metadata rows (RecordID, DateTimeStamp). Used for
the "X of Y complete" figure. Calls `read_rows`, so it inherits copy-first.

### Filename → mill
`FILENAME_TO_MILL` maps exact filenames to mill codes (e.g.
`Opelika PLC Tags.xlsx` → `OPE`). Match is **exact** — a differently-named file
falls back to using the filename as the code, producing a phantom card. (Listed
for fuzzy-matching improvement.)


## 3. Status model & the worker (`run_program.py`)

### Status buckets (workbook-driven)
The four buckets are purely **workbook-status-driven**; the SQL reconcile only
raises *flags*, it never changes a tag's bucket.
- **New** = blank status meeting the new-tag criteria (`is_new_tag`).
- **In Progress** = status "In Progress".
- **Faulty** = status "Faulty".
- **Complete** = status "Complete".

### Reconcile flags (live vs. historian)
`sql_reconcile.reconcile_row` raises flags where the workbook status disagrees
with what the historian shows (e.g. a New tag already writing; a Complete tag
with no data). Flags are advisory — surfaced for a human, never auto-applied.

### The write
`run_program.run()`:
1. `DELETE FROM dbo.WorkbookStatus` (full wipe).
2. Parse every workbook, compute buckets, reconcile.
3. `INSERT` one row per tag, including `LastRun` = `SYSUTCDATETIME()` (UTC) on
   every row.
4. Commit (rollback on error).

Wipe-and-replace keeps the snapshot simple and consistent; every row carries the
same `LastRun` (harmless duplication — the timestamp is one fact about the run).

### `LastRun` / timezone
- Stored **UTC** (`SYSUTCDATETIME()`).
- The API appends `Z` to the ISO string so the browser parses it as UTC and
  converts to its own local time.
- This avoids a server-zone vs. browser-zone mismatch (the naive-timestamp bug
  that pinned the label to "just now" because the browser read it as a future
  time). Historian `DateTimeStamp` is *local* — a separate basis; don't conflate.


## 4. API (`api.py`, FastAPI, :8100)

| Endpoint | Method | Returns |
| --- | --- | --- |
| `/workbooks/overview` | GET | per-mill counts, totals, flag count, `last_run` |
| `/workbooks/{code}/detail` | GET | that mill's tags grouped by table |
| `/workbooks/{code}/reconcile` | GET | live flags vs. historian (suggest-only) |
| `/workbooks/run` | POST | run the worker now (the Refresh button) |
| `/workbooks/{code}/send-new` | POST | open an Outlook draft of new tags |
| `/workbooks/{code}/llm-suggest` | POST | LLM data-quality suggestions |
| `/dashboard` | GET | per-mill requests + coverage for the charts |

- **CORS** is open (`allow_origins=["*"]`) so the 5500 origin can read 8100.
- `MILL_NAMES` maps code → display name (server-side).
- `overview` reads counts from `WorkbookStatus` and `total` from a live
  `count_tags` per workbook.
- `/dashboard` joins three queries: `DataParcMills` (names + connected flag),
  `TagRequestItems` (submitted/complete, `ColumnExists=1 AND HasData=1`), and the
  historian `_PlcOptTagStatus_AllMills` (PLC + optimizer active counts). Sets
  `hist.timeout = 300` on the connection (the OSB cross-linked-server query is
  slow).


## 5. Frontend

- **`form.js`** — tab switching, the tag-request form, header title per tab. The
  mill list is **static `<option>`s in `index.html`**, not fetched. The form
  fetches `/getUnits` and `/getTableNames` from Kevin's relay
  (`roc-laa01:3000`) and submits to `/submit-tags` there. (The form side is
  independent of the Workbook Manager API.)
- **`workbook.js`** — overview, per-mill detail, reconcile, rollout report,
  LLM-suggestions tab. Talks only to the Workbook Manager API (`WB_API`,
  `:8100`).
- **`dashboard.js`** — two clustered SVG bar charts (no chart library). Per-mill
  spacing 130px; axis top = highest + 100 (requests) / + 200 (coverage) via the
  `axisPad` arg.
- **Case sensitivity:** `index.html` references `workbook.js` (lowercase). A
  stray uppercase `Workbook.js` breaks serving (Windows is case-insensitive for
  files but the server/browser are case-sensitive for URLs). Keep one lowercase
  file. (Git is the real long-term fix.)

### Rollout report
- Editable **Comments** input per row (session-only).
- **Copy table** rebuilds the table as inline-styled HTML (email clients ignore
  external CSS): green header, white text, status dots per label, and each
  Comments cell's typed **`input.value`**. `_esc()` escapes typed text.

### "Last updated" ticking label
- `wbLastRunIso` holds the timestamp; a 60s `setInterval` recomputes the label
  from it (pure local math, no server call). One timer at a time (cleared before
  restart); self-cleans if the element is gone. An open page won't catch the
  hourly job's new time until it re-fetches (reload / Refresh).


## 6. LLM suggestions (data-quality checker)

Flagging is **deterministic** (`flagger.py`); suggesting is the **LLM**
(`llm_suggest.py`). Suggest-only — nothing is written.

### `flagger.py` — three flag types
1. **INCOMPLETE** — 3 PLC inputs filled (PLC Tag/Name/IP) but a required field
   blank. Required = table, tag_name, data_type, tag_purpose. **Notes is NOT
   required**; DefaultSource/Criticality never required. BIT is **not** flagged.
2. **MULTI_PATH** — the PLC Tag cell looks like 2+ tags crammed together (split
   on whitespace/newline/comma/slash + repeated-token heuristic). Loose on
   purpose; the LLM confirms.
3. **SHARED_IDENTITY** — tags sharing PLC identity (tag+name+ip, tag+name, or
   tag+ip), **grouped via union-find** so all tags on one path form ONE group,
   reported once.

`flag_rows()` returns `{"row_flags": [...], "shared_groups": [...]}`.

### `llm_suggest.py` — one call per mill, batched & parallel
- Model `claude-sonnet-5`, `MAX_TOKENS=8000`, `BATCH_SIZE=8`. Batches run in
  parallel (ThreadPoolExecutor). `standards.md` is the cached system prompt
  (prompt caching → cheap across a review session).
- **Two prompts / tasks:**
  - **Row suggestions** (INCOMPLETE, MULTI_PATH) → original-vs-suggested rows.
  - **Group comments** (SHARED_IDENTITY) → one comment per group.
- **Output kept small:** Claude does **not** echo the original row (the endpoint
  re-attaches it from our own data by `row_index`); reasons ≤ 2 sentences;
  confidence one word; suggested rows complete (for pasting).
- **Verdict-first reasons:** INCOMPLETE → `"[Column] should be [value]."`;
  MULTI_PATH → `"Expansion needed."` / `"No expansion needed."`; SHARED →
  `"Duplicates look correct."` / `"Duplicates may not be correct."`

### Tuning lessons baked in
- `max_tokens` too high → SDK demands streaming (>10-min estimate). Keep it
  right-sized (8000 for 8-row batches).
- `max_tokens` too low / batch too big → truncated JSON. Smaller batches fix it.
- **Group-index misattribution (fixed):** batches numbered groups locally
  (0,1,2…) but the merge matched against the global list, so comments landed on
  the wrong groups. Fixed with a **stable global `gid`** carried through
  batching. (The infamous "TBPushDistanceSP/fan" comment on the Canter group was
  this — a *misattached* real comment, not a hallucination.)
- Group prompt forbids referencing data not shown and gets each tag's TagPurpose,
  so comments reason from real descriptions.

### `standards.md`
Derived from the official Process Data Collection Standard Procedure plus
patterns mined from real workbooks. Encodes naming rules (TitleCase, Left/Right,
delineation at end), the BIT→INT rule, the raw-tag rule (no aggregates), the
data-type-by-tag-kind conventions, and that **StandardTagName depends on the
table/area**. Edit this file to change LLM behavior — no code change; picked up
next call.



## 7. Send-new-tags (`outlook_sender.py`)

- `POST /workbooks/{code}/send-new` gathers the mill's New tags and opens an
  **Outlook draft** (Classic Outlook, via COM) on roc-wf-rv, in request-column
  format, for review before sending.
- Recipient: **WFM-OT-TagsRUsAgent@westfraser.com**. Subject ends with
  **"Workbook"** (e.g. `MAP Tag Request: N new tags Workbook`) — confirmed not to
  break the agent's subject matching.
- Draft is **displayed, not sent** — human reviews first.


## 8. Scheduling

- Windows Task Scheduler on roc-wf-rv runs `run_program.py` **hourly**.
  - Trigger: Daily, **Repeat every 1 hour, duration Indefinitely**.
  - Action: program `py`, args `-3.12 run_program.py`, **Start in** = project
    folder (so it finds files and loads `.env`; `WORKBOOK_DIR` comes from there).
- **Failure mode seen:** the task hung "Running" (froze the schedule) when a
  workbook was open in Excel. Fixed by copy-first read (§2). Recovery: End the
  run, Run manually with Excel closed, F5 to refresh the (sometimes stale) Status
  column, confirm next-run advances.


## 9. Data & environment

- **Databases (TRD_MSTR on ROC-HIST01):** `WorkbookStatus` (per-tag status +
  `LastRun`), `TagRequestItems` / `TagRequestLog` (agent's request tracking),
  `DataParcMills` (mill dimension), `TagRequestItems_Test` (diagnostic).
- **Historian (ROC-HIST01):** `_PlcOptTagStatus_AllMills` (PLC/optimizer active
  tag counts per site) for the dashboard coverage chart.
- **Env (`.env`):** `TRD_*`, `HIST_*`, `WORKBOOK_DIR`, `CLAUDE_API_KEY`.
- **Machines:** roc-wf-rv (this tool), ROC-HIST01 (SQL), ROC-TAGSRUS (the
  web-hosted form, separate), roc-laa01 (Kevin's relay `:3000` for the form).


## 10. Key files

Backend: `run_program.py` (worker), `api.py` (API), `workbook_manager.py`
(parser + copy-first loader), `sql_reconcile.py` (reconcile/flags),
`flagger.py` (deterministic flags), `llm_suggest.py` (LLM calls),
`outlook_sender.py` (Outlook draft), `standards.md` (LLM system prompt).

Frontend: `index.html`, `form.js`, `workbook.js`, `dashboard.js`, `styles.css`.
