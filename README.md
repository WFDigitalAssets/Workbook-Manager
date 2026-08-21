# Workbook Manager

A tool for tracking the rollout status of PLC tags across West Fraser mills. It
reads the per-mill tag workbooks, reconciles them against the SQL historian,
and presents everything through a web interface: a dashboard, per-mill status,
a rollout report, an AI-assisted data-quality checker, and a one-click way to
send new tag requests to the TagsRUs agent.

It runs locally on **roc-wf-rv** and is the PowerBI+Workbook companion to the
web-hosted Tag Request form (which runs separately on ROC-TAGSRUS).


## What it does

- **Dashboard** — two charts across all mills: tag-request status (submitted /
  in progress / complete) and tag coverage (PLC + optimizer tags actively
  recording). Replaces the old Power BI embed.
- **Per-mill status** — for each mill, tags grouped by table and filtered by
  status (New / In Progress / Faulty / Complete), plus the embedded read-only
  Excel workbook.
- **Rollout report** — a summary table of every mill's counts with editable
  comments, and a "Copy table" button that produces email-ready styled HTML.
- **Review flags** — live reconcile of a mill's workbook against the historian,
  surfacing mismatches (e.g. a tag marked New that is already writing data).
- **LLM suggestions** — an AI data-quality checker that flags problem tags
  (incomplete rows, multiple PLC paths crammed in one cell, tags sharing a PLC
  path) and suggests fixes. Click-to-run; suggest-only.
- **Send new tags** — opens an Outlook draft with a mill's new tags formatted for
  the TagsRUs agent, for you to review and send.


## How it works (high level)

```
  Mill workbooks (.xlsx, OneDrive-synced)
            │  parsed by the worker
            ▼
   dbo.WorkbookStatus  (SQL, TRD_MSTR)  ◄─── reconciled against the historian
            │  served as JSON
            ▼
     FastAPI  (api.py, :8100)
            │
            ▼
     Web UI  (workbook.js + form.js + dashboard.js)
```

- A **worker** (`run_program.py`) reads every mill workbook, computes each tag's
  status bucket, reconciles against the historian, and writes one row per tag to
  **`dbo.WorkbookStatus`**. This runs on a schedule and on demand.
- A **FastAPI** service (`api.py`, port 8100) serves that data as JSON, plus the
  dashboard, reconcile, LLM-suggestion, and send-new endpoints.
- The **web UI** reads those endpoints. The page loads instantly from the cached
  SQL snapshot; the expensive parse/reconcile only happens during a refresh.

**Why SQL is in the middle** (not reading workbooks directly on every load): fast
reads, the historian reconcile can't come from the workbook, decoupling from
Excel file locks, safe concurrent reads, and — importantly — the UI can run
anywhere that can reach SQL, not only the machine that has the synced files.

## Running it

Two processes, both from the project folder on roc-wf-rv:

```powershell
# 1. Backend API
py -3.12 -m uvicorn api:app --host 0.0.0.0 --port 8100

# 2. Frontend (static files)
py -3.12 -m http.server 5500
```

Then open the form's page and go to the **Workbook** tab. If the page says
"couldn't load workbooks / failed to fetch," the backend (uvicorn on 8100) isn't
running — start it.

Configuration lives in `.env` (same folder):

- `TRD_SERVER`, `TRD_DATABASE`, `TRD_UID`, `TRD_PWD` — the status DB (TRD_MSTR).
- `HIST_SERVER`, `HIST_DATABASE`, `HIST_UID`, `HIST_PWD` — the historian
  (reconcile + dashboard coverage).
- `WORKBOOK_DIR` — folder of the OneDrive-synced `.xlsx` workbooks.
- `CLAUDE_API_KEY` — for the LLM suggestions feature.

## Keeping status current

- **Scheduled refresh** — a Windows Task Scheduler job runs `run_program.py`
  **every hour** on roc-wf-rv, so the cached status stays fresh without anyone
  clicking. (Task config: trigger "repeat every 1 hour, indefinitely"; action
  `py -3.12 run_program.py`; **Start in** = the project folder so it finds the
  files and `.env`.)
- **Last updated label** — the rollout report shows "Updated X minutes ago,"
  read from the `LastRun` timestamp stored on each refresh (stored UTC, shown in
  the browser's local time). The label ticks up on its own every minute without
  hitting the server.
- **Manual "Refresh numbers"** — runs the worker on demand and reloads. Use this
  before emailing a report if you just edited a workbook, since the hourly job
  won't have picked up your change yet.

Data changes slowly (mills edit workbooks ~1–3×/day), so hourly is plenty; the
manual button covers the "I need it right now" case.


## Common tasks

- **Update a mill's status** → edit the workbook, then hit **Refresh numbers** to
  pull it into the site.
- **Send a report** → Workbook tab → rollout report → type any comments →
  **Copy table** → paste into an email.
- **Check a mill for issues** → open the mill → **Review flags** (live) or the
  **LLM suggestions** tab (click "Get suggestions").
- **Submit new tags** → open the mill → **New** filter → **Send N new tags to
  agent** (opens an Outlook draft to review before sending).


## Handoffs / room for improvement

### Infrastructure & deployment

- **Daily emailed rollout report** — a scheduled Mon–Fri 11 AM email of the
  report to Valentin and Ron is planned but not built. Decide **SMTP vs Outlook
  COM** first: SMTP works headless (recommended, like the agent's escalation
  emails); Outlook COM needs an interactive login, so a scheduled send may not
  fire unless someone is logged in.

- **Move to ROC-TAGSRUS (hosted).** LLM (`api.anthropic.com:443`) and SQL
  (`ROC-HIST01:1433`) outbound are already open there. The unsolved piece is how
  the workbook `.xlsx` files. Microsoft 365 is already in roc-tagsrus so you should theoretically be able to derive the onedrive sync to the roc-tagsrus machine. Additionally, you will need to move the PLC Tag folder to a centralized area (i.e. The TagsRUs folder) for security purposes.

- **Change credentials to company LLM API account** - Current API Key is using my own personal account

- **Access control for a hosted Workbook tab.** Hiding the tab is not security
  (the API is reachable by URL). Real options: backend token auth (moderate) or
  Entra/SSO tied to a WF group (robust, IT-owned). Kevin/Valentin/IT decision.

- **Run in the cloud (optional).** Consider running the API wrapper / hosting off
  a personal business machine (e.g. a cloud host) so it doesn't depend on
  roc-wf-rv being up and logged in.

### Planned LLM / feature enhancements

- **Cross-compare new tags against the existing workbook (high priority).** Send
  a mill's new tags to Claude and have it cross-reference against what's already
  in that workbook — catch duplicates and near-duplicates before they're
  requested. (Called out repeatedly in the working notes as the next big LLM
  improvement.)

- **Suggest tag-name convention fixes.** The LLM currently handles INCOMPLETE,
  MULTI_PATH, and shared-path groups, but not dedicated name-convention fixes
  (TitleCase, Left/Right, delineation-at-end, no special chars). `standards.md`
  already encodes the rules — this would be a focused name-linting suggestion
  type.

- **Suggest creating new tables.** Have the LLM flag when a requested tag doesn't
  fit any existing table and a new table is warranted.

- **Workbook-level SQL "already exists" check.** Before requesting, check the
  historian to see whether a tag is already writing — surface existing matches at
  the workbook level (the agent does this at request time; doing it earlier has
  real value). Likely the highest-value non-duplicative addition.

- **LLM suggestions need iteration** on real usage — prompt tuning, batch size,
  and confidence calibration as more mills are reviewed.

- **Reduce LLM / OneDrive latency further** — already improved (Sonnet + parallel
  batches), but ongoing.

### Status automation

- **Drive status from `TagRequestLog` / new completion logic.** Auto-update a
  tag's workbook status from the agent's tracking tables. **Caution:** the agent's
  current "has data" is history-based and unreliable (see the CanterProducing
  finding) — auto-*setting* status would enshrine that flaw. Prefer
  auto-*suggesting* (surface for review) over auto-setting, and settle the
  recency-window definition first (see the agent handoffs / `check_tag_completion`
  work).

- **Overdue / error notifications.** An "overdue" notice for tags stuck
  in-progress, and an error-message follow-up if the agent fails a request (the
  agent's `send_reply` is the natural place — a failed run could email a short
  "we hit an error, we're looking at it" notice to restore user-facing feedback).

- **"Requests not seen as unread" in the shared mailbox** — investigate/fix so
  the agent reliably picks up new requests.

### Known bugs / cleanup

- **`# DASHBOARD LAYER` parsed as a table.** The parser treats the
  `# DASHBOARD LAYER` separator row as a table name. Should be skipped.

- **`dryerdatatable` question** — an open "why dryerdatatable??" note; verify
  energy/dryer/kiln table naming is consistent across mills (some were merged
  into `KilnDataTable`).

- **Data inputs not showing in some views / "what is review?"** — a couple of
  UI/data-display inconsistencies noted during testing; audit and resolve.

- **Filename→mill matching is brittle** (exact match). Switch to keyword/fuzzy
  matching so small renames don't create phantom cards.

- **First-table header detection.** Some workbooks have a non-standard first
  header the parser can miss (caused an ANG undercount once, fixed by
  standardizing that header). A more robust first-header catch would remove the
  manual step.

- **Excel formula-apply error** when someone inserts/deletes workbook rows —
  formulas shift/break; needs handling on the workbook side.

- **Remove non-relevant tags** (e.g. fuel tags for mills that don't have them)
  and reconcile duplicate-equipment rows.

### Possible workbook schema additions

- **Min/Max columns** and **Units of Measure (UoM) columns** — considered for the
  workbooks; not yet added. (Units can often be derived from TagPurpose/Notes.)

- **Comments persistence.** Rollout-report comments are session-only (cleared on
  reload). If they need to survive, store them (SQL or local file).

### Other

- **Auto-refresh on view (optional).** Current design is scheduled + manual. A
  "serve cached instantly, refresh in the background, update when done" approach
  is nicer UX but more complex and probably overkill for slow-changing data —
  documented as an option, not recommended for now.

- **PowerBI week-slicer issue.** The week slicer didn't work for certain users
  (e.g. Valentin) while year/month did — narrowed to the `RequestWeek`
  field/slicer specifically; still unresolved.

- **Scalability.** General goal to keep the parser/flow easy to extend as mills
  and tables grow (header detection is already layout-tolerant; keep it that way).

See `ARCHITECTURE.md` for the technical detail behind all of this.
