# Workbook Manager

A tool for tracking the rollout status of PLC tags across West Fraser mills. It
reads the per-mill tag workbooks, reconciles them against the SQL historian,
and presents everything through a web interface: a dashboard, per-mill status,
a rollout report, an AI-assisted data-quality checker, and a one-click way to
send new tag requests to the TagsRUs agent.

It runs locally on **roc-wf-rv** and is the PowerBI+Workbook companion to the
web-hosted Tag Request form (which runs separately on ROC-TAGSRUS).

## TLDR
**Read these things at least**

How to use the workbook manager:
- The workbook-manager is locally web-hosted in roc-wfm-rv
- The vm pw is `helpme` and the pin is `1111110`
- 1.To use the workbook manager, ensure that you have an updated version of the statuses
- 2.Check the new tag inputs by putting them through the LLM call
- The LLM call will not work and will need to be wired to a new api key
- 3.Click `send new tags` in the new tag tab
- 4.Change the status of the tags in the workbook to `in progress` once forwarded to the ROC
- 5.Change the status of those tags to `complete` once ROC has completed the request or `faulty` if it requires fixing

**The workbook manager is still in testing phase and may require additional tweaking for efficient and smooth usage**



Two scheduled tasks:
- **Hourly workbook refresh** - Re-analyzes the workbooks once every hour and re-writes to the statuses to `WorkbookStatus` in `TRD_MSTR`. 
- **Daily send out report** - Uses the smtp send relay from `tagsrus-agent@automationewp.com` to send out a report every weekday at 12PM.

The handoffs for the workbook management are as follows:
- **Better tag completion detection.** The current `check_tag_completion` writes to the `TagRequestLog` and `TagRequestItems` and determines if a tag is complete by seeing if it has on non-null value in it's history. I've made a better version called `check_tag_completion_TEST` that tests writing in a month recency window so PLCs that are under an unscheduled trigger, discoennected or turned off aren't counted has "not completed". 

- **CC the controls tech of each mill's workbook when sending to the agent** - Currently by default, the payload is using my email for `email_from`. For better visibility, adding a condition that sends the controls tech email for each mill workbook would work.
- **More versatile parsing** - The workbook parser breaks if the columns are out of order and if the headers don't follow the standard naming. Making it more versatile in terms of parsing avoids these hiccups.
- **Normalize the workbooks** - You can derive all the missing inputs so that it directly maps to the columns in the tag request form (i.e. min max, units) from the other columns in the workbook. This may take some time though.
- **Making it easier to delete faulty tag requests** - test runs and faulty tag requests leak into `tagrequestItems` and `tagrequestLog` making deleting these records time-consuming. 



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


- **Move to ROC-TAGSRUS (hosted).** LLM (`api.anthropic.com:443`) and SQL
  (`ROC-HIST01:1433`) outbound are already open there. The unsolved piece is how
  the workbook `.xlsx` files. Microsoft 365 is already in roc-tagsrus so you should theoretically be able to derive the onedrive sync to the roc-tagsrus machine. Additionally, you will need to move the PLC Tag folder to a centralized area (i.e. The TagsRUs folder) for security purposes. IF YOU DO THIS YOU WILL NEED TO IMPLEMENT ACCESS CONTROL

- **Change credentials to company LLM API account** - Current API Key is using my own personal account

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


### Known bugs / cleanup

- **Data inputs not showing in some views / "what is review?"** — a couple of
  UI/data-display inconsistencies noted during testing; audit and resolve.

- **Filename→mill matching is finnicky** - file names in onedrive folder must share the same naming conventions

- **First-table header detection.** Some workbooks have a non-standard first
  header the parser can miss (caused an ANG undercount once, fixed by
  standardizing that header). A more robust first-header catch would remove the
  manual step.

- **Remove non-relevant tags** (e.g. fuel tags for mills that don't have them)
  and reconcile duplicate-equipment rows.

### Possible workbook schema additions

- **Min/Max columns** and **Units of Measure (UoM) columns** — considered for the
  workbooks; not yet added. (Units can often be derived from TagPurpose/Notes.)

- **Comments persistence.** Rollout-report comments are session-only (cleared on
  reload). If they need to survive, store them (SQL or local file).

### Other


- **Scalability.** General goal to keep the parser/flow easy to extend as mills
  and tables grow (header detection is already layout-tolerant; keep it that way).
- **Faster loading/startup time** 
- **Updated PowerBI (Home) section**
- 

See `ARCHITECTURE.md` for the technical detail behind all of this.
