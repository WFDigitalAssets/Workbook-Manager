/* ===================================================================
   Workbook Manager — client module
   ---------------------------------------------------------------
   Pattern A: a backend (roc-laa01 / Python service) reads the
   OneDrive-synced .xlsx workbooks and exposes JSON endpoints. This
   file renders that JSON. Right now it runs against MOCK data so the
   UI works with no backend; flip USE_MOCK to false once the real
   endpoints exist.

   The embedded Excel views (SharePoint iframes) are configured in
   EMBED_URLS below and are independent of USE_MOCK — they render the
   real view-only workbooks regardless of whether the data layer is
   mock or live.

   DATA CONTRACT (what the backend must return)
   -------------------------------------------------------------------
   GET /workbooks/overview
     -> { mills: [ { code, name, counts:{new,inProgress,faulty,complete},
                     total, flags } ] }

   GET /workbooks/<code>/detail
     -> { code, name,
          tables: [ { name,
                      rows: [ { status, tagName, plcTag, plcName, plcIp,
                                dataType, description } ] } ] }

   GET /workbooks/<code>/reconcile
     -> { code, name,
          suggestions: [ { row, table, tagName, kind, current,
                           suggested, reason, confidence } ] }

   POST /workbooks/<code>/submit-new
     body { tags: [ ...rows... ] }
   =================================================================== */

const USE_MOCK = false;
const WB_API = "http://localhost:8100";

/* ===================================================================
   Embedded Excel workbooks (SharePoint, view-only).
   One embed URL per mill. Dimensions are applied in the render, not
   here, so they're consistent across mills.
   =================================================================== */

const EMBED_URLS = {
  OPE: "https://westfraser4-my.sharepoint.com/personal/reymart_velasco_westfraser_com/_layouts/15/Doc.aspx?sourcedoc={bce8624f-4582-4efa-8ed9-1146ad476ae5}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True",
  JOY: "https://westfraser4-my.sharepoint.com/personal/ron_fachini_westfraser_com/_layouts/15/Doc.aspx?sourcedoc={cd0225e3-63e8-4282-adff-9007bb462792}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True",
  MAP: "https://westfraser4-my.sharepoint.com/personal/ron_fachini_westfraser_com/_layouts/15/Doc.aspx?sourcedoc={01935d33-1bda-4bd1-a83e-9fb2555e7a1c}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True",
  HNM: "https://westfraser4-my.sharepoint.com/personal/ron_fachini_westfraser_com/_layouts/15/Doc.aspx?sourcedoc={b6ce746b-83b5-4961-abad-858e65bb9295}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True",
  ANG: "https://westfraser4-my.sharepoint.com/personal/ron_fachini_westfraser_com/_layouts/15/Doc.aspx?sourcedoc={c14a6c48-56d8-4ee1-af1d-55334104b01a}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True",
  ARM: "https://westfraser4-my.sharepoint.com/personal/valentin_kurz_westfraser_com/_layouts/15/Doc.aspx?sourcedoc={5959b74d-3482-41a2-933b-f40780221821}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True",
  NEW: "https://westfraser4-my.sharepoint.com/personal/ron_fachini_westfraser_com/_layouts/15/Doc.aspx?sourcedoc={7788e7cb-cf54-4680-a3ef-fadd71f16369}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True",
  DGM: "https://westfraser4-my.sharepoint.com/personal/valentin_kurz_westfraser_com/_layouts/15/Doc.aspx?sourcedoc={f4f737cf-3fcc-4fc7-bf9e-7e4149ae8d47}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True",
  NWB: "https://westfraser4-my.sharepoint.com/personal/valentin_kurz_westfraser_com/_layouts/15/Doc.aspx?sourcedoc={7bb1f622-ae68-4152-86f8-2ffa7b7b59be}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True",
  RUS: "https://westfraser4-my.sharepoint.com/personal/ron_fachini_westfraser_com/_layouts/15/Doc.aspx?sourcedoc={a5deebf6-0562-47f8-8119-3bd584347414}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True",
};

// Embed dimensions. Width fills the container; height fills most of the
// viewport (minus header/nav/filter chrome) so the sheet nearly fills the page.
const EMBED_HEIGHT = "calc(100vh - 220px)";

/* ===================================================================
   Data layer — the only place that talks to the backend.
   =================================================================== */

async function wbFetchOverview() {
  if (USE_MOCK) return MOCK_OVERVIEW;
  const res = await fetch(`${WB_API}/workbooks/overview`);
  return res.json();
}

async function wbFetchDetail(code) {
  if (USE_MOCK) return MOCK_DETAIL[code] || { code, name: code, tables: [] };
  const res = await fetch(`${WB_API}/workbooks/${code}/detail`);
  return res.json();
}

async function wbFetchReconcile(code) {
  if (USE_MOCK) return MOCK_RECONCILE[code] || { code, name: code, suggestions: [] };
  const res = await fetch(`${WB_API}/workbooks/${code}/reconcile`);
  return res.json();
}

async function wbSubmitNew(code, tags) {
  if (USE_MOCK) {
    console.log("MOCK submit-new", code, tags);
    return { ok: true, submitted: tags.length };
  }
  const res = await fetch(`${WB_API}/workbooks/${code}/submit-new`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ tags }),
  });
  return res.json();
}

/* ===================================================================
   State
   =================================================================== */

let wbCurrentMill = null;
let wbDetailFilter = "all";
let wbLastRunIso = null;      // most recent refresh time, held for the ticking label
let wbLastRunTimer = null;    // interval that keeps "Updated X ago" current

/* ===================================================================
   OVERVIEW — all mills, per-mill counts, completion, flag badge
   =================================================================== */

async function wbRenderOverview() {
  const host = document.getElementById("wbOverview");
  host.innerHTML = `<div class="wb-loading">Loading workbooks…</div>`;

  let data;
  try {
    data = await wbFetchOverview();
  } catch (err) {
    host.innerHTML = `<div class="wb-error">Couldn't load workbooks. ${err.message}</div>`;
    return;
  }

  const cards = data.mills.map(m => {
    const c = m.counts;
    const pct = m.total ? Math.round((c.complete / m.total) * 100) : 0;
    const flagBadge = m.flags
      ? `<span class="wb-flag-badge">${m.flags} to review</span>`
      : "";
    return `
      <button class="wb-mill-card" onclick="wbOpenMill('${m.code}')">
        <div class="wb-mill-head">
          <span class="wb-mill-code">${m.code}</span>
          <span class="wb-mill-name">${m.name}</span>
          ${flagBadge}
        </div>
        <div class="wb-progress">
          <div class="wb-progress-bar" style="width:${pct}%"></div>
        </div>
        <div class="wb-mill-stats">
          <span class="wb-stat"><b>${c.complete}</b> of ${m.total} complete</span>
        </div>
        <div class="wb-mill-counts">
          <span class="wb-pill wb-pill-new">${c.new} new</span>
          <span class="wb-pill wb-pill-prog">${c.inProgress} in progress</span>
          <span class="wb-pill wb-pill-fault">${c.faulty} faulty</span>
          <span class="wb-pill wb-pill-done">${c.complete} complete</span>
        </div>
      </button>`;
  }).join("");

  const roll = data.mills.reduce((a, m) => {
    a.new += m.counts.new; a.prog += m.counts.inProgress;
    a.fault += m.counts.faulty; a.done += m.counts.complete;
    a.total += m.total; a.flags += m.flags || 0;
    return a;
  }, { new: 0, prog: 0, fault: 0, done: 0, total: 0, flags: 0 });

  // Rollout report table — same data as the cards, in the report layout
  // Reymart uses (dots colored by status, "complete of total", blank comments).
  const reportRows = data.mills.map(m => {
    const c = m.counts;
    return `
      <tr>
        <td class="wb-rep-mill">${m.code}</td>
        <td>${c.new}</td>
        <td>${c.inProgress}</td>
        <td>${c.faulty}</td>
        <td>${c.complete} of ${m.total}</td>
        <td class="wb-rep-comment">
          <input type="text" class="wb-comment-input" placeholder="Add a comment…">
        </td>
      </tr>`;
  }).join("");

  const updated = wbFormatLastRun(data.last_run);
  const reportTable = `
    <div class="card wb-report-card">
      <div class="wb-report-head">
        <div class="section-label">Rollout report</div>
        <div class="wb-report-actions">
          <span class="wb-last-updated" id="wbLastUpdated">${updated}</span>
          <button class="btn btn-grey" onclick="wbRefreshAll()" id="wbRefreshBtn">Refresh numbers</button>
          <button class="btn btn-grey" onclick="wbCopyReport()">Copy table</button>
        </div>
      </div>
      <div class="table-wrap">
        <table class="wb-report" id="wbReportTable">
          <thead>
            <tr>
              <th>Mill</th>
              <th>New <span class="wb-dot wb-dot-new"></span></th>
              <th>In Progress <span class="wb-dot wb-dot-prog"></span></th>
              <th>Faulty <span class="wb-dot wb-dot-fault"></span></th>
              <th>Completed <span class="wb-dot wb-dot-done"></span></th>
              <th>Comments</th>
            </tr>
          </thead>
          <tbody>${reportRows}</tbody>
        </table>
      </div>
    </div>`;

  host.innerHTML = `
    <div class="card">
      <div class="section-label">All mills — rollout status</div>
      <div class="wb-rollup">
        <div class="wb-roll-item"><b>${roll.total}</b><span>total tags</span></div>
        <div class="wb-roll-item"><b>${roll.done}</b><span>complete</span></div>
        <div class="wb-roll-item"><b>${roll.prog}</b><span>in progress</span></div>
        <div class="wb-roll-item"><b>${roll.new}</b><span>new</span></div>
        <div class="wb-roll-item"><b>${roll.fault}</b><span>faulty</span></div>
        <div class="wb-roll-item wb-roll-flags"><b>${roll.flags}</b><span>to review</span></div>
      </div>
    </div>
    <div class="wb-mill-grid">${cards}</div>
    ${reportTable}`;

  // Keep the "Updated X ago" label current without re-hitting the server:
  // store the timestamp and recompute the label text every 60s from it.
  wbLastRunIso = data.last_run || null;
  if (wbLastRunTimer) clearInterval(wbLastRunTimer);
  wbLastRunTimer = setInterval(wbTickLastUpdated, 60000);
}

/* Recompute just the "Updated X ago" label from the stored timestamp. Pure
   local math — no server call. */
function wbTickLastUpdated() {
  const el = document.getElementById("wbLastUpdated");
  if (!el) { clearInterval(wbLastRunTimer); wbLastRunTimer = null; return; }
  el.textContent = wbFormatLastRun(wbLastRunIso);
}

/* Copy the rollout report as styled HTML so it pastes into email/Word with
   the green header and status dots intact. */
function _esc(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/* Format the last-run timestamp as a friendly "updated X ago" string. */
function wbFormatLastRun(iso) {
  if (!iso) return "Not yet refreshed";
  const then = new Date(iso);
  if (isNaN(then)) return "Not yet refreshed";
  const mins = Math.round((Date.now() - then.getTime()) / 60000);
  let ago;
  if (mins < 1)        ago = "just now";
  else if (mins === 1) ago = "1 minute ago";
  else if (mins < 60)  ago = `${mins} minutes ago`;
  else {
    const hrs = Math.round(mins / 60);
    ago = hrs === 1 ? "1 hour ago" : `${hrs} hours ago`;
  }
  return `Updated ${ago}`;
}

/* Run the full refresh (run_program) then reload the overview. */
async function wbRefreshAll() {
  const btn = document.getElementById("wbRefreshBtn");
  if (USE_MOCK) { alert("(mock) would run refresh"); return; }
  if (btn) { btn.disabled = true; btn.textContent = "Refreshing…"; }
  try {
    const res = await fetch(`${WB_API}/workbooks/run`, { method: "POST" });
    if (!res.ok) {
      const d = await res.json().catch(() => ({}));
      alert("Refresh failed: " + (d.detail || res.status));
    } else {
      await wbRenderOverview();   // reloads counts + the new last-updated time
    }
  } catch (err) {
    alert("Couldn't reach the server: " + err.message);
  }
  if (btn && document.body.contains(btn)) { btn.disabled = false; btn.textContent = "Refresh numbers"; }
}

function wbCopyReport() {
  const src = document.getElementById("wbReportTable");
  if (!src) return;

  // Rebuild with inline styles (email clients ignore external CSS / classes).
  const WF = "#017e3a";
  const dot = (color, border) =>
    `<span style="display:inline-block;width:9px;height:9px;border-radius:50%;` +
    `background:${color};${border ? "border:1px solid #999;" : ""}` +
    `margin-left:6px;vertical-align:middle;"></span>`;

  const headers = [
    "Mill",
    "New " + dot("#ffffff", true),
    "In Progress " + dot("#d97706"),
    "Faulty " + dot("#dc2626"),
    "Completed " + dot("#22c55e"),
    "Comments",
  ];
  const headHtml = headers.map(h =>
    `<th style="background:${WF};color:#fff;text-align:left;padding:8px 12px;` +
    `border:1px solid #cbd5e1;">${h}</th>`).join("");

  const bodyRows = Array.from(src.querySelectorAll("tbody tr")).map(tr => {
    const tds = Array.from(tr.children).map(td => {
      // If the cell holds an input (the Comments cell), copy its typed value;
      // otherwise copy the text.
      const input = td.querySelector("input");
      const value = input ? input.value : td.textContent;
      return `<td style="padding:7px 12px;border:1px solid #cbd5e1;">${_esc(value)}</td>`;
    }).join("");
    return `<tr>${tds}</tr>`;
  }).join("");

  const html =
    `<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;">` +
    `<thead><tr>${headHtml}</tr></thead><tbody>${bodyRows}</tbody></table>`;

  if (navigator.clipboard && navigator.clipboard.write) {
    const blob = new Blob([html], { type: "text/html" });
    navigator.clipboard.write([new ClipboardItem({ "text/html": blob })])
      .then(() => alert("Report copied. Paste into your email."))
      .catch(() => wbFallbackCopy(html));
  } else {
    wbFallbackCopy(html);
  }
}

function wbFallbackCopy(html) {
  const tmp = document.createElement("div");
  tmp.style.position = "fixed";
  tmp.style.left = "-9999px";
  tmp.contentEditable = "true";
  tmp.innerHTML = html;
  document.body.appendChild(tmp);
  const range = document.createRange();
  range.selectNodeContents(tmp);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  try { document.execCommand("copy"); alert("Report copied. Paste into your email."); }
  catch { alert("Couldn't copy automatically."); }
  document.body.removeChild(tmp);
}

/* ===================================================================
   DETAIL — one mill. "All" shows the embedded Excel file;
   the status filters show the parsed data table.
   =================================================================== */

async function wbOpenMill(code) {
  wbCurrentMill = code;
  wbDetailFilter = "all";
  showWorkbookSubview("detail");
  await wbRenderDetail();
}

async function wbRenderDetail() {
  const host = document.getElementById("wbDetail");
  host.innerHTML = `<div class="wb-loading">Loading ${wbCurrentMill}…</div>`;

  let data;
  try {
    data = await wbFetchDetail(wbCurrentMill);
  } catch (err) {
    host.innerHTML = `<div class="wb-error">Couldn't load ${wbCurrentMill}. ${err.message}</div>`;
    return;
  }

  const filterBtns = ["all", "new", "in progress", "faulty", "complete", "llm"].map(f => {
    const active = wbDetailFilter === f ? " active" : "";
    const label = f === "all" ? "All (workbook)"
                : f === "llm" ? "LLM suggestions"
                : f.replace(/\b\w/g, ch => ch.toUpperCase());
    return `<button class="wb-filter${active}" onclick="wbSetFilter('${f}')">${label}</button>`;
  }).join("");

  // Body depends on the active filter.
  let body;
  if (wbDetailFilter === "all") {
    // Show the embedded, view-only Excel workbook.
    const url = EMBED_URLS[data.code];
    body = url
      ? `<div class="card wb-embed-card">
           <iframe
             title="${data.name} workbook"
             src="${url}"
             class="wb-embed"
             style="width:100%; height:${EMBED_HEIGHT}; border:none;"
             frameborder="0"
             scrolling="no">
           </iframe>
         </div>`
      : `<div class="empty-state">No embedded workbook configured for ${data.code}.</div>`;
  } else if (wbDetailFilter === "llm") {
    // LLM suggestions: nothing runs until the user clicks the button (costs money).
    body = `
      <div class="card">
        <div class="wb-llm-intro">
          <p>Get AI suggestions for flagged tags (incomplete rows, multiple PLC paths
             in one cell, and tags sharing PLC identity). Suggestions are advisory —
             review before applying.</p>
          <button class="btn btn-primary" id="wbLlmRunBtn" onclick="wbRunLlm('${data.code}')">
            Get suggestions
          </button>
        </div>
        <div id="wbLlmResults"></div>
      </div>`;
  } else {
    // Show the parsed data table filtered by status.
    const tables = data.tables.map(t => {
      const rows = t.rows
        .filter(r => (r.status || "").trim().toLowerCase() === wbDetailFilter ||
                     (wbDetailFilter === "new" && !(r.status || "").trim()))
        .map(r => {
          const status = (r.status || "").trim() || "New";
          const cls = "wb-status-" + status.toLowerCase().replace(/\s+/g, "");
          return `
            <tr>
              <td><span class="wb-status ${cls}">${status}</span></td>
              <td>${r.tagName || ""}</td>
              <td>${r.plcTag || ""}</td>
              <td>${r.plcName || ""}</td>
              <td>${r.plcIp || ""}</td>
              <td>${r.dataType || ""}</td>
              <td>${r.description || ""}</td>
            </tr>`;
        }).join("");

      if (!rows) return "";
      return `
        <div class="card">
          <div class="wb-table-head">${t.name}</div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Status</th><th>Tag name</th><th>PLC tag</th>
                  <th>PLC name</th><th>PLC IP</th><th>Type</th><th>Description</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>
          </div>
        </div>`;
    }).join("");
    body = tables || `<div class="empty-state">No tags match this filter.</div>`;
  }

  // "Send new tags to agent" — only in the New filter, only if there are any.
  // Opens an Outlook draft on the server machine (roc-wfm-rv) with the new
  // tags in request-column format, for the user to review and send.
  let sendBar = "";
  if (wbDetailFilter === "new") {
    const newCount = data.tables.reduce((n, t) =>
      n + t.rows.filter(r => {
        const s = (r.status || "").trim().toLowerCase();
        return s === "new" || s === "";
      }).length, 0);
    if (newCount > 0) {
      sendBar = `
        <div class="wb-send-bar">
          <button class="btn btn-primary" onclick="wbSendNew('${data.code}')">
            Send ${newCount} new tag${newCount === 1 ? "" : "s"} to agent
          </button>
          <span class="wb-send-note">Opens an Outlook draft to review before sending.</span>
        </div>`;
    }
  }

  host.innerHTML = `
    <div class="wb-detail-bar">
      <button class="btn btn-grey" onclick="showWorkbookSubview('overview')">← All mills</button>
      <h2 class="wb-detail-title">${data.name} <span>(${data.code})</span></h2>
      <button class="btn btn-primary" onclick="wbOpenReconcile('${data.code}')">Review flags</button>
    </div>
    <div class="wb-filters">${filterBtns}</div>
    ${sendBar}
    ${body}`;
}

async function wbSendNew(code) {
  if (USE_MOCK) { alert("(mock) would open Outlook draft for " + code); return; }
  try {
    const res = await fetch(`${WB_API}/workbooks/${code}/send-new`, { method: "POST" });
    const data = await res.json();
    if (!res.ok) {
      alert("Couldn't open draft: " + (data.detail || res.status));
      return;
    }
    alert(`Outlook draft opened with ${data.opened} new tag(s). Review, then send.`);
  } catch (err) {
    alert("Couldn't reach the server: " + err.message);
  }
}

/* Run the one-per-mill LLM call and render the suggestions. Only fires when the
   user clicks the button — never automatically. */
async function wbRunLlm(code) {
  const btn = document.getElementById("wbLlmRunBtn");
  const out = document.getElementById("wbLlmResults");
  if (USE_MOCK) { out.innerHTML = "<p>(mock) would call the LLM.</p>"; return; }
  if (btn) { btn.disabled = true; btn.textContent = "Thinking… (this can take a bit)"; }
  out.innerHTML = `<div class="wb-loading">Analyzing flagged tags…</div>`;
  try {
    const res = await fetch(`${WB_API}/workbooks/${code}/llm-suggest`, { method: "POST" });
    const data = await res.json();
    if (!res.ok) {
      out.innerHTML = `<div class="wb-error">Couldn't get suggestions: ${data.detail || res.status}</div>`;
    } else {
      out.innerHTML = wbRenderLlm(data);
    }
  } catch (err) {
    out.innerHTML = `<div class="wb-error">Couldn't reach the server: ${err.message}</div>`;
  }
  if (btn) { btn.disabled = false; btn.textContent = "Get suggestions"; }
}

function wbRenderLlm(data) {
  const c = data.counts || {};
  const groups = data.shared_group_comments || [];
  const summary = `
    <div class="wb-llm-summary">
      Flagged: ${c.INCOMPLETE || 0} incomplete · ${c.MULTI_PATH || 0} multi-path ·
      ${c.SHARED_IDENTITY_GROUPS || groups.length || 0} shared-path group(s)
    </div>`;

  const hasRows = data.suggestions && data.suggestions.length;
  if (!hasRows && !groups.length) {
    return summary + `<div class="empty-state">No suggestions — nothing flagged.</div>`;
  }

  // --- Per-row suggestions (INCOMPLETE, MULTI_PATH) ---
  const cards = (data.suggestions || []).map(s => {
    const conf = (s.confidence || "medium").toLowerCase();
    const flags = (s.flag_types || []).join(", ");
    const orig = s.original_row || {};
    const suggestedRows = (s.suggested_rows || []).map(r => wbRowLine(r)).join("");
    const copy = (s.copyable || "").replace(/"/g, "&quot;");
    return `
      <div class="wb-sug">
        <div class="wb-sug-head">
          <span class="wb-conf wb-conf-${conf}">${conf}</span>
          <span class="wb-sug-flags">${flags}</span>
          <button class="btn btn-grey wb-sug-copy" data-copy="${copy}"
                  onclick="wbCopySuggestion(this)">Copy suggested row(s)</button>
        </div>
        <div class="wb-sug-reason">${s.reason || ""}</div>
        <div class="wb-sug-cols">
          <div class="wb-sug-col">
            <div class="wb-sug-label">Original</div>
            ${wbRowLine(orig)}
          </div>
          <div class="wb-sug-col">
            <div class="wb-sug-label">Suggested</div>
            ${suggestedRows}
          </div>
        </div>
      </div>`;
  }).join("");

  // --- Shared-path groups: one card per group listing all its tags + a comment ---
  const groupCards = groups.map(g => {
    const conf = (g.confidence || "medium").toLowerCase();
    const tags = (g.tags || []);
    const path = tags.length ? (tags[0].plc_tag || "") : "";
    const tagLines = tags.map(t => `
      <div class="wb-rowline">
        <b>${t.table || ""}</b> · ${t.tag_name || "<i>(no name)</i>"}
        <div class="wb-rowline-detail">
          PLC: ${t.plc_tag || ""} | ${t.plc_name || ""} | ${t.plc_ip || ""}
          &nbsp;·&nbsp; ${t.data_type || "—"}
        </div>
      </div>`).join("");
    return `
      <div class="wb-sug wb-sug-group">
        <div class="wb-sug-head">
          <span class="wb-conf wb-conf-${conf}">${conf}</span>
          <span class="wb-sug-flags">${tags.length} tags share PLC path: <code>${path}</code></span>
        </div>
        <div class="wb-sug-reason">${g.comment || ""}</div>
        <div class="wb-sug-label">Tags sharing this path</div>
        ${tagLines}
      </div>`;
  }).join("");

  let out = summary;
  if (cards) {
    out += `<div class="wb-llm-section-label">Row suggestions</div>` + cards;
  }
  if (groupCards) {
    out += `<div class="wb-llm-section-label">Shared PLC paths</div>` + groupCards;
  }
  return out;
}

/* Render one row dict as a compact labeled line. */
function wbRowLine(r) {
  const f = (k) => (r[k] == null ? "" : String(r[k]));
  return `
    <div class="wb-rowline">
      <b>${f("table")}</b> · ${f("tag_name") || "<i>(no name)</i>"}
      <div class="wb-rowline-detail">
        PLC: ${f("plc_tag")} | ${f("plc_name")} | ${f("plc_ip")}
        &nbsp;·&nbsp; Type: ${f("data_type") || "—"}
        &nbsp;·&nbsp; ${f("tag_purpose") || ""}
      </div>
    </div>`;
}

function wbCopySuggestion(btn) {
  const text = (btn.getAttribute("data-copy") || "").replace(/&quot;/g, '"');
  navigator.clipboard.writeText(text)
    .then(() => { btn.textContent = "Copied!"; setTimeout(() => btn.textContent = "Copy suggested row(s)", 1500); })
    .catch(() => alert("Couldn't copy. Here it is:\n\n" + text));
}

function wbSetFilter(f) {
  wbDetailFilter = f;
  wbRenderDetail();}

/* ===================================================================
   RECONCILE — flagged issues + suggested fixes, human approves
   =================================================================== */

async function wbOpenReconcile(code) {
  wbCurrentMill = code;
  showWorkbookSubview("reconcile");
  await wbRenderReconcile();
}

async function wbRenderReconcile() {
  const host = document.getElementById("wbReconcile");
  host.innerHTML = `<div class="wb-loading">Checking ${wbCurrentMill} against SQL…</div>`;

  let data;
  try {
    data = await wbFetchReconcile(wbCurrentMill);
  } catch (err) {
    host.innerHTML = `<div class="wb-error">Couldn't reconcile ${wbCurrentMill}. ${err.message}</div>`;
    return;
  }

  const rows = data.suggestions.map(s => {
    const conf = s.confidence === "high"
      ? `<span class="wb-conf wb-conf-high">high</span>`
      : `<span class="wb-conf wb-conf-low">needs review</span>`;
    return `
      <tr>
        <td><span class="wb-kind wb-kind-${s.kind}">${s.kind}</span></td>
        <td>${s.table}</td>
        <td>${s.tagName}</td>
        <td>${s.current || "—"}</td>
        <td class="wb-suggested">${s.suggested || "—"}</td>
        <td>${s.reason || ""} ${conf}</td>
      </tr>`;
  }).join("");

  host.innerHTML = `
    <div class="wb-detail-bar">
      <button class="btn btn-grey" onclick="showWorkbookSubview('detail')">← ${data.name}</button>
      <h2 class="wb-detail-title">Flags for ${data.name}</h2>
    </div>
    <div class="card">
      <p class="wb-reconcile-note">
        Suggestions only — nothing is changed automatically. Review each,
        then apply in the workbook yourself. High-confidence items are
        rule-based; "needs review" items are judgment calls.
      </p>
      ${data.suggestions.length ? `
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Type</th><th>Table</th><th>Tag</th>
                <th>Current</th><th>Suggested</th><th>Why</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </div>` : `<div class="empty-state">No flags — everything reconciles.</div>`}
    </div>`;
}

/* ===================================================================
   Subview switching within the Workbook tab
   =================================================================== */

function showWorkbookSubview(which) {
  ["overview", "detail", "reconcile"].forEach(v => {
    document.getElementById("wb" + v.charAt(0).toUpperCase() + v.slice(1))
      .style.display = (v === which) ? "block" : "none";
  });
  if (which === "overview") wbRenderOverview();
}

/* ===================================================================
   MOCK DATA — remove or ignore once USE_MOCK = false
   =================================================================== */

const MOCK_OVERVIEW = {
  mills: [
    { code: "ANG", name: "Angelina",     counts: { new: 0,  inProgress: 0,  faulty: 0, complete: 53  }, total: 862,  flags: 2 },
    { code: "ARM", name: "Armour",       counts: { new: 0,  inProgress: 0,  faulty: 0, complete: 0   }, total: 879,  flags: 0 },
    { code: "DGM", name: "Dudley",       counts: { new: 0,  inProgress: 0,  faulty: 0, complete: 1   }, total: 851,  flags: 1 },
    { code: "HNM", name: "Henderson",    counts: { new: 0,  inProgress: 0,  faulty: 0, complete: 5   }, total: 860,  flags: 0 },
    { code: "JOY", name: "Joyce",        counts: { new: 0,  inProgress: 13, faulty: 9, complete: 122 }, total: 1022, flags: 4 },
    { code: "MAP", name: "Maplesville",  counts: { new: 4,  inProgress: 0,  faulty: 0, complete: 213 }, total: 882,  flags: 3 },
    { code: "NEW", name: "New Boston",   counts: { new: 0,  inProgress: 0,  faulty: 0, complete: 0   }, total: 860,  flags: 0 },
    { code: "NWB", name: "Newberry",     counts: { new: 0,  inProgress: 0,  faulty: 0, complete: 0   }, total: 860,  flags: 0 },
    { code: "OPE", name: "Opelika",      counts: { new: 2,  inProgress: 5,  faulty: 0, complete: 40  }, total: 770,  flags: 1 },
    { code: "RUS", name: "Russellville", counts: { new: 0,  inProgress: 0,  faulty: 0, complete: 0   }, total: 860,  flags: 0 },
  ],
};

const MOCK_DETAIL = {
  JOY: {
    code: "JOY", name: "Joyce",
    tables: [
      { name: "DryStackerDataTable", rows: [
        { status: "Complete",    tagName: "BoardLength",   plcTag: "Board_Len", plcName: "JoyceCDK", plcIp: "192.168.10.5", dataType: "float", description: "Board length at dry stacker" },
        { status: "In Progress", tagName: "BoardWidth",    plcTag: "Board_Wid", plcName: "JoyceCDK", plcIp: "192.168.10.5", dataType: "float", description: "Board width at dry stacker" },
        { status: "Faulty",      tagName: "ProductLength", plcTag: "Prod_Len",  plcName: "JoyceCDK", plcIp: "192.168.10.5", dataType: "float", description: "Product length" },
      ]},
      { name: "EdgerDataTable", rows: [
        { status: "Complete",    tagName: "PieceCount",    plcTag: "Piece_Cnt", plcName: "JoyceEdg", plcIp: "192.168.10.8", dataType: "int",   description: "Pieces through edger" },
        { status: "",            tagName: "LineSpeed",     plcTag: "Line_Spd",  plcName: "JoyceEdg", plcIp: "192.168.10.8", dataType: "float", description: "Edger line speed" },
      ]},
    ],
  },
  MAP: {
    code: "MAP", name: "Maplesville",
    tables: [
      { name: "FuelEnergyDataTable", rows: [
        { status: "", tagName: "FuelUsageTotal", plcTag: "Fuel_Tons", plcName: "MaplesvilleCDK", plcIp: "192.168.220.51", dataType: "float", description: "Accumulated fuel usage Burner 1" },
        { status: "", tagName: "FuelUsageTotal", plcTag: "Fuel_Tons", plcName: "MaplesvilleCDK", plcIp: "192.168.220.51", dataType: "float", description: "Accumulated fuel usage Burner 2" },
      ]},
    ],
  },
};

const MOCK_RECONCILE = {
  JOY: {
    code: "JOY", name: "Joyce",
    suggestions: [
      { row: 45, table: "EdgerDataTable",      tagName: "LineSpeed",     kind: "status", current: "New",    suggested: "Complete", reason: "Tag is actively writing in SQL.",       confidence: "high" },
      { row: 12, table: "DryStackerDataTable", tagName: "ProductLength", kind: "status", current: "Faulty", suggested: "Complete", reason: "Marked faulty but data is recording.",  confidence: "high" },
    ],
  },
  MAP: {
    code: "MAP", name: "Maplesville",
    suggestions: [
      { row: 8, table: "FuelEnergyDataTable", tagName: "FuelUsageTotal", kind: "duplicate", current: "FuelUsageTotal", suggested: "FuelUsageTotalBurner1 / FuelUsageTotalBurner2", reason: "Two rows share a name but are different burners; scope and description are crossed.", confidence: "low" },
    ],
  },
};

/* ===================================================================
   Init — called by the tab switcher when Workbook tab opens
   =================================================================== */

function wbInit() {
  showWorkbookSubview("overview");
}