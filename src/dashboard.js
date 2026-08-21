/* ===================================================================
   dashboard.js — native (no Power BI) dashboard for the Home tab.
   Pulls the same SQL the Power BI report used, via /dashboard:
     - Tag requests per mill (submitted / in progress / complete)
     - Tag coverage per mill (PLC / optimizer tags recording)
   Both rendered as WF-colored CLUSTERED bar charts, sorted by total
   tags descending (most on the left).
   =================================================================== */

const DASH_API = WB_API;   // same backend host as the workbook API

// WF palette for the bars.
const DASH_COLORS = {
  submitted:  "#60a5fa",   // light blue
  inProgress: "#d97706",   // amber
  complete:   "#017e3a",   // WF green
  plc:        "#017e3a",   // WF green
  optimizer:  "#7cc59f",   // light WF green
};

let dashLoaded = false;

async function dashInit() {
  if (dashLoaded) return;   // render once per session; reload on demand elsewhere
  await dashRender();
}

async function dashRender() {
  const reqHost = document.getElementById("dashRequests");
  const covHost = document.getElementById("dashCoverage");
  reqHost.innerHTML = `<div class="wb-loading">Loading dashboard…</div>`;
  covHost.innerHTML = "";

  let mills;
  try {
    const res = await fetch(`${DASH_API}/dashboard`);
    mills = (await res.json()).mills || [];
  } catch (err) {
    reqHost.innerHTML = `<div class="wb-error">Couldn't load dashboard. ${err.message}</div>`;
    return;
  }
  dashLoaded = true;

  // ---- Tag requests: sort by total submitted, desc ----
  const reqData = [...mills].sort((a, b) => b.submitted - a.submitted);
  drawClustered(
    reqHost,
    reqData.map(m => ({
      label: m.millCode,
      series: [
        { key: "submitted",  value: m.submitted },
        { key: "inProgress", value: m.inProgress },
        { key: "complete",   value: m.complete },
      ],
    }))
  );
  legend("dashRequestsLegend", [
    ["Submitted", DASH_COLORS.submitted],
    ["In Progress", DASH_COLORS.inProgress],
    ["Complete", DASH_COLORS.complete],
  ]);

  // ---- Tag coverage: sort by total recording (plc+opt), desc ----
  const covData = [...mills].sort(
    (a, b) => (b.plcTags + b.optimizerTags) - (a.plcTags + a.optimizerTags)
  );
  drawClustered(
    covHost,
    covData.map(m => ({
      label: m.millCode,
      connected: m.connected,
      series: [
        { key: "plc",       value: m.plcTags },
        { key: "optimizer", value: m.optimizerTags },
      ],
    }))
  );
  legend("dashCoverageLegend", [
    ["PLC", DASH_COLORS.plc],
    ["Optimizer", DASH_COLORS.optimizer],
  ]);
}

/* Draw a clustered (grouped) bar chart as inline SVG. */
function drawClustered(host, groups) {
  const W = Math.max(720, groups.length * 90);
  const H = 340;
  const pad = { top: 20, right: 16, bottom: 46, left: 52 };
  const plotW = W - pad.left - pad.right;
  const plotH = H - pad.top - pad.bottom;

  const maxVal = Math.max(
    1, ...groups.flatMap(g => g.series.map(s => s.value))
  );
  // "nice" round top for the axis
  const top = niceCeil(maxVal);

  const groupW = plotW / groups.length;
  const barGap = 6;
  const nSeries = groups[0]?.series.length || 1;
  const barW = (groupW - barGap * (nSeries + 1)) / nSeries;

  const y = v => pad.top + plotH - (v / top) * plotH;

  // gridlines + y labels (5 steps)
  let grid = "";
  for (let i = 0; i <= 5; i++) {
    const val = Math.round((top / 5) * i);
    const yy = y(val);
    grid += `<line x1="${pad.left}" y1="${yy}" x2="${W - pad.right}" y2="${yy}"
              stroke="#e2e8f0" stroke-width="1"/>`;
    grid += `<text x="${pad.left - 8}" y="${yy + 4}" text-anchor="end"
              font-size="11" fill="#64748b">${val}</text>`;
  }

  let bars = "";
  const SERIES_LABELS = {
    submitted: "Submitted", inProgress: "In Progress", complete: "Complete",
    plc: "PLC", optimizer: "Optimizer",
  };
  groups.forEach((g, gi) => {
    const gx = pad.left + gi * groupW;
    g.series.forEach((s, si) => {
      const x = gx + barGap + si * (barW + barGap);
      const h = (s.value / top) * plotH;
      const yy = pad.top + plotH - h;
      const label = SERIES_LABELS[s.key] || s.key;
      bars += `<rect class="dash-bar" x="${x}" y="${yy}" width="${barW}" height="${h}"
                fill="${DASH_COLORS[s.key]}" rx="2"
                data-mill="${g.label}" data-series="${label}" data-value="${s.value}"></rect>`;
    });
    // x label (mill code); mark not-connected with an asterisk
    const star = g.connected === false ? "*" : "";
    bars += `<text x="${gx + groupW / 2}" y="${H - pad.bottom + 18}"
              text-anchor="middle" font-size="12" fill="#334155"
              font-weight="600">${g.label}${star}</text>`;
  });

  const note = groups.some(g => g.connected === false)
    ? `<text x="${W - pad.right}" y="${H - 6}" text-anchor="end"
        font-size="10" fill="#94a3b8">* not connected to historian</text>`
    : "";

  host.innerHTML =
    `<div class="dash-wrap">
       <svg viewBox="0 0 ${W} ${H}" class="dash-svg" preserveAspectRatio="xMinYMin meet">
         ${grid}${bars}${note}
       </svg>
       <div class="dash-tip" style="display:none;"></div>
     </div>`;

  // Hover tooltip that follows the cursor.
  const tip = host.querySelector(".dash-tip");
  host.querySelectorAll(".dash-bar").forEach(bar => {
    bar.addEventListener("mousemove", e => {
      tip.innerHTML =
        `<b>${bar.dataset.mill}</b><br>${bar.dataset.series}: ${bar.dataset.value}`;
      tip.style.display = "block";
      const r = host.getBoundingClientRect();
      tip.style.left = (e.clientX - r.left + 12) + "px";
      tip.style.top = (e.clientY - r.top + 12) + "px";
    });
    bar.addEventListener("mouseleave", () => { tip.style.display = "none"; });
  });
}

function legend(id, items) {
  const el = document.getElementById(id);
  if (!el) return;
  el.innerHTML = items.map(([label, color]) =>
    `<span class="dash-leg-item">
       <span class="dash-leg-swatch" style="background:${color}"></span>${label}
     </span>`).join("");
}

// Round a max value up to a clean axis top.
function niceCeil(v) {
  if (v <= 5) return 5;
  const mag = Math.pow(10, Math.floor(Math.log10(v)));
  const n = v / mag;
  const nice = n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10;
  return nice * mag;
}