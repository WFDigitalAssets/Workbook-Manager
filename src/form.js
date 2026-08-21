/* ===================================================================
   Tag Request Form — client logic
   Sections:
     1. UI updates (enable/disable, tabs, urgent toggle)
     2. Row building (addRow, makeInputBoxes, rowsCount)
     3. Input validation (tag / IP / number / duplicates)
     4. Clear-all confirmation
     5. Support helpers (warnings, check/uncheck)
     6. Submit + table serialization
     7. Excel paste import
     8. Endpoint fetches (units, table names, template)
     9. Search dropdown widget
   =================================================================== */


/* ===================================================================
   1. UI UPDATES
   =================================================================== */

// Enable/disable buttons and tabs based on whether a mill is selected.
function updateUI() {
  getTableNames();

  const millSelect = document.getElementById("millSelect");

  if (millSelect.value !== "") {
    document.getElementById("addRow").disabled = false;
    document.getElementById("clrAll").disabled = false;
    document.getElementById("chkAll").disabled = false;
    document.getElementById("unchkAll").disabled = false;
    document.getElementById("submitReq").disabled = false;
    document.getElementById("manualOption").disabled = false;
    document.getElementById("pasteOption").disabled = false;
    document.getElementById("pasteTextBox").disabled = false;
    document.getElementById("importButton").disabled = false;
  }

  if (millSelect.value === "") {
    document.getElementById("addRow").disabled = true;
    document.getElementById("clrAll").disabled = true;
    document.getElementById("chkAll").disabled = true;
    document.getElementById("unchkAll").disabled = true;
    document.getElementById("submitReq").disabled = true;
    document.getElementById("manualOption").disabled = true;
    document.getElementById("pasteOption").disabled = true;
    document.getElementById("pasteTextBox").disabled = true;
    document.getElementById("importButton").disabled = true;
  }
}

// Toggle the "urgent" styling on its label.
function toggleUrgent(urgentBtn) {
  if (urgentBtn.checked) {
    urgentBtn.parentElement.classList.add("on");
  } else {
    urgentBtn.parentElement.classList.remove("on");
  }
}

// Switch between the "manual" and "paste" entry panels.
function switchTab(mode) {
  if (mode === "manual") {
    document.getElementById("manualOption").classList.toggle("active", true);
    document.getElementById("pasteOption").classList.toggle("active", false);
    document.getElementById("pastePanel").classList.toggle("visible", false);
  }

  if (mode === "paste") {
    document.getElementById("manualOption").classList.toggle("active", false);
    document.getElementById("pasteOption").classList.toggle("active", true);
    document.getElementById("pastePanel").classList.toggle("visible", true);
  }
}


/* ===================================================================
   2. ROW BUILDING
   =================================================================== */

// Build one table row. `data` comes from the Add Row button (empty) or
// from processExcelData (populated).
function addRow(data = {}) {
  const tableBody = document.getElementById("tbody");
  const row = document.createElement("tr");

  // Table name (searchable dropdown)
  const tableCell = document.createElement("td");
  tableCell.appendChild(createSearchDropdown("tableName", validTables, data.tableName || ""));
  tableCell.querySelector("input").placeholder = "e.g FormingTable";
  row.appendChild(tableCell);

  // Free-text cells
  row.appendChild(makeInputBoxes("tagName", "e.g. BladeSpeedRpm", data.tagName));
  row.appendChild(makeInputBoxes("plcAddress", "192.168.x.x", data.plcAddress));
  row.appendChild(makeInputBoxes("plcPath", "PLC tag path", data.plcPath));
  row.appendChild(makeInputBoxes("dataType", "— type —", data.dataType));

  // Units (searchable dropdown)
  const unitsCell = document.createElement("td");
  unitsCell.appendChild(
    createSearchDropdown("units", validUnits.map(u => u.uom + " (" + u.description + ")"), data.units || "")
  );
  unitsCell.querySelector("input").placeholder = "e.g Int";
  row.appendChild(unitsCell);

  // More free-text cells
  row.appendChild(makeInputBoxes("transactionFrequency", "e.g. 5 s, 1 min", data.transactionFrequency));
  row.appendChild(makeInputBoxes("tagDescription", "What does this tag measure?", data.tagDescription));
  row.appendChild(makeInputBoxes("min", "e.g. 0", data.min));
  row.appendChild(makeInputBoxes("max", "e.g. 100", data.max));

  // "New table?" checkbox
  const checkBoxCell = document.createElement("td");
  checkBoxCell.className = "checkbox-cell";
  const checkBox = document.createElement("input");
  checkBox.type = "checkbox";
  checkBox.className = "newTable";
  if (data.newTable) {
    checkBox.checked = true;
  }
  checkBoxCell.appendChild(checkBox);
  row.appendChild(checkBoxCell);

  // Remove-row button
  const removeCell = document.createElement("td");
  const removeBtn = document.createElement("button");
  removeBtn.className = "remove-btn";
  removeBtn.textContent = "x";
  removeBtn.onclick = function () {
    row.remove();
  };
  removeCell.appendChild(removeBtn);
  row.appendChild(removeCell);

  tableBody.appendChild(row);
  rowsCount();

  // Attach validation on blur.
  const tagInput = row.querySelector(".tagName");
  tagInput.addEventListener("blur", function () {
    valTag(tagInput);
    chkDups();
  });

  const ipInput = row.querySelector(".plcAddress");
  ipInput.addEventListener("blur", function () {
    valIP(ipInput);
  });

  const minInput = row.querySelector(".min");
  minInput.addEventListener("blur", function () {
    valNumber(minInput);
  });

  const maxInput = row.querySelector(".max");
  maxInput.addEventListener("blur", function () {
    valNumber(maxInput);
  });
}

// Build a plain text <input> wrapped in a <td>.
function makeInputBoxes(className, placeholder, value) {
  const cell = document.createElement("td");
  const input = document.createElement("input");
  input.type = "text";
  input.className = "cell-input " + className;
  input.placeholder = placeholder || "";
  input.value = value || "";
  cell.appendChild(input);
  return cell;
}

// Update the row counter and empty-state message.
function rowsCount() {
  const rowCount = document.querySelectorAll("#tbody tr").length;
  const noTagsMessage = document.getElementById("emptyState");
  noTagsMessage.style.display = rowCount === 0 ? "block" : "none";
  document.getElementById("rowCounter").textContent =
    rowCount + (rowCount === 1 ? " row" : " rows");
}


/* ===================================================================
   3. INPUT VALIDATION
   =================================================================== */

// Tag name: warn on underscores.
function valTag(input) {
  clearMsg(input);
  const value = input.value.trim();
  if (!value) return;
  if (value.includes("_")) {
    addWarning(input, "Avoid underscores in the tag name");
  }
}

// PLC IP address: must match the IP pattern; no commas.
function valIP(input) {
  const IP_PATTERN = /^(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)(?::\d{1,5})?$/;
  clearMsg(input);
  const value = input.value.trim();

  if (!value) return;

  if (value.includes(",")) {
    addWarning(input, "Avoid adding , in PLC IP Adresses. Use . instead");
    return;
  }

  if (!IP_PATTERN.test(value)) {
    addWarning(input, "Avoid adding , in PLC IP Adresses. Use . instead");
    return;
  }
}

// Min/Max: must be a number; no commas.
function valNumber(input) {
  clearMsg(input);
  const value = input.value.trim();
  if (!value) return;

  if (value.includes(",")) {
    addWarning(input, "Use . instead of ,");
    return;
  }

  if (!/^-?\d*\.?\d*$/.test(value)) {
    addWarning(input, "Must be a number");
  }
}

// Flag duplicate (tableName | tagName) pairs across all rows.
function chkDups() {
  const rows = document.querySelectorAll("#tbody tr");
  const counts = {};

  rows.forEach(function (row) {
    const tableName = (row.querySelector(".tableName")?.value || "").trim().toLowerCase();
    const tagName = (row.querySelector(".tagName")?.value || "").trim().toLowerCase();
    if (tagName) {
      const key = tableName + "|" + tagName;
      counts[key] = (counts[key] || 0) + 1;
    }
  });

  rows.forEach(function (row) {
    const tagInput = row.querySelector(".tagName");
    const tableName = (row.querySelector(".tableName")?.value || "").trim().toLowerCase();
    const tagName = (row.querySelector(".tagName")?.value || "").trim().toLowerCase();

    clearMsg(tagInput);
    const key = tableName + "|" + tagName;
    if (counts[key] > 1) {
      addWarning(tagInput, "Duplicate tags are not allowed");
    }
  });
}


/* ===================================================================
   4. CLEAR-ALL CONFIRMATION
   =================================================================== */

function clearAll() {
  const clearAllWarning = document.querySelector(".clearAllWarning");
  clearAllWarning.classList.add("show");

  const Warning = document.createElement("div");
  Warning.className = "warning";

  const warningText = document.createElement("div");
  warningText.textContent = "Are you sure you want to clear all rows? This action cannot be undone.";

  const confirmBtn = document.createElement("input");
  confirmBtn.type = "button";
  confirmBtn.value = "Confirm";

  const cancelBtn = document.createElement("input");
  cancelBtn.type = "button";
  cancelBtn.value = "Cancel";

  Warning.appendChild(warningText);
  Warning.appendChild(confirmBtn);
  Warning.appendChild(cancelBtn);
  clearAllWarning.appendChild(Warning);

  confirmBtn.addEventListener("click", function () {
    Warning.remove();
    const rowBody = document.querySelectorAll("#tbody tr");
    rowBody.forEach(row => row.remove());
    clearAllWarning.classList.remove("show");
  });

  cancelBtn.addEventListener("click", function () {
    Warning.remove();
    clearAllWarning.classList.remove("show");
  });
}


/* ===================================================================
   5. SUPPORT HELPERS
   =================================================================== */

// Add a warning message under an input and flag the input.
function addWarning(input, text) {
  const msg = document.createElement("div");
  msg.className = "msg warn";
  msg.textContent = text;
  input.classList.add("warn");
  input.parentElement.appendChild(msg);
}

// Remove any existing warning message under an input.
function clearMsg(input) {
  const oldMsg = input.parentElement.querySelector(".msg");
  if (oldMsg) oldMsg.remove();
}

// Check every "new table" box.
function checkAll() {
  const rows = document.querySelectorAll("#tbody tr");
  rows.forEach(row => {
    row.querySelector(".newTable").checked = true;
  });
}

// Uncheck every "new table" box.
function uncheckAll() {
  const rows = document.querySelectorAll("#tbody tr");
  rows.forEach(row => {
    row.querySelector(".newTable").checked = false;
  });
}


/* ===================================================================
   6. SUBMIT + TABLE SERIALIZATION
   =================================================================== */

// POST the request to the agent endpoint.
async function doSubmit() {
  const mill = document.getElementById("millSelect").value;
  const count = document.querySelectorAll("#tbody tr").length;
  const html = buildTableHTML();
  const subject = mill + " Tag Request: " + count + " tags on " + new Date().toLocaleDateString("en-US");

  try {
    const response = await fetch("http://roc-laa01:3000/submit-tags", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        email_html: html,
        email_subject: subject,
        email_from: "wfm-tagsrus@westfraser.com",
      }),
    });
    const data = await response.json();
    alert("Request submitted !");
  } catch (error) {
    alert("Submit failed: " + error.message);
  }
}

// Serialize all rows into an HTML table string.
function buildTableHTML() {
  const rows = document.querySelectorAll("#tbody tr");

  let html =
    "<table border='1'><thead><tr>" +
    "<th>TableName</th><th>TagName</th><th>PLCIPAddress</th><th>PLCPath</th>" +
    "<th>DataType</th><th>Units</th><th>TransactionFrequency</th>" +
    "<th>TagDescription</th><th>Min</th><th>Max</th><th>NewTable?</th>" +
    "</tr></thead><tbody>";

  rows.forEach(row => {
    html += "<tr>";
    html += "<td>" + row.querySelector(".tableName").value + "</td>";
    html += "<td>" + row.querySelector(".tagName").value + "</td>";
    html += "<td>" + row.querySelector(".plcAddress").value + "</td>";
    html += "<td>" + row.querySelector(".plcPath").value + "</td>";
    html += "<td>" + row.querySelector(".dataType").value + "</td>";
    html += "<td>" + row.querySelector(".units").value + "</td>";
    html += "<td>" + row.querySelector(".transactionFrequency").value + "</td>";
    html += "<td>" + row.querySelector(".tagDescription").value + "</td>";
    html += "<td>" + row.querySelector(".min").value + "</td>";
    html += "<td>" + row.querySelector(".max").value + "</td>";
    html += "<td>" + (row.querySelector(".newTable").checked ? "Yes" : "No") + "</td>";
    html += "</tr>";
  });

  html += "</tbody></table>";
  return html;
}


/* ===================================================================
   7. EXCEL PASTE IMPORT
   =================================================================== */

function processExcelData() {
  const raw = document.getElementById("pasteTextBox").value;
  if (!raw.trim()) return;

  // Vertical format (no tabs): 10 lines per record.
  if (!raw.includes("\t")) {
    const lines = raw.trim().split(/\r?\n/).filter(l => l.trim() !== "");
    for (let i = 0; i + 10 <= lines.length; i += 10) {
      // Skip a header block.
      if (
        lines[i] === "TableName" || lines[i + 1] === "TagName" || lines[i + 2] === "PLC IP Address" ||
        lines[i + 3] === "PLC Path" || lines[i + 4] === "Data Type" || lines[i + 5] === "Units" ||
        lines[i + 6] === "Desired Transaction Frequency" || lines[i + 7] === "Tag Description" ||
        lines[i + 8] === "Min Value" || lines[i + 9] === "Max Value"
      ) {
        continue;
      }
      addRow({
        tableName: lines[i],
        tagName: lines[i + 1],
        plcAddress: lines[i + 2],
        plcPath: lines[i + 3],
        dataType: lines[i + 4],
        units: lines[i + 5],
        transactionFrequency: lines[i + 6],
        tagDescription: lines[i + 7],
        min: lines[i + 8],
        max: lines[i + 9],
      });
    }
    return;
  }

  // Tab-delimited format: one record per line.
  const result = Papa.parse(raw.trim(), { delimiter: "\t" });
  const rows = result.data;

  rows.forEach(row => {
    if (row.every(c => !c || !c.trim())) return;

    // Skip a header row.
    if (
      row[0] === "Table Name" || row[1] === "Tag Name DataParc" || row[2] === "PLC IP Address" ||
      row[3] === "PLC Path/PLC Ctrl Lgx" || row[4] === "Data Type" || row[5] === "Units" ||
      row[6] === "Transaction Frequency" || row[7] === "Tag Description" || row[8] === "Min" ||
      row[9] === "Max" || row[10] === "NewTable?"
    ) {
      return;
    }

    addRow({
      tableName: row[0],
      tagName: row[1],
      plcAddress: row[2],
      plcPath: row[3],
      dataType: row[4],
      units: row[5],
      transactionFrequency: row[6],
      tagDescription: row[7],
      min: row[8],
      max: row[9],
    });
  });

  chkDups();
  document.getElementById("pasteTextBox").value = "";
}

function clearText() {
  const clearTextButton = document.getElementById("pasteTextBox");
  clearTextButton.value = "";
}


/* ===================================================================
   8. ENDPOINT FETCHES
   =================================================================== */

// Units dropdown source.
let validUnits = [];
async function getUnits() {
  try {
    const res = await fetch("http://roc-laa01.westfrasertimber.ca:3000/getUnits");
    validUnits = await res.json();
  } catch (err) {
    console.error(err);
  }
}

// Table names for the selected mill.
let validTables = [];
async function getTableNames() {
  try {
    const mill = document.getElementById("millSelect").value;
    if (!mill) return;

    const select = document.getElementById("millSelect");
    const selectedOption = select.options[select.selectedIndex];

    let grp = 0;
    if (
      selectedOption.parentElement.tagName === "OPTGROUP" &&
      selectedOption.parentElement.label === "Sawmills"
    ) {
      grp = 1;
    }

    const res = await fetch(`http://roc-laa01.westfrasertimber.ca:3000/getTableNames?mill=${mill}&grp=${grp}`);
    validTables = await res.json();
  } catch (err) {
    console.error(err);
  }
}

// Download the Excel template.
function openExcelTemplate() {
  window.location.href = "http://roc-laa01.westfrasertimber.ca:3000/downloadExcelTemplate";
}


/* ===================================================================
   9. SEARCH DROPDOWN WIDGET
   =================================================================== */

// Build a text input with a filterable dropdown list beneath it.
function createSearchDropdown(className, items, selectedValue) {
  const wrap = document.createElement("div");
  wrap.style.position = "relative";

  const input = document.createElement("input");
  input.type = "text";
  input.className = "cell-input " + className;
  input.value = selectedValue || "";
  input.autocomplete = "off";

  const list = document.createElement("div");
  list.className = "search-drop";

  function render(filter) {
    list.innerHTML = "";
    const f = (filter || "").toLowerCase();
    const matches = items.filter(i => i.toLowerCase().includes(f));

    if (!matches.length) {
      list.style.display = "none";
      return;
    }

    matches.forEach(t => {
      const opt = document.createElement("div");
      opt.className = "search-opt";
      opt.textContent = t;
      opt.addEventListener("mousedown", e => {
        e.preventDefault();
        input.value = t;
        list.style.display = "none";
      });
      list.appendChild(opt);
    });

    list.style.display = "block";
  }

  input.addEventListener("focus", () => render(input.value));
  input.addEventListener("input", () => render(input.value));
  input.addEventListener("blur", () => setTimeout(() => (list.style.display = "none"), 150));

  wrap.appendChild(input);
  wrap.appendChild(list);
  return wrap;
}


/* ===================================================================
   TOP-LEVEL TABS  (Home / Tag Request / Workbook)
   =================================================================== */

function showTab(which) {
  const views = { home: "homeView", tag: "formView", workbook: "workbookView" };
  Object.entries(views).forEach(([key, id]) => {
    const el = document.getElementById(id);
    if (el) el.style.display = (key === which) ? "block" : "none";
  });

  // Swap the header title/subtitle to match the active tab.
  const titles = {
    home:     ["Dashboard", "ROC — Tag Requests & Coverage"],
    tag:      ["Tag Request Form", "ROC — OSB & Sawmill Form"],
    workbook: ["Workbook Management", "ROC — Tag Rollout Status"],
  };
  const t = titles[which];
  if (t) {
    const h = document.getElementById("appTitle");
    const s = document.getElementById("appSubtitle");
    if (h) h.textContent = t[0];
    if (s) s.textContent = t[1];
  }

  // Reflect active state on the header buttons
  document.querySelectorAll(".nav-btn").forEach(b => b.classList.remove("active"));
  const btn = document.getElementById("nav-" + which);
  if (btn) btn.classList.add("active");

  if (which === "workbook" && typeof wbInit === "function") wbInit();
  if (which === "home" && typeof dashInit === "function") dashInit();
}

function homeTab()     { showTab("home"); }
function tagTab()      { showTab("tag"); }
function workbookTab() { showTab("workbook"); }


/* ===================================================================
   STARTUP
   =================================================================== */

window.addEventListener("DOMContentLoaded", () => {
  getUnits();
  showTab("tag");   // default view
});