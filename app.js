/* =========================================================
   CSLB Sales Dashboard – app.js
   Static site, loaded after XLSX via CDN in index.html
   ========================================================= */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const WORKBOOK_PATH = "sales-dashboard.xlsx";

// Possible sheet name variants (first match wins)
const SHEET_VARIANTS = {
  region: ["Region", "REGION", "Regions"],
  district: ["District", "DISTRICT", "Districts"],
  store: ["Store", "STORE", "Stores"],
  location: ["Location List", "Locations", "Location"],
  metrics: ["Metric Rankings", "Metrics", "Metric Ranking"],
};

// Metric groups – value/rank list fallbacks so renamed columns still resolve
const METRIC_GROUPS = [
  {
    label: "GP/LH",
    weight: "15%",
    valueFallbacks: ["GP Per Labor Hour %Tgt", "GP Per Labor Hour % Tgt", "GP/LH %Tgt", "GP/LH % Tgt", "GP Per Labor Hour Pct Tgt"],
    rankFallbacks: ["GP Per Labor Hour Rank", "GP/LH Rank", "GPLH Rank"],
    isCsat: false,
    actualFallbacks: ["GP Per Labor Hour Actual", "GP/LH Actual"],
  },
  {
    label: "PP Act",
    weight: "15%",
    valueFallbacks: ["PP Act %Tgt", "PP Act % Tgt", "PP Act % Target", "PP Acts %Tgt", "PP Acts Pct Tgt"],
    rankFallbacks: ["PP Act Attain Rank", "PP Act Rank", "PP Acts Rank"],
    isCsat: false,
    actualFallbacks: ["PP Acts", "PP Act"],
  },
  {
    label: "ReBiz Conv",
    weight: "15%",
    valueFallbacks: ["ReBiz Conv %Tgt", "ReBiz Conv % Tgt", "Rebiz Conv %Tgt", "ReBiz Conv % Target"],
    rankFallbacks: ["ReBiz Conv Rank", "Rebiz Conv Rank"],
    isCsat: false,
    actualFallbacks: ["Rebiz Conv", "ReBiz Conv"],
  },
  {
    label: "Acc GP",
    weight: "10%",
    valueFallbacks: ["Acc GP Pct Actual", "Acc GP % Actual", "Acc GP Pct", "Acc GP %"],
    rankFallbacks: ["Acc GP Pct Rank", "Acc GP Rank", "Acc GP % Rank"],
    isCsat: false,
    actualFallbacks: ["Acc GP Pct Actual", "Acc GP Actual"],
  },
  {
    label: "CSAT",
    weight: "10%",
    valueFallbacks: ["CSAT Actual", "CSAT Score", "CSAT"],
    rankFallbacks: ["CSAT Rank"],
    isCsat: true,
    actualFallbacks: ["CSAT Actual", "CSAT Score"],
  },
  {
    label: "Visa Priority",
    weight: "15%",
    valueFallbacks: ["Visa Priority Rate %Tg", "Visa Priority Rate %Tgt", "Visa Priority Rate % Tgt", "Visa Priority %Tgt"],
    rankFallbacks: ["Visa Priority Rate Rank", "Visa Priority Rank"],
    isCsat: false,
    actualFallbacks: ["Visa Priority Rate", "Visa Priority Apps"],
  },
  {
    label: "P360 Attach",
    weight: "10%",
    valueFallbacks: ["P360 Attach Rate %Tgt", "P360 Attach Rate % Tgt", "P360 %Tgt", "Indexed P360 Attach Rate"],
    rankFallbacks: ["P360 Attach Rate Rank", "P360 Rank"],
    isCsat: false,
    actualFallbacks: ["Indexed P360 Attach Rate", "P360 Attach Rate"],
  },
  {
    label: "Premium Mix",
    weight: "10%",
    valueFallbacks: ["Premium Mix Rate %Tgt", "Premium Mix Rate % Tgt", "Premium Mix %Tgt"],
    rankFallbacks: ["Premium Rate Plan Rank", "Premium Mix Rate Rank", "Premium Mix Rank"],
    isCsat: false,
    actualFallbacks: ["Premium Mix Rate", "Premium Lines"],
  },
];

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------

const state = {
  workbook: null,
  regions: [],
  districts: [],
  stores: [],
  locations: [],
  metricRankingRows: [],
  regionRankTable: [],
  districtRankTable: [],
  storeRankTable: [],
  selectedDistrict: "",
  outlierMetric: null,
  outlierMode: "lowest",
  // resolved column maps per metric group per sheet
  resolvedColumns: { region: [], district: [], store: [] },
};

// ---------------------------------------------------------------------------
// DOM refs (populated after DOMContentLoaded)
// ---------------------------------------------------------------------------

let el = {};

function initEl() {
  const ids = [
    "statusBar", "dashboardView", "outliersView",
    "btnDashboard", "btnOutliers",
    "regionMeta", "regionRankHost", "regionWeightHost", "regionTable",
    "districtMeta", "districtRankHost", "districtWeightHost", "districtTable",
    "districtSelect",
    "selectedDistrictSummary", "storeRankHost", "storeWeightHost", "storeTable",
    "outlierMetricSelect", "outlierModeSelect",
    "outliersMeta", "outliersRankHost", "outliersTable",
  ];
  ids.forEach((id) => {
    el[id] = document.getElementById(id);
  });
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

function norm(s) {
  return String(s ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function resolveKey(obj, candidates) {
  if (!obj) return undefined;
  const keys = Object.keys(obj);

  // 1) Exact case-insensitive match (preserves special chars like % so %Tgt != Tgt)
  for (const c of candidates) {
    const cl = c.toLowerCase();
    const match = keys.find((k) => k.toLowerCase() === cl);
    if (match !== undefined) return match;
  }

  // 2) Normalized exact match (strips punctuation)
  for (const c of candidates) {
    const nc = norm(c);
    const match = keys.find((k) => norm(k) === nc);
    if (match !== undefined) return match;
  }

  // 3) Fuzzy partial (last resort)
  for (const c of candidates) {
    const nc = norm(c);
    const match = keys.find((k) => norm(k).includes(nc) || nc.includes(norm(k)));
    if (match !== undefined) return match;
  }

  return undefined;
}

function getVal(row, candidates) {
  if (!row) return "";
  const key = resolveKey(row, candidates);
  return key !== undefined ? row[key] : "";
}

function fmtNum(v) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") {
    if (Number.isInteger(v)) return String(v);
    return v.toFixed(2);
  }
  return String(v);
}

function fmtPct(v) {
  if (v === null || v === undefined || v === "") return "";
  const n = typeof v === "number" ? v : parseFloat(v);
  if (isNaN(n)) return String(v);
  return (n * 100).toFixed(1) + "%";
}

function quintileClass(qp) {
  const n = parseInt(qp, 10);
  if (n >= 1 && n <= 5) return `q${n}`;
  return "";
}

// ---------------------------------------------------------------------------
// Sheet name resolution
// ---------------------------------------------------------------------------

function findSheet(wb, variants) {
  for (const v of variants) {
    const match = wb.SheetNames.find((s) => norm(s) === norm(v));
    if (match) return wb.Sheets[match];
  }
  // fuzzy
  for (const v of variants) {
    const match = wb.SheetNames.find((s) => norm(s).includes(norm(v)));
    if (match) return wb.Sheets[match];
  }
  return null;
}

// ---------------------------------------------------------------------------
// Column resolution (resolved once after workbook load)
// ---------------------------------------------------------------------------

function buildResolvedColumns(rows) {
  if (!rows || !rows.length) return METRIC_GROUPS.map(() => ({ valueKey: null, rankKey: null }));
  const sample = rows[0];
  return METRIC_GROUPS.map((g) => ({
    valueKey: resolveKey(sample, g.valueFallbacks),
    rankKey: resolveKey(sample, g.rankFallbacks),
  }));
}

// ---------------------------------------------------------------------------
// Load workbook
// ---------------------------------------------------------------------------

function setStatus(msg, isError = false) {
  if (!el.statusBar) return;
  el.statusBar.textContent = msg;
  el.statusBar.style.color = isError ? "#c00" : "";
}

async function loadWorkbook() {
  setStatus("Loading workbook…");
  try {
    const resp = await fetch(WORKBOOK_PATH);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const buf = await resp.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    state.workbook = wb;

    parseSheets(wb);
    populateDistrictSelect();
    populateOutlierMetricSelect();
    renderDashboard();
    setStatus(`Loaded: ${state.regions.length} regions, ${state.districts.length} districts, ${state.stores.length} stores.`);
  } catch (e) {
    setStatus(`Error loading workbook: ${e.message}`, true);
  }
}

// ---------------------------------------------------------------------------
// Parse sheets
// ---------------------------------------------------------------------------

function sheetRows(wb, variants) {
  const ws = findSheet(wb, variants);
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

function parseRankingTables(rows) {
  let section = null;
  const tables = { region: [], district: [], store: [] };
  const sectionMap = {
    "region ranking table": "region",
    "district ranking table": "district",
    "store ranking table": "store",
  };
  for (const row of rows) {
    const first = norm(String(Object.values(row)[0] ?? ""));
    if (sectionMap[first] !== undefined) { section = sectionMap[first]; continue; }
    if (first === "quintile position") continue; // header row
    const vals = Object.values(row);
    const qp = parseInt(vals[0], 10);
    const minR = parseInt(vals[1], 10);
    const maxR = parseInt(vals[2], 10);
    if (!isNaN(qp) && !isNaN(minR) && !isNaN(maxR) && section) {
      tables[section].push([qp, minR, maxR]);
    }
  }
  // fallback defaults
  if (!tables.region.length) tables.region = [[1,1,1],[2,2,3],[3,3,3],[4,4,4]];
  if (!tables.district.length) tables.district = [[1,1,7],[2,8,14],[3,15,22],[4,23,29],[5,30,37]];
  if (!tables.store.length) tables.store = [[1,1,72],[2,73,144],[3,145,216],[4,217,288],[5,289,360]];
  return tables;
}

function parseSheets(wb) {
  state.regions = sheetRows(wb, SHEET_VARIANTS.region);
  state.districts = sheetRows(wb, SHEET_VARIANTS.district);
  state.stores = sheetRows(wb, SHEET_VARIANTS.store);
  state.locations = sheetRows(wb, SHEET_VARIANTS.location);
  state.metricRankingRows = sheetRows(wb, SHEET_VARIANTS.metrics);

  const tables = parseRankingTables(state.metricRankingRows);
  state.regionRankTable = tables.region;
  state.districtRankTable = tables.district;
  state.storeRankTable = tables.store;

  state.resolvedColumns.region = buildResolvedColumns(state.regions);
  state.resolvedColumns.district = buildResolvedColumns(state.districts);
  state.resolvedColumns.store = buildResolvedColumns(state.stores);

  // Default selected district
  if (state.districts.length) {
    state.selectedDistrict = getVal(state.districts[0], ["District", "District Name"]);
  }
}

// ---------------------------------------------------------------------------
// Quintile from rank
// ---------------------------------------------------------------------------

function quintileFromRank(rank, rankTable) {
  const r = parseInt(rank, 10);
  if (isNaN(r)) return "";
  for (const [qp, min, max] of rankTable) {
    if (r >= min && r <= max) return String(qp);
  }
  return "";
}

// ---------------------------------------------------------------------------
// Populate selects
// ---------------------------------------------------------------------------

function populateDistrictSelect() {
  if (!el.districtSelect) return;
  el.districtSelect.innerHTML = "";
  state.districts.forEach((row) => {
    const name = getVal(row, ["District", "District Name"]);
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    el.districtSelect.appendChild(opt);
  });
  if (state.selectedDistrict) el.districtSelect.value = state.selectedDistrict;
}

function populateOutlierMetricSelect() {
  if (!el.outlierMetricSelect) return;
  el.outlierMetricSelect.innerHTML = "";
  METRIC_GROUPS.forEach((g, i) => {
    const opt = document.createElement("option");
    opt.value = String(i);
    opt.textContent = g.label;
    el.outlierMetricSelect.appendChild(opt);
  });
  state.outlierMetric = 0;
}

// ---------------------------------------------------------------------------
// Render helpers: rank table + weight band
// ---------------------------------------------------------------------------

function renderRankTable(container, rows) {
  if (!container) return;
  container.innerHTML = "";
  const tbl = document.createElement("table");
  tbl.className = "rank-table";
  const thead = document.createElement("thead");
  const hrow = document.createElement("tr");
  ["Quintile Position", "Min Rank", "Max Rank"].forEach((h) => {
    const th = document.createElement("th");
    th.textContent = h;
    hrow.appendChild(th);
  });
  thead.appendChild(hrow);
  tbl.appendChild(thead);
  const tbody = document.createElement("tbody");
  rows.forEach(([qp, min, max]) => {
    const tr = document.createElement("tr");
    [qp, min, max].forEach((v, i) => {
      const td = document.createElement("td");
      td.textContent = v;
      if (i === 0) td.classList.add(quintileClass(v));
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  tbl.appendChild(tbody);
  container.appendChild(tbl);
}

function renderWeightBand(container) {
  if (!container) return;
  container.innerHTML = "";
  const band = document.createElement("div");
  band.className = "weight-band";
  const spacer = document.createElement("div");
  spacer.className = "weight-spacer";
  band.appendChild(spacer);
  const grid = document.createElement("div");
  grid.className = "weight-grid";
  METRIC_GROUPS.forEach((g) => {
    const cell = document.createElement("div");
    cell.className = "weight-cell";
    cell.textContent = g.weight;
    grid.appendChild(cell);
  });
  band.appendChild(grid);
  container.appendChild(band);
}

// ---------------------------------------------------------------------------
// Render grouped table
// ---------------------------------------------------------------------------

function buildGroupedTable(tableEl, rows, labelFallbacks, resolvedCols, rankTable) {
  if (!tableEl) return;
  tableEl.innerHTML = "";

  if (!rows || !rows.length) {
    tableEl.innerHTML = "<thead><tr><th>No data</th></tr></thead><tbody><tr><td>Nothing to display.</td></tr></tbody>";
    return;
  }

  const thead = document.createElement("thead");

  // Row 1: base headers (rowspan=2) + metric group headers (colspan=2)
  const groupRow = document.createElement("tr");

  const baseHeaders = [
    { label: "Month", fallbacks: ["Month", "MONTH", "Month Name"] },
    { label: "Name", fallbacks: labelFallbacks },
    { label: "Quintile Position", fallbacks: ["Quintile Position", "Quintile", "Q"] },
    { label: "Overall Rank", fallbacks: ["Overall Rank", "OverallRank", "Rank"] },
    { label: "*", fallbacks: ["*"] },
  ];

  baseHeaders.forEach((h) => {
    const th = document.createElement("th");
    th.rowSpan = 2;
    th.textContent = h.label;
    groupRow.appendChild(th);
  });

  METRIC_GROUPS.forEach((g) => {
    const th = document.createElement("th");
    th.colSpan = 2;
    th.textContent = `${g.label} ${g.weight}`;
    groupRow.appendChild(th);
  });

  // Row 2: metric sub-headers (value + rank)
  const subRow = document.createElement("tr");
  METRIC_GROUPS.forEach((g) => {
    const thV = document.createElement("th");
    thV.textContent = g.label;
    subRow.appendChild(thV);
    const thR = document.createElement("th");
    thR.textContent = "Rank";
    subRow.appendChild(thR);
  });

  thead.appendChild(groupRow);
  thead.appendChild(subRow);
  tableEl.appendChild(thead);

  const tbody = document.createElement("tbody");

  rows.forEach((row) => {
    const tr = document.createElement("tr");

    const qpRaw = getVal(row, ["Quintile Position", "Quintile", "Q"]);
    const qpStr = String(qpRaw);
    const qCls = quintileClass(qpStr);

    baseHeaders.forEach((h) => {
      const td = document.createElement("td");
      if (h.label === "*") {
        td.textContent = "*";
      } else {
        const v = getVal(row, h.fallbacks);
        td.textContent = fmtNum(v);
      }
      if (h.label === "Quintile Position" || h.label === "Overall Rank") {
        if (qCls) td.classList.add(qCls);
      }
      tr.appendChild(td);
    });

    METRIC_GROUPS.forEach((g, i) => {
      const col = resolvedCols[i] || {};
      const vKey = col.valueKey;
      const rKey = col.rankKey;

      const valueTd = document.createElement("td");
      const rawVal = vKey ? row[vKey] : "";
      if (g.isCsat) {
        valueTd.textContent = fmtNum(rawVal);
      } else {
        valueTd.textContent = fmtPct(rawVal);
      }
      tr.appendChild(valueTd);

      const rankTd = document.createElement("td");
      const rawRank = rKey ? row[rKey] : "";
      rankTd.textContent = fmtNum(rawRank);
      if (rawRank !== "" && rankTable) {
        const qp = quintileFromRank(rawRank, rankTable);
        const cls = quintileClass(qp);
        if (cls) rankTd.classList.add(cls);
      }
      tr.appendChild(rankTd);
    });

    tbody.appendChild(tr);
  });

  tableEl.appendChild(tbody);
}

// ---------------------------------------------------------------------------
// Summary grid (district drilldown header)
// ---------------------------------------------------------------------------

function renderSummaryGrid(container, row) {
  if (!container) return;
  container.innerHTML = "";
  if (!row) return;

  const grid = document.createElement("div");
  grid.className = "summary-grid";

  const fields = [
    { label: "District", fallbacks: ["District", "District Name"] },
    { label: "Month", fallbacks: ["Month", "MONTH"] },
    { label: "Quintile Position", fallbacks: ["Quintile Position", "Quintile"] },
    { label: "Overall Rank", fallbacks: ["Overall Rank", "Overall"] },
    { label: "Overall Score", fallbacks: ["Overall Score"] },
  ];

  fields.forEach((f) => {
    const v = getVal(row, f.fallbacks);
    if (v === "" && f.label === "Overall Score") return;
    const div = document.createElement("div");
    div.className = "summary-item";
    const label = document.createElement("span");
    label.className = "summary-label";
    label.textContent = f.label;
    const val = document.createElement("span");
    val.className = "summary-value";
    val.textContent = fmtNum(v);
    const qp = getVal(row, ["Quintile Position", "Quintile"]);
    if (f.label === "Quintile Position" || f.label === "Overall Rank") {
      const cls = quintileClass(String(qp));
      if (cls) div.classList.add(cls);
    }
    div.appendChild(label);
    div.appendChild(val);
    grid.appendChild(div);
  });

  container.appendChild(grid);
}

// ---------------------------------------------------------------------------
// Outlier rank table
// ---------------------------------------------------------------------------

function renderOutlierRankTable(container, rankTable) {
  if (!container) return;
  container.innerHTML = "";
  renderRankTable(container, rankTable);
}

// ---------------------------------------------------------------------------
// Outlier table
// ---------------------------------------------------------------------------

function renderOutliersTable() {
  const metricIdx = parseInt(el.outlierMetricSelect?.value ?? "0", 10);
  const mode = el.outlierModeSelect?.value ?? "lowest";
  const g = METRIC_GROUPS[metricIdx];
  if (!g) return;

  // Update meta label
  if (el.outliersMeta) {
    el.outliersMeta.textContent = `Metric: ${g.label} | View: ${mode === "lowest" ? "Lowest 15" : "Highest 15"}`;
  }

  // Resolve the value column from store rows
  const colInfo = state.resolvedColumns.store[metricIdx] || {};
  const valueKey = colInfo.valueKey;

  // Dynamic second column header
  const valueColHeader = g.isCsat ? "CSAT Actual" : "Percent to Target";

  // Render rank-range table
  renderOutlierRankTable(el.outliersRankHost, state.storeRankTable);

  // Sort and slice
  let rows = state.stores.filter((r) => {
    const v = valueKey ? r[valueKey] : "";
    return v !== "" && v !== null && v !== undefined;
  });

  rows.sort((a, b) => {
    const va = valueKey ? parseFloat(a[valueKey]) : 0;
    const vb = valueKey ? parseFloat(b[valueKey]) : 0;
    return mode === "lowest" ? va - vb : vb - va;
  });

  rows = rows.slice(0, 15);

  // Build table
  if (!el.outliersTable) return;
  el.outliersTable.innerHTML = "";

  const thead = document.createElement("thead");
  const hrow = document.createElement("tr");
  ["Store Name", valueColHeader].forEach((h) => {
    const th = document.createElement("th");
    th.textContent = h;
    hrow.appendChild(th);
  });
  thead.appendChild(hrow);
  el.outliersTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.forEach((row) => {
    const tr = document.createElement("tr");

    // Store name
    const nameTd = document.createElement("td");
    nameTd.textContent = getVal(row, [
      "Sap: Loaction", "Sap: Location", "SAP Location",
      "Store Name", "STORE NAME", "Store", "SAP", "STORE CODE",
    ]);
    tr.appendChild(nameTd);

    // Value
    const valTd = document.createElement("td");
    const rawVal = valueKey ? row[valueKey] : "";
    valTd.textContent = g.isCsat ? fmtNum(rawVal) : fmtPct(rawVal);
    tr.appendChild(valTd);

    tbody.appendChild(tr);
  });

  el.outliersTable.appendChild(tbody);
}

// ---------------------------------------------------------------------------
// Main render
// ---------------------------------------------------------------------------

function renderDashboard() {
  if (!el.regionMeta) return;
  el.regionMeta.textContent = `${state.regions.length} rows`;
  if (el.districtMeta) el.districtMeta.textContent = `${state.districts.length} rows`;

  renderRankTable(el.regionRankHost, state.regionRankTable);
  renderRankTable(el.districtRankHost, state.districtRankTable);

  renderWeightBand(el.regionWeightHost);
  renderWeightBand(el.districtWeightHost);

  buildGroupedTable(
    el.regionTable, state.regions,
    ["Region", "Region Name", "REGION"],
    state.resolvedColumns.region,
    state.regionRankTable
  );

  buildGroupedTable(
    el.districtTable, state.districts,
    ["District", "District Name", "DISTRICT"],
    state.resolvedColumns.district,
    state.districtRankTable
  );

  renderDistrictView();
}

function renderDistrictView() {
  const districtName = state.selectedDistrict;
  const districtRow = state.districts.find((r) => {
    const d = getVal(r, ["District", "District Name"]);
    return norm(d) === norm(districtName);
  }) || null;

  const storeRows = state.stores.filter((r) => {
    const d = getVal(r, ["District", "District Name", "DISTRICT"]);
    return norm(d) === norm(districtName);
  });

  renderSummaryGrid(el.selectedDistrictSummary, districtRow);

  renderRankTable(el.storeRankHost, state.storeRankTable);
  renderWeightBand(el.storeWeightHost);

  buildGroupedTable(
    el.storeTable, storeRows,
    ["Sap: Loaction", "Sap: Location", "SAP Location", "Store Name", "STORE NAME", "Store", "SAP"],
    state.resolvedColumns.store,
    state.storeRankTable
  );
}

// ---------------------------------------------------------------------------
// View switching
// ---------------------------------------------------------------------------

function showView(viewId) {
  ["dashboardView", "outliersView"].forEach((id) => {
    const v = el[id];
    if (v) v.classList.toggle("hidden", id !== viewId);
  });
}

// ---------------------------------------------------------------------------
// Event wiring
// ---------------------------------------------------------------------------

function wireEvents() {
  el.btnDashboard?.addEventListener("click", () => showView("dashboardView"));

  el.btnOutliers?.addEventListener("click", () => {
    showView("outliersView");
    renderOutliersTable();
  });

  el.districtSelect?.addEventListener("change", (e) => {
    state.selectedDistrict = e.target.value;
    renderDistrictView();
  });

  el.outlierMetricSelect?.addEventListener("change", () => renderOutliersTable());
  el.outlierModeSelect?.addEventListener("change", () => renderOutliersTable());
}

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------

document.addEventListener("DOMContentLoaded", () => {
  initEl();
  wireEvents();
  loadWorkbook();
});
