/* global XLSX */

const WORKBOOK_URL = "sales-dashboard.xlsx";

const QUINTILE_CLASSES = {
  1: "q1",
  2: "q2",
  3: "q3",
  4: "q4",
  5: "q5",
};

const SHEET_CANDIDATES = {
  region: ["Region Sheet", "Region", "Region Data"],
  district: ["District Sheet", "District", "District Data"],
  store: ["Store Sheet", "Store", "Store Data"],
  locations: ["Location List Sheet", "Location List", "Locations"],
  metricRules: ["Metric Rankings Sheet", "Metric Rankings", "Rankings"],
};

const METRIC_CONFIGS = [
  {
    label: "GP Per Labor Hour",
    aliases: ["GP Per Labor Hour"],
    percentKeys: ["GP Per Labor Hour %Tgt", "GP Per Labor Hour % Tgt"],
    rankKeys: ["GP Per Labor Hour Rank"],
  },
  {
    label: "PP Act",
    aliases: ["PP Act", "PP Acts"],
    percentKeys: ["PP Act %Tgt", "PP Act % Tgt"],
    rankKeys: ["PP Act Attain Rank"],
  },
  {
    label: "ReBiz Conv",
    aliases: ["ReBiz Conv"],
    percentKeys: ["ReBiz Conv %Tgt", "ReBiz Conv % Tgt"],
    rankKeys: ["ReBiz Conv Rank"],
  },
  {
    label: "Acc GP",
    aliases: ["Acc GP", "Acc GP Pct"],
    percentKeys: ["Acc GP Pct Actual", "Acc GP Pct"],
    rankKeys: ["Acc GP Pct Rank"],
  },
  {
    label: "CSAT",
    aliases: ["CSAT"],
    percentKeys: ["CSAT Actual", "CSAT"],
    rankKeys: ["CSAT Rank"],
  },
  {
    label: "Visa Priority",
    aliases: ["Visa Priority"],
    percentKeys: ["Visa Priority Rate %Tg", "Visa Priority Rate %Tgt", "Visa Priority Rate"],
    rankKeys: ["Visa Priority Rate Rank"],
  },
  {
    label: "P360 Attach",
    aliases: ["P360 Attach", "Indexed P360 Attach Rate"],
    percentKeys: ["P360 Attach Rate %Tgt", "P360 Attach Rate % Tgt"],
    rankKeys: ["P360 Attach Rate Rank"],
  },
  {
    label: "Premium Mix",
    aliases: ["Premium Mix"],
    percentKeys: ["Premium Mix Rate %Tgt", "Premium Mix Rate % Tgt"],
    rankKeys: ["Premium Rate Plan Rank"],
  },
];

const state = {
  workbook: null,
  regions: [],
  districts: [],
  stores: [],
  locations: [],
  metricRules: {
    region: [],
    district: [],
    store: [],
    weights: [],
  },
  selectedDistrict: null,
  selectedOutlierMetric: null,
  currentView: "dashboard",
};

const el = {
  statusBar: document.getElementById("statusBar"),
  dashboardView: document.getElementById("dashboardView"),
  outliersView: document.getElementById("outliersView"),
  btnDashboard: document.getElementById("btnDashboard"),
  btnOutliers: document.getElementById("btnOutliers"),
  regionTable: document.getElementById("regionTable"),
  districtTable: document.getElementById("districtTable"),
  storeTable: document.getElementById("storeTable"),
  outliersTable: document.getElementById("outliersTable"),
  districtSelect: document.getElementById("districtSelect"),
  outlierMetricSelect: document.getElementById("outlierMetricSelect"),
  outlierModeSelect: document.getElementById("outlierModeSelect"),
  regionMeta: document.getElementById("regionMeta"),
  districtMeta: document.getElementById("districtMeta"),
  selectedDistrictSummary: document.getElementById("selectedDistrictSummary"),
  outliersMeta: document.getElementById("outliersMeta"),
};

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function normalizeKey(key) {
  return String(key || "").replace(/\s+/g, " ").trim();
}

function normalizeText(value) {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function normalizeLookup(value) {
  return normalizeText(value).toLowerCase();
}

function parseNumeric(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return value;
  const raw = String(value).trim();
  if (!raw) return null;

  const negative = /^\(.*\)$/.test(raw);
  const stripped = raw.replace(/[(),$%]/g, "").replace(/\s+/g, "");
  if (!stripped) return null;

  const num = Number(stripped);
  if (Number.isNaN(num)) return null;
  return negative ? -num : num;
}

function parsePercentText(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number") return `${value.toFixed(2)}%`;
  return normalizeText(value);
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return rows.map((row) => {
    const normalized = {};
    Object.entries(row).forEach(([key, value]) => {
      normalized[normalizeKey(key)] = typeof value === "string" ? normalizeText(value) : value;
    });
    return normalized;
  });
}

function getSheetWithFallback(workbook, candidates, includeContains = true) {
  const names = workbook?.SheetNames || [];
  for (const candidate of candidates) {
    const exact = names.find((sheetName) => normalizeLookup(sheetName) === normalizeLookup(candidate));
    if (exact) return workbook.Sheets[exact];
  }
  if (!includeContains) return null;
  for (const candidate of candidates) {
    const found = names.find((sheetName) => normalizeLookup(sheetName).includes(normalizeLookup(candidate)));
    if (found) return workbook.Sheets[found];
  }
  return null;
}

function findColumn(row, possibleNames) {
  const keys = Object.keys(row || {});
  return (
    keys.find((k) => possibleNames.some((candidate) => normalizeLookup(k) === normalizeLookup(candidate))) ||
    keys.find((k) => possibleNames.some((candidate) => normalizeLookup(k).includes(normalizeLookup(candidate))))
  );
}

function determineQuintileFromRank(rankValue, rules) {
  const rank = parseNumeric(rankValue);
  if (rank === null || !Array.isArray(rules)) return "";
  for (const rule of rules) {
    const min = parseNumeric(rule["Min Rank"]);
    const max = parseNumeric(rule["Max Rank"]);
    const q = rule["Quintile Position"];
    if (min === null || max === null || q === undefined || q === "") continue;
    if (rank >= min && rank <= max) return String(q);
  }
  return "";
}

function quintileClass(q) {
  return QUINTILE_CLASSES[String(q).trim()] || "";
}

function isSeparatorColumn(header) {
  return normalizeLookup(header) === "*";
}

function cleanRows(rows) {
  return rows.map((row) => {
    const out = {};
    Object.entries(row).forEach(([key, value]) => {
      if (isSeparatorColumn(key)) return;
      out[key] = value;
    });
    return out;
  });
}

function getPrimaryLabel(row, type) {
  if (!row) return "";
  if (type === "region") return normalizeText(row.Region || row["REGION"]);
  if (type === "district") return normalizeText(row.District || row["DISTRICT"]);
  if (type === "store") {
    return normalizeText(
      row["Sap: Loaction"] ||
        row["Sap: Location"] ||
        row["STORE NAME"] ||
        row.SAP ||
        row["STORE CODE"]
    );
  }
  return "";
}

function detectMetricColumns(row) {
  const headers = Object.keys(row || {});
  const metrics = [];

  for (const cfg of METRIC_CONFIGS) {
    const percentColumn = headers.find((header) =>
      cfg.percentKeys.some((candidate) => normalizeLookup(header) === normalizeLookup(candidate))
    );
    const rankColumn = headers.find((header) =>
      cfg.rankKeys.some((candidate) => normalizeLookup(header) === normalizeLookup(candidate))
    );

    if (percentColumn || rankColumn) {
      metrics.push({
        label: cfg.label,
        percentColumn: percentColumn || "",
        rankColumn: rankColumn || "",
      });
    }
  }

  return metrics;
}

function parseMetricRules(sheetRows) {
  const rules = {
    region: [],
    district: [],
    store: [],
    weights: [],
  };

  let section = "weights";

  for (const row of sheetRows) {
    const values = Object.values(row).map((v) => normalizeText(v)).filter(Boolean);
    if (!values.length) continue;

    const first = normalizeLookup(values[0]);

    if (first === "region ranking table") {
      section = "region";
      continue;
    }
    if (first === "district ranking table") {
      section = "district";
      continue;
    }
    if (first === "store ranking table") {
      section = "store";
      continue;
    }
    if (first === "metric" && values.length > 1) {
      section = "weights";
      continue;
    }

    if (section === "weights" && row.Metric) {
      rules.weights.push({ Metric: row.Metric, Weighting: row.Weighting });
    } else if (section === "region" && row["Quintile Position"]) {
      rules.region.push(row);
    } else if (section === "district" && row["Quintile Position"]) {
      rules.district.push(row);
    } else if (section === "store" && row["Quintile Position"]) {
      rules.store.push(row);
    }
  }

  return rules;
}

function enrichRows(rows, type, quintileRules) {
  return rows.map((row) => {
    const copy = { ...row };
    copy.__type = type;
    copy.__label = getPrimaryLabel(copy, type);

    if (!copy["Quintile Position"]) {
      const rankKey = findColumn(copy, ["Overall Rank"]);
      copy["Quintile Position"] = determineQuintileFromRank(rankKey ? copy[rankKey] : null, quintileRules);
    }

    copy.__quintile = String(copy["Quintile Position"] || "").trim();
    return copy;
  });
}

function sortByLabel(rows, type) {
  return rows.sort((a, b) => {
    const aLabel = getPrimaryLabel(a, type).toLowerCase();
    const bLabel = getPrimaryLabel(b, type).toLowerCase();
    return aLabel.localeCompare(bLabel);
  });
}

function findDistrictRow(districtName) {
  const target = normalizeLookup(districtName);
  return state.districts.find((row) => normalizeLookup(row.District || row.DISTRICT) === target) || null;
}

function getDistrictStores(districtName) {
  const target = normalizeLookup(districtName);
  return state.stores.filter((row) => normalizeLookup(row.District || row.DISTRICT) === target);
}

function formatCellValue(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") {
    if (Number.isInteger(value)) return String(value);
    return value.toFixed(2);
  }
  return String(value);
}

function buildTable(tableEl, rows, type, rankRules = []) {
  tableEl.innerHTML = "";

  if (!rows || !rows.length) {
    tableEl.innerHTML = `<thead><tr><th>No data found</th></tr></thead><tbody><tr><td>Nothing to display.</td></tr></tbody>`;
    return;
  }

  const headers = Object.keys(rows[0]).filter((h) => !isSeparatorColumn(h) && !String(h).startsWith("__"));
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");

  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);
  tableEl.appendChild(thead);

  const tbody = document.createElement("tbody");

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    headers.forEach((header) => {
      const td = document.createElement("td");
      const value = row[header];
      td.textContent = formatCellValue(value);

      const isMetricLike = /%Tgt|%Tg|Rank$|Actual$|Tgt$|Rate$|Score$/i.test(header);
      const q = row.__quintile || determineQuintileFromRank(row["Overall Rank"], rankRules);
      const cls = quintileClass(q);

      if (isMetricLike && cls) td.classList.add("metric-cell", cls);
      if (header === "Quintile Position" && cls) td.classList.add("metric-cell", cls);

      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  tableEl.appendChild(tbody);
}

function renderSummaryGrid(container, row) {
  if (!row) {
    container.innerHTML = `<div class="summary-grid"><div class="summary-item"><span class="summary-label">No selection</span><span class="summary-value">No data available</span></div></div>`;
    return;
  }

  const priorityKeys = ["Month", "Region", "District", "SAP", "Sap: Loaction", "Quintile Position", "Overall Rank", "Overall Score"];
  const keys = Object.keys(row).filter((k) => !k.startsWith("__") && !isSeparatorColumn(k));

  const ordered = [
    ...priorityKeys.filter((k) => keys.includes(k)),
    ...keys.filter((k) => !priorityKeys.includes(k)).slice(0, 10),
  ];

  container.innerHTML = `
    <div class="summary-grid">
      ${ordered
        .map(
          (key) => `
            <div class="summary-item">
              <span class="summary-label">${escapeHtml(key)}</span>
              <span class="summary-value">${escapeHtml(formatCellValue(row[key]))}</span>
            </div>`
        )
        .join("")}
    </div>
  `;
}

function renderDashboard() {
  el.regionMeta.textContent = `${state.regions.length} rows`;
  el.districtMeta.textContent = `${state.districts.length} rows`;

  buildTable(el.regionTable, state.regions, "region", state.metricRules.region);
  buildTable(el.districtTable, state.districts, "district", state.metricRules.district);
  renderDistrictView();
}

function renderDistrictView() {
  const districtName = state.selectedDistrict;
  const districtRow = findDistrictRow(districtName);
  const storeRows = getDistrictStores(districtName);

  renderSummaryGrid(el.selectedDistrictSummary, districtRow);
  buildTable(el.storeTable, storeRows, "store", state.metricRules.store);
}

function renderOutliers() {
  const metricColumn = state.selectedOutlierMetric;
  const metricLabel = el.outlierMetricSelect.options[el.outlierMetricSelect.selectedIndex]?.text || metricColumn || "";
  const mode = el.outlierModeSelect?.value || "lowest";
  const sortDirection = mode === "highest" ? -1 : 1;

  const rows = state.stores
    .map((row) => {
      const storeName = normalizeText(
        row["Sap: Loaction"] || row["Sap: Location"] || row["STORE NAME"] || row["STORE CODE"] || row.SAP
      );

      const metricValue = parseNumeric(row[metricColumn]);

      return {
        District: normalizeText(row.District || row.DISTRICT),
        Store: storeName,
        [metricLabel]: metricValue,
        "% to target": parsePercentText(row[metricColumn]),
        __metricValue: metricValue,
      };
    })
    .filter((row) => row.Store && row.__metricValue !== null)
    .sort((a, b) => {
      if (a.__metricValue === b.__metricValue) return a.Store.toLowerCase().localeCompare(b.Store.toLowerCase());
      return (a.__metricValue - b.__metricValue) * sortDirection;
    })
    .slice(0, 15)
    .map(({ __metricValue, ...rest }) => rest);

  el.outliersMeta.innerHTML = `
    <div class="summary-grid">
      <div class="summary-item">
        <span class="summary-label">Metric</span>
        <span class="summary-value">${escapeHtml(metricLabel)}</span>
      </div>
      <div class="summary-item">
        <span class="summary-label">View</span>
        <span class="summary-value">${escapeHtml(mode === "highest" ? "Highest 15" : "Lowest 15")}</span>
      </div>
      <div class="summary-item">
        <span class="summary-label">Stores shown</span>
        <span class="summary-value">${rows.length}</span>
      </div>
    </div>
  `;

  buildTable(el.outliersTable, rows, "store");
}

function populateDistrictSelector() {
  const districts = [...new Set(state.locations.map((r) => normalizeText(r.DISTRICT || r.District)).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b)
  );

  el.districtSelect.innerHTML = districts.map((district) => `<option value="${escapeHtml(district)}">${escapeHtml(district)}</option>`).join("");
  state.selectedDistrict = districts[0] || null;
  if (state.selectedDistrict) el.districtSelect.value = state.selectedDistrict;
}

function populateOutlierMetrics() {
  const row = state.stores[0] || {};
  const metrics = detectMetricColumns(row);

  el.outlierMetricSelect.innerHTML = metrics
    .map((metric) => `<option value="${escapeHtml(metric.percentColumn)}">${escapeHtml(metric.label)}</option>`)
    .join("");

  state.selectedOutlierMetric = metrics[0]?.percentColumn || null;
  if (state.selectedOutlierMetric) el.outlierMetricSelect.value = state.selectedOutlierMetric;
}

function switchView(view) {
  state.currentView = view;
  el.dashboardView.classList.toggle("hidden", view !== "dashboard");
  el.outliersView.classList.toggle("hidden", view !== "outliers");
}

function attachEvents() {
  el.btnDashboard.addEventListener("click", () => switchView("dashboard"));
  el.btnOutliers.addEventListener("click", () => switchView("outliers"));

  el.districtSelect.addEventListener("change", (e) => {
    state.selectedDistrict = e.target.value;
    renderDistrictView();
  });

  el.outlierMetricSelect.addEventListener("change", (e) => {
    state.selectedOutlierMetric = e.target.value;
    renderOutliers();
  });

  el.outlierModeSelect.addEventListener("change", () => {
    renderOutliers();
  });
}

function parseWorkbook() {
  const regionSheet = getSheetWithFallback(state.workbook, SHEET_CANDIDATES.region);
  const districtSheet = getSheetWithFallback(state.workbook, SHEET_CANDIDATES.district);
  const storeSheet = getSheetWithFallback(state.workbook, SHEET_CANDIDATES.store);
  const locationSheet = getSheetWithFallback(state.workbook, SHEET_CANDIDATES.locations);
  const metricSheet = getSheetWithFallback(state.workbook, SHEET_CANDIDATES.metricRules);

  const regionRows = cleanRows(sheetToObjects(regionSheet));
  const districtRows = cleanRows(sheetToObjects(districtSheet));
  const storeRows = cleanRows(sheetToObjects(storeSheet));
  const locationRows = cleanRows(sheetToObjects(locationSheet));
  const metricRows = cleanRows(sheetToObjects(metricSheet));

  state.metricRules = parseMetricRules(metricRows);
  state.locations = locationRows;
  state.regions = sortByLabel(enrichRows(regionRows, "region", state.metricRules.region), "region");
  state.districts = sortByLabel(enrichRows(districtRows, "district", state.metricRules.district), "district");
  state.stores = sortByLabel(enrichRows(storeRows, "store", state.metricRules.store), "store");
}

async function loadWorkbook() {
  try {
    const response = await fetch(WORKBOOK_URL);
    if (!response.ok) throw new Error(`Could not fetch ${WORKBOOK_URL} (${response.status})`);
    const buffer = await response.arrayBuffer();
    state.workbook = XLSX.read(buffer, { type: "array" });

    parseWorkbook();
    populateDistrictSelector();
    populateOutlierMetrics();

    renderDashboard();
    renderOutliers();
    switchView("dashboard");

    el.statusBar.innerHTML = `Loaded <strong>${escapeHtml(WORKBOOK_URL)}</strong> successfully.`;
  } catch (err) {
    console.error(err);
    el.statusBar.innerHTML = `
      <strong>Unable to load workbook.</strong>
      Please make sure <code>${escapeHtml(WORKBOOK_URL)}</code> exists in the repository root and that the sheet names are correct.
      <br /><br />
      ${escapeHtml(err.message || String(err))}
    `;
  }
}

attachEvents();
loadWorkbook();
