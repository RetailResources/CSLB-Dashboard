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

const METRIC_GROUPS = [
  {
    label: "GP Per Labor Hour",
    value: "GP Per Labor Hour %Tgt",
    rank: "GP Per Labor Hour Rank",
    weight: "15%",
  },
  {
    label: "PP Act",
    value: "PP Act %Tgt",
    rank: "PP Act Attain Rank",
    weight: "15%",
  },
  {
    label: "ReBiz Conv",
    value: "ReBiz Conv %Tgt",
    rank: "ReBiz Conv Rank",
    weight: "15%",
  },
  {
    label: "Acc GP Pct",
    value: "Acc GP Pct Actual",
    rank: "Acc GP Pct Rank",
    weight: "10%",
  },
  {
    label: "CSAT",
    value: "CSAT Actual",
    rank: "CSAT Rank",
    weight: "10%",
  },
  {
    label: "Visa Priority",
    value: "Visa Priority Rate %Tg",
    rank: "Visa Priority Rate Rank",
    weight: "15%",
  },
  {
    label: "P360 Attach",
    value: "P360 Attach Rate %Tgt",
    rank: "P360 Attach Rate Rank",
    weight: "10%",
  },
  {
    label: "Premium Mix",
    value: "Premium Mix Rate %Tgt",
    rank: "Premium Rate Plan Rank",
    weight: "10%",
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
  outliersRankHost: document.getElementById("outliersRankHost"),
  regionMeta: document.getElementById("regionMeta"),
  districtMeta: document.getElementById("districtMeta"),
  selectedDistrictSummary: document.getElementById("selectedDistrictSummary"),
  outliersMeta: document.getElementById("outliersMeta"),
  regionRankHost: document.getElementById("regionRankHost"),
  regionWeightHost: document.getElementById("regionWeightHost"),
  districtRankHost: document.getElementById("districtRankHost"),
  districtWeightHost: document.getElementById("districtWeightHost"),
  storeRankHost: document.getElementById("storeRankHost"),
  storeWeightHost: document.getElementById("storeWeightHost"),
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

  for (const cfg of METRIC_GROUPS) {
    const percentColumn = headers.find((header) =>
      cfg.value === header || normalizeLookup(header) === normalizeLookup(cfg.value)
    );
    const rankColumn = headers.find((header) =>
      cfg.rank === header || normalizeLookup(header) === normalizeLookup(cfg.rank)
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

function buildTable(tableEl, rows, type, rankRules = [], visibleColumns = null) {
  tableEl.innerHTML = "";

  if (!rows || !rows.length) {
    tableEl.innerHTML = `<thead><tr><th>No data found</th></tr></thead><tbody><tr><td>Nothing to display.</td></tr></tbody>`;
    return;
  }

  const headers = (visibleColumns && visibleColumns.length)
    ? visibleColumns.filter((h) => Object.prototype.hasOwnProperty.call(rows[0], h))
    : Object.keys(rows[0]).filter((h) => !isSeparatorColumn(h) && !String(h).startsWith("__"));

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

function renderRankTable(container, rows) {
  container.innerHTML = `
    <table class="data-table rank-table">
      <thead>
        <tr>
          <th>Quintile Position</th>
          <th>Min Rank</th>
          <th>Max Rank</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (r) => `
            <tr>
              <td class="${quintileClass(r[0])}">${r[0]}</td>
              <td class="${quintileClass(r[0])}">${r[1]}</td>
              <td class="${quintileClass(r[0])}">${r[2]}</td>
            </tr>`
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function renderWeightBand(container) {
  const weights = ["15%", "15%", "15%", "10%", "10%", "15%", "10%", "10%"];
  container.innerHTML = `
    <div class="weight-band">
      <div class="weight-spacer"></div>
      <div class="weight-grid">
        ${weights.map((w) => `<div class="weight-cell">${w}</div>`).join("")}
      </div>
    </div>
  `;
}

function renderOutlierRankTable(container) {
  container.innerHTML = `
    <table class="data-table rank-table">
      <thead>
        <tr>
          <th>Quintile Position</th>
          <th>Min Rank</th>
          <th>Max Rank</th>
        </tr>
      </thead>
      <tbody>
        <tr><td class="q1">1</td><td class="q1">1</td><td class="q1">1</td></tr>
        <tr><td class="q2">2</td><td class="q2">2</td><td class="q2">3</td></tr>
        <tr><td class="q3">3</td><td class="q3">3</td><td class="q3">3</td></tr>
        <tr><td class="q4">4</td><td class="q4">4</td><td class="q4">4</td></tr>
        <tr><td class="q5">5</td><td class="q5">5</td><td class="q5">5</td></tr>
      </tbody>
    </table>
  `;
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

function buildGroupedTable(tableEl, rows, labelField) {
  tableEl.innerHTML = "";

  if (!rows || !rows.length) {
    tableEl.innerHTML = `<thead><tr><th>No data found</th></tr></thead><tbody><tr><td>Nothing to display.</td></tr></tbody>`;
    return;
  }

  const metricGroups = resolveMetricGroupColumns(rows[0]);
  const baseHeaders = ["Month", labelField, "Quintile Position", "Overall Rank", "*"];
  const thead = document.createElement("thead");

  const groupRow = document.createElement("tr");
  baseHeaders.forEach((h) => {
    const th = document.createElement("th");
    th.rowSpan = 2;
    th.textContent = h;
    groupRow.appendChild(th);
  });

  metricGroups.forEach((g) => {
    const th = document.createElement("th");
    th.colSpan = 2;
    th.textContent = g.weight;
    groupRow.appendChild(th);
  });

  const subRow = document.createElement("tr");
  metricGroups.forEach((g) => {
    const thValue = document.createElement("th");
    thValue.textContent = g.label;
    subRow.appendChild(thValue);

    const thRank = document.createElement("th");
    thRank.textContent = `${g.label} Rank`;
    subRow.appendChild(thRank);
  });

  thead.appendChild(groupRow);
  thead.appendChild(subRow);
  tableEl.appendChild(thead);

  const tbody = document.createElement("tbody");

  rows.forEach((row) => {
    const tr = document.createElement("tr");

    baseHeaders.forEach((header) => {
      const td = document.createElement("td");

      if (header === "Month") {
        td.textContent = formatCellValue(getRowValue(row, ["Month", "MONTH", "Month Name"]));
      } else if (header === "Quintile Position") {
        td.textContent = formatCellValue(getRowValue(row, ["Quintile Position", "Quintile", "Q", "Qtr"]));
      } else if (header === "Overall Rank") {
        td.textContent = formatCellValue(getRowValue(row, ["Overall Rank", "OverallRank", "Rank", "Overall"]));
      } else if (header === "*") {
        td.textContent = "*";
      } else {
        td.textContent = formatCellValue(
          getRowValue(row, [
            labelField,
            `${labelField} Name`,
            "Region",
            "Region Name",
            "District",
            "District Name",
            "Store",
            "Store Name",
            "SAP",
            "Sap: Loaction",
            "Sap: Location",
            "STORE NAME",
            "STORE CODE",
          ])
        );
      }

      const q = getRowValue(row, ["Quintile Position", "Quintile", "Q", "Qtr"]);
      if (header === "Quintile Position") td.classList.add(quintileClass(q));
      if (header === "Overall Rank") td.classList.add(quintileClass(q));

      tr.appendChild(td);
    });

    metricGroups.forEach((g) => {
      const valueTd = document.createElement("td");
      valueTd.textContent = formatCellValue(
        getRowValue(row, [g.percentColumn, `${g.label} %Tgt`, `${g.label} % Target`, `${g.label}`, `${g.label} Actual`])
      );

      const rankValue = getRowValue(row, [g.rankColumn, `${g.label} Rank`, "Rank"]);
      const rankTd = document.createElement("td");
      rankTd.textContent = formatCellValue(rankValue);

      const quintileRules = state.metricRules[labelField.toLowerCase()] || [];
      const quintile = determineQuintileFromRank(rankValue, quintileRules);
      if (quintile) rankTd.classList.add(quintileClass(quintile));

      tr.appendChild(valueTd);
      tr.appendChild(rankTd);
    });

    tbody.appendChild(tr);
  });

  tableEl.appendChild(tbody);
}

function renderDashboard() {
  el.regionMeta.textContent = `${state.regions.length} rows`;
  el.districtMeta.textContent = `${state.districts.length} rows`;

  renderRankTable(el.regionRankHost, [
    [1, 1, 1],
    [2, 2, 3],
    [3, 3, 3],
    [4, 4, 4],
  ]);

  renderRankTable(el.districtRankHost, [
    [1, 1, 7],
    [2, 8, 14],
    [3, 15, 22],
    [4, 23, 29],
    [5, 30, 37],
  ]);

  renderWeightBand(el.regionWeightHost);
  renderWeightBand(el.districtWeightHost);

  buildGroupedTable(el.regionTable, state.regions, "Region");
  buildGroupedTable(el.districtTable, state.districts, "District");
  renderDistrictView();
}

function renderDistrictView() {
  const districtName = state.selectedDistrict;
  const districtRow = findDistrictRow(districtName);
  const storeRows = getDistrictStores(districtName);

  renderSummaryGrid(el.selectedDistrictSummary, districtRow);

  renderRankTable(el.storeRankHost, [
    [1, 1, 72],
    [2, 73, 144],
    [3, 145, 216],
    [4, 217, 288],
    [5, 289, 360],
  ]);

  renderWeightBand(el.storeWeightHost);

  buildGroupedTable(el.storeTable, storeRows, "Store");
}

function renderOutliers() {
  const metricColumn = state.selectedOutlierMetric;
  const metricLabel =
    el.outlierMetricSelect.options[el.outlierMetricSelect.selectedIndex]?.text || metricColumn || "";
  const mode = el.outlierModeSelect?.value || "lowest";
  const sortDirection = mode === "highest" ? -1 : 1;

  renderOutlierRankTable(el.outliersRankHost);

  const isCsat = normalizeLookup(metricLabel).includes("csat");
  const valueHeader = isCsat ? "CSAT Actual" : "Percent to Target";

  const formatOutlierValue = (value) => {
    const numeric = parseNumeric(value);
    if (numeric === null) return "";
    return isCsat ? numeric.toFixed(2) : `${numeric.toFixed(2)}%`;
  };

  const rows = state.stores
    .map((row) => {
      const storeName = normalizeText(
        row["Sap: Loaction"] || row["Sap: Location"] || row["STORE NAME"] || row["STORE CODE"] || row.SAP
      );
      const metricValue = parseNumeric(row[metricColumn]);

      return {
        "Store Name": storeName,
        [valueHeader]: formatOutlierValue(row[metricColumn]),
        __metricValue: metricValue,
      };
    })
    .filter((row) => row["Store Name"] && row.__metricValue !== null)
    .sort((a, b) => {
      if (a.__metricValue === b.__metricValue) {
        return a["Store Name"].toLowerCase().localeCompare(b["Store Name"].toLowerCase());
      }
      return (a.__metricValue - b.__metricValue) * sortDirection;
    })
    .slice(0, 20)
    .map(({ __metricValue, ...rest }) => rest);

  el.outliersMeta.innerHTML = `
    <div class="summary-grid">
      <div class="summary-item">
        <span class="summary-label">Metric</span>
        <span class="summary-value">${escapeHtml(metricLabel)}</span>
      </div>
      <div class="summary-item">
        <span class="summary-label">View</span>
        <span class="summary-value">${escapeHtml(mode === "highest" ? "Highest 20" : "Lowest 20")}</span>
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
  state.regions = enrichRows(regionRows, "region", state.metricRules.region).sort(
    (a, b) => parseNumeric(a["Overall Rank"]) - parseNumeric(b["Overall Rank"])
  );
  state.districts = enrichRows(districtRows, "district", state.metricRules.district).sort(
    (a, b) => parseNumeric(a["Overall Rank"]) - parseNumeric(b["Overall Rank"])
  );
  state.stores = enrichRows(storeRows, "store", state.metricRules.store).sort(
    (a, b) => parseNumeric(a["Overall Rank"]) - parseNumeric(b["Overall Rank"])
  );
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
