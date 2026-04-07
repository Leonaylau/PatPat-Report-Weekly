const state = {
  rawRows: [],
  filteredRows: [],
  excludedRows: [],
  currentStart: "",
  currentEnd: "",
  compareStart: "",
  compareEnd: "",
  pushCategoryFilter: "all",
  pushSearch: "",
  zeroRowFilter: "show",
};

const overviewMetrics = [
  { label: "Sessions", key: "sessions", type: "number", inverse: false },
  { label: "Users", key: "users", type: "number", inverse: false },
  { label: "Purchasers", key: "purchasers", type: "number", inverse: false },
  { label: "Revenue", key: "revenue", type: "currency", inverse: false },
];

document.addEventListener("DOMContentLoaded", async () => {
  bindEvents();
  await initializePushData();
});

function bindEvents() {
  document.getElementById("applyDateRange").addEventListener("click", () => {
    syncDateInputsToState();
    applyFiltersAndRender();
  });

  document.getElementById("usePreviousPeriod").addEventListener("click", () => {
    syncDateInputsToState();
    autoFillPreviousPeriod();
    syncStateToDateInputs();
    applyFiltersAndRender();
  });

  const pushCategoryFilter = document.getElementById("pushCategoryFilter");
  if (pushCategoryFilter) {
    pushCategoryFilter.addEventListener("change", (e) => {
      state.pushCategoryFilter = e.target.value;
      applyFiltersAndRender();
    });
  }

  const pushSearch = document.getElementById("pushSearch");
  if (pushSearch) {
    pushSearch.addEventListener("input", (e) => {
      state.pushSearch = e.target.value.trim().toLowerCase();
      applyFiltersAndRender();
    });
  }

  const zeroRowFilter = document.getElementById("zeroRowFilter");
  if (zeroRowFilter) {
    zeroRowFilter.addEventListener("change", (e) => {
      state.zeroRowFilter = e.target.value;
      applyFiltersAndRender();
    });
  }
}

async function initializePushData() {
  try {
    const pathCandidates = [
      "./data/current.csv",
      "./data/raw/current.csv",
      "./data/raw/push.csv",
    ];

    let text = null;
    let usedPath = "";

    for (const path of pathCandidates) {
      try {
        const response = await fetch(path, { cache: "no-store" });
        if (response.ok) {
          text = await response.text();
          usedPath = path;
          break;
        }
      } catch {
        // ignore
      }
    }

    if (!text) {
      throw new Error("No app-push CSV found in ./data/");
    }

    const parsed = parsePushCsv(text);
    state.rawRows = parsed.rows;

    if (state.rawRows.length) {
      const dates = state.rawRows
        .map((row) => row.dateObj)
        .filter(Boolean)
        .sort((a, b) => a - b);

      const minDate = dates[0];
      const maxDate = dates[dates.length - 1];

      state.currentStart = formatDateInput(minDate);
      state.currentEnd = formatDateInput(maxDate);
      autoFillPreviousPeriod();
      syncStateToDateInputs();
    }

    document.getElementById("fileMeta").textContent = `Loaded ${state.rawRows.length} app-push rows from ${usedPath}`;
    applyFiltersAndRender();
  } catch (error) {
    document.getElementById("fileMeta").textContent = `Failed to load app-push data: ${error.message}`;
    renderEmptyStates(error.message);
  }
}

function parsePushCsv(text) {
  const cleanLines = text
    .split(/\r?\n/)
    .filter((line) => line.trim() !== "")
    .filter((line) => !line.trim().startsWith("#"));

  const csvText = cleanLines.join("\n");
  const records = parseCsv(csvText);

  const rows = records
    .map(mapPushRow)
    .filter((row) => row.dateObj && row.pushName && row.pushName.toLowerCase() !== "grand total");

  return { rows };
}

function mapPushRow(raw) {
  const dateText = normalizeText(raw["Date"]);
  const pushName = normalizeText(raw["Session manual term"]);
  const dateObj = parseFlexibleDate(dateText);

  return {
    date: dateText,
    dateObj,
    pushName,
    sessions: parseNumber(raw["Sessions"]),
    users: parseNumber(raw["Total users"]),
    purchasers: parseNumber(raw["Total purchasers"]),
    revenue: parseNumber(raw["Total revenue"]),
  };
}

function syncDateInputsToState() {
  state.currentStart = document.getElementById("currentStart").value;
  state.currentEnd = document.getElementById("currentEnd").value;
  state.compareStart = document.getElementById("compareStart").value;
  state.compareEnd = document.getElementById("compareEnd").value;
}

function syncStateToDateInputs() {
  document.getElementById("currentStart").value = state.currentStart || "";
  document.getElementById("currentEnd").value = state.currentEnd || "";
  document.getElementById("compareStart").value = state.compareStart || "";
  document.getElementById("compareEnd").value = state.compareEnd || "";
}

function autoFillPreviousPeriod() {
  const currentStart = parseInputDate(state.currentStart);
  const currentEnd = parseInputDate(state.currentEnd);
  if (!currentStart || !currentEnd) return;

  const days = Math.round((currentEnd - currentStart) / 86400000) + 1;
  const compareEnd = new Date(currentStart);
  compareEnd.setDate(compareEnd.getDate() - 1);

  const compareStart = new Date(compareEnd);
  compareStart.setDate(compareStart.getDate() - (days - 1));

  state.compareStart = formatDateInput(compareStart);
  state.compareEnd = formatDateInput(compareEnd);
}

function applyFiltersAndRender() {
  const currentStart = parseInputDate(state.currentStart);
  const currentEnd = parseInputDate(state.currentEnd);

  if (!currentStart || !currentEnd) {
    renderEmptyStates("Please select a valid current period.");
    return;
  }

  const classified = state.rawRows.map((row) => {
    const pushCategory = classifyPushCategory(row.pushName);
    return { ...row, pushCategory };
  });

  state.excludedRows = [];

  state.filteredRows = classified.filter((row) => isRowIncluded(row, currentStart, currentEnd));

  renderAll();
}

function classifyPushCategory(pushName) {
  const name = String(pushName || "").trim().toLowerCase();
  return name.startsWith("push") ? "manual" : "automation";
}

function renderAll() {
  renderHeaderSummary();
  renderOverview();
  renderPushTable();
  renderContributionCharts();
  renderTrendCharts();
  renderExcludedTable();
}

function renderHeaderSummary() {
  const labelMap = {
    all: "All",
    manual: "Manual Push",
    automation: "Automation Push",
  };

  document.getElementById("headerSummary").textContent =
    `Current Period: ${state.currentStart || "--"} ~ ${state.currentEnd || "--"} ｜ Compare Period: ${state.compareStart || "--"} ~ ${state.compareEnd || "--"} ｜ Push Category: ${labelMap[state.pushCategoryFilter]}`;
}

function renderOverview() {
  const container = document.getElementById("overviewGrid");
  const context = document.getElementById("overviewContext");

  if (!state.filteredRows.length) {
    context.textContent = "Current period KPI summary and period-over-period movement.";
    container.innerHTML = '<div class="empty-state">No app-push rows match the selected filters.</div>';
    return;
  }

  const currentAgg = aggregateRows(state.filteredRows);
  const compareRows = getCompareRows();
  const compareAgg = aggregateRows(compareRows);

  context.textContent =
    `Current: ${state.currentStart} ~ ${state.currentEnd} · Compare: ${state.compareStart} ~ ${state.compareEnd}`;

  container.innerHTML = overviewMetrics.map((metric) => {
    const wow = buildWowObject(
      currentAgg[metric.key],
      compareAgg[metric.key],
      metric.type,
      metric.inverse
    );

    return `
      <article class="metric-card">
        <p class="metric-label">${escapeHtml(metric.label)}</p>
        <p class="metric-value">${escapeHtml(formatValue(currentAgg[metric.key], metric.type))}</p>
        <p class="metric-compare">Compare: ${escapeHtml(formatValue(compareAgg[metric.key], metric.type))}</p>
        <div class="metric-wow ${wow.className}">
          <span class="metric-wow-arrow">${escapeHtml(wow.arrow)}</span>
          <span class="metric-wow-value">${escapeHtml(wow.pctOnlyLabel)}</span>
        </div>
      </article>
    `;
  }).join("");
}

function getCompareRows() {
  const compareStart = parseInputDate(state.compareStart);
  const compareEnd = parseInputDate(state.compareEnd);
  if (!compareStart || !compareEnd) return [];

  return state.rawRows
    .map((row) => {
      const pushCategory = classifyPushCategory(row.pushName);
      return { ...row, pushCategory };
    })
    .filter((row) => isRowIncluded(row, compareStart, compareEnd));
}

function isRowIncluded(row, startDate, endDate) {
  const inDateRange = row.dateObj >= startDate && row.dateObj <= endDate;
  if (!inDateRange) return false;

  const categoryMatch =
    state.pushCategoryFilter === "all" ||
    (state.pushCategoryFilter === "manual" && row.pushCategory === "manual") ||
    (state.pushCategoryFilter === "automation" && row.pushCategory === "automation");

  const searchMatch =
    !state.pushSearch || row.pushName.toLowerCase().includes(state.pushSearch);

  const zeroRowMatch =
    state.zeroRowFilter === "show" || (toNumber(row.purchasers) || 0) !== 0;

  return categoryMatch && searchMatch && zeroRowMatch;
}

function aggregateRows(rows) {
  return rows.reduce(
    (acc, row) => {
      acc.sessions += toNumber(row.sessions) || 0;
      acc.users += toNumber(row.users) || 0;
      acc.purchasers += toNumber(row.purchasers) || 0;
      acc.revenue += toNumber(row.revenue) || 0;
      return acc;
    },
    { sessions: 0, users: 0, purchasers: 0, revenue: 0 }
  );
}

function renderPushTable() {
  const thead = document.querySelector("#pushTable thead");
  const tbody = document.querySelector("#pushTable tbody");
  const context = document.getElementById("pushTableContext");

  context.textContent =
    `Current: ${state.currentStart} ~ ${state.currentEnd} · Category: ${document.getElementById("pushCategoryFilter")?.selectedOptions?.[0]?.text || "All"}`;

  const grouped = groupByPushName(state.filteredRows);

  thead.innerHTML = `
    <tr>
      <th>Push Name</th>
      <th>Push Category</th>
      <th>Sessions</th>
      <th>Users</th>
      <th>Purchasers</th>
      <th>Revenue</th>
    </tr>
  `;

  if (!grouped.length) {
    tbody.innerHTML = '<tr><td class="empty-cell" colspan="6">No push rows match the current filters.</td></tr>';
    return;
  }

  tbody.innerHTML = grouped.map((row) => `
    <tr>
      <td>${escapeHtml(row.pushName)}</td>
      <td>${renderPushCategoryPill(row.pushCategory)}</td>
      <td>${escapeHtml(formatNumber(row.sessions))}</td>
      <td>${escapeHtml(formatNumber(row.users))}</td>
      <td>${escapeHtml(formatNumber(row.purchasers))}</td>
      <td>${escapeHtml(formatCurrency(row.revenue))}</td>
    </tr>
  `).join("");
}

function groupByPushName(rows) {
  const map = new Map();

  rows.forEach((row) => {
    const key = row.pushName;
    if (!map.has(key)) {
      map.set(key, {
        pushName: row.pushName,
        pushCategory: row.pushCategory,
        sessions: 0,
        users: 0,
        purchasers: 0,
        revenue: 0,
      });
    }

    const item = map.get(key);
    item.sessions += toNumber(row.sessions) || 0;
    item.users += toNumber(row.users) || 0;
    item.purchasers += toNumber(row.purchasers) || 0;
    item.revenue += toNumber(row.revenue) || 0;
  });

  return [...map.values()].sort((a, b) => b.sessions - a.sessions);
}

function renderPushCategoryPill(type) {
  const labelMap = {
    manual: "Manual Push",
    automation: "Automation Push",
  };

  const classMap = {
    manual: "push-type-pill push-current",
    automation: "push-type-pill push-previous",
  };

  return `<span class="${classMap[type]}">${labelMap[type]}</span>`;
}

function renderContributionCharts() {
  const grouped = groupByPushCategory(state.filteredRows);

  renderBarChart(
    document.getElementById("sessionsShareChart"),
    buildShareDataFromCategory(grouped, "sessions")
  );

  renderBarChart(
    document.getElementById("revenueShareChart"),
    buildShareDataFromCategory(grouped, "revenue")
  );
}

function groupByPushCategory(rows) {
  const base = {
    manual: { label: "Manual Push", sessions: 0, revenue: 0 },
    automation: { label: "Automation Push", sessions: 0, revenue: 0 },
  };

  rows.forEach((row) => {
    if (!base[row.pushCategory]) return;
    base[row.pushCategory].sessions += toNumber(row.sessions) || 0;
    base[row.pushCategory].revenue += toNumber(row.revenue) || 0;
  });

  return Object.values(base);
}

function buildShareDataFromCategory(rows, key) {
  const total = rows.reduce((sum, row) => sum + (toNumber(row[key]) || 0), 0);
  return rows.map((row) => ({
    label: row.label,
    share: total ? row[key] / total : 0,
  }));
}

function renderBarChart(container, data) {
  if (!data.length) {
    container.classList.add("empty-chart");
    container.innerHTML = "No data available for the current filter.";
    return;
  }

  container.classList.remove("empty-chart");
  container.innerHTML = data.map((item) => `
    <div class="bar-row">
      <div class="bar-label">${escapeHtml(item.label)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(item.share * 100, 0)}%"></div></div>
      <div class="bar-value">${escapeHtml(formatPercent(item.share))}</div>
    </div>
  `).join("");
}

function renderTrendCharts() {
  const container = document.getElementById("trendGrid");
  const currentStart = parseInputDate(state.currentStart);
  const currentEnd = parseInputDate(state.currentEnd);

  if (!currentStart || !currentEnd) {
    container.innerHTML = '<div class="chart-card empty-chart">Please select a valid current period.</div>';
    return;
  }

  const daily = buildDailySeries(currentStart, currentEnd, state.filteredRows);

  container.innerHTML = [
    { label: "Sessions Trend", key: "sessions", type: "number" },
    { label: "Revenue Trend", key: "revenue", type: "currency" },
  ].map((metric) => `
    <div class="chart-card">
      <h3>${escapeHtml(metric.label)}</h3>
      <div class="line-chart">${renderLineSvg(daily, metric)}</div>
    </div>
  `).join("");
}

function buildDailySeries(start, end, rows) {
  const dayMap = new Map();
  const cursor = new Date(start);

  while (cursor <= end) {
    const key = formatDateInput(cursor);
    dayMap.set(key, { label: key, sessions: 0, revenue: 0 });
    cursor.setDate(cursor.getDate() + 1);
  }

  rows.forEach((row) => {
    const key = formatDateInput(row.dateObj);
    if (!dayMap.has(key)) return;
    const item = dayMap.get(key);
    item.sessions += toNumber(row.sessions) || 0;
    item.revenue += toNumber(row.revenue) || 0;
  });

  return [...dayMap.values()];
}

function renderLineSvg(series, metric) {
  const points = series.map((item) => ({ label: item.label, value: item[metric.key] }));
  if (!points.length) return '<div class="empty-chart">No data available.</div>';

  const width = 620;
  const height = 260;
  const padding = { top: 20, right: 20, bottom: 52, left: 58 };
  const innerWidth = width - padding.left - padding.right;
  const innerHeight = height - padding.top - padding.bottom;
  const maxValue = Math.max(...points.map((point) => toNumber(point.value) || 0), 0.01);

  const scaleX = (index) =>
    padding.left + (points.length === 1 ? innerWidth / 2 : (index / (points.length - 1)) * innerWidth);
  const scaleY = (value) =>
    padding.top + innerHeight - ((toNumber(value) || 0) / maxValue) * innerHeight;

  const path = points
    .map((point, index) => `${index === 0 ? "M" : "L"}${scaleX(index)} ${scaleY(point.value)}`)
    .join(" ");

  const gridLines = [0.25, 0.5, 0.75, 1].map((factor) => {
    const y = padding.top + innerHeight - innerHeight * factor;
    return `<line class="grid-line" x1="${padding.left}" y1="${y}" x2="${padding.left + innerWidth}" y2="${y}"></line>`;
  }).join("");

  const dots = points.map((point, index) => {
    const x = scaleX(index);
    const y = scaleY(point.value);
    return `
      <g>
        <circle class="line-dot" cx="${x}" cy="${y}" r="4"></circle>
        <text class="line-value" x="${x}" y="${y - 10}" text-anchor="middle">${escapeHtml(formatValue(point.value, metric.type))}</text>
      </g>
    `;
  }).join("");

  const xTicks = points.map((point, index) =>
    `<text class="tick-label" x="${scaleX(index)}" y="${height - 18}" text-anchor="middle">${escapeHtml(shortDate(point.label))}</text>`
  ).join("");

  return `
    <svg class="line-svg" viewBox="0 0 ${width} ${height}" role="img">
      ${gridLines}
      <line class="grid-line" x1="${padding.left}" y1="${padding.top + innerHeight}" x2="${padding.left + innerWidth}" y2="${padding.top + innerHeight}"></line>
      <line class="grid-line" x1="${padding.left}" y1="${padding.top}" x2="${padding.left}" y2="${padding.top + innerHeight}"></line>
      <path class="line-path" d="${path}"></path>
      ${dots}
      ${xTicks}
    </svg>
  `;
}

function renderExcludedTable() {
  const thead = document.querySelector("#excludedTable thead");
  const tbody = document.querySelector("#excludedTable tbody");

  thead.innerHTML = `
    <tr>
      <th>日期</th>
      <th>Push Name</th>
      <th>Reason</th>
    </tr>
  `;

  if (!state.excludedRows.length) {
    tbody.innerHTML = '<tr><td class="empty-cell" colspan="3">No excluded push rows in the current dataset.</td></tr>';
    return;
  }

  tbody.innerHTML = state.excludedRows.map((row) => `
    <tr>
      <td>${escapeHtml(row.date)}</td>
      <td>${escapeHtml(row.pushName)}</td>
      <td>Excluded by classification rule</td>
    </tr>
  `).join("");
}

function buildWowObject(currentValue, priorValue, type, inverse) {
  const current = toNumber(currentValue) || 0;
  const prior = toNumber(priorValue) || 0;

  if (prior === 0) {
    return {
      pctOnlyLabel: "—",
      className: "wow-neutral",
      arrow: "→",
    };
  }

  const abs = current - prior;
  const pct = abs / prior;

  const good = inverse ? abs < 0 : abs > 0;
  const bad = inverse ? abs > 0 : abs < 0;

  return {
    pctOnlyLabel: formatSignedPercent(pct),
    className: good ? "wow-good" : bad ? "wow-bad" : "wow-neutral",
    arrow: abs > 0 ? "↑" : abs < 0 ? "↓" : "→",
  };
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let current = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      row.push(current);
      current = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(current);
      if (row.some((cell) => cell !== "")) rows.push(row);
      row = [];
      current = "";
    } else {
      current += char;
    }
  }

  if (current.length || row.length) {
    row.push(current);
    rows.push(row);
  }

  const [header = [], ...body] = rows;
  return body.map((cells) => {
    const record = {};
    header.forEach((key, headerIndex) => {
      record[key] = cells[headerIndex] ?? "";
    });
    return record;
  });
}

function parseFlexibleDate(value) {
  if (!value) return null;
  const raw = String(value).trim();

  const compact = raw.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (compact) {
    const year = Number(compact[1]);
    const month = Number(compact[2]);
    const day = Number(compact[3]);
    return new Date(year, month - 1, day);
  }

  const tryDate = new Date(raw);
  if (!isNaN(tryDate.getTime())) {
    return new Date(tryDate.getFullYear(), tryDate.getMonth(), tryDate.getDate());
  }

  const m = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
  if (m) {
    const month = Number(m[1]);
    const day = Number(m[2]);
    const year = Number(m[3].length === 2 ? `20${m[3]}` : m[3]);
    return new Date(year, month - 1, day);
  }

  return null;
}

function parseInputDate(value) {
  if (!value) return null;
  const date = new Date(value);
  if (isNaN(date.getTime())) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatDateInput(date) {
  if (!date) return "";
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function shortDate(value) {
  if (!value) return "";
  return value.slice(5);
}

function normalizeText(value) {
  return String(value ?? "").trim();
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const normalized = String(value).trim().replace(/[$,]/g, "");
  if (!normalized) return null;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function toNumber(value) {
  return typeof value === "number" ? value : parseNumber(value);
}

function formatValue(value, type) {
  if (value === null || value === undefined || value === "") return "-";
  if (type === "currency") return formatCurrency(value);
  return formatNumber(value);
}

function formatNumber(value) {
  const number = toNumber(value);
  if (!Number.isFinite(number)) return "-";
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(number);
}

function formatCurrency(value) {
  const number = toNumber(value);
  if (!Number.isFinite(number)) return "-";
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(number);
}

function formatPercent(value) {
  const number = toNumber(value);
  if (!Number.isFinite(number)) return "-";
  return new Intl.NumberFormat("en-US", {
    style: "percent",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(number);
}

function formatSignedPercent(value) {
  const number = toNumber(value);
  if (!Number.isFinite(number)) return "—";
  const sign = number > 0 ? "+" : number < 0 ? "-" : "";
  return `${sign}${formatPercent(Math.abs(number))}`;
}

function renderEmptyStates(message) {
  document.getElementById("overviewGrid").innerHTML = `<div class="empty-state">${escapeHtml(message)}</div>`;
  document.querySelector("#pushTable thead").innerHTML = "";
  document.querySelector("#pushTable tbody").innerHTML = `<tr><td class="empty-cell">${escapeHtml(message)}</td></tr>`;
  document.getElementById("sessionsShareChart").innerHTML = message;
  document.getElementById("revenueShareChart").innerHTML = message;
  document.getElementById("trendGrid").innerHTML = `<div class="chart-card empty-chart">${escapeHtml(message)}</div>`;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
