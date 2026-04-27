const MANIFEST_PATH = "./data/manifest.json";
const STORAGE_KEY = "fulfillmentDashboardColumns.v4";

const GA_FILE_CANDIDATES = [
  "./data/fulfillment-GA.csv",
  "./data/fulfillment_GA.csv",
  "./data/fulfillment-ga.csv",
  "./data/raw/fulfillment-GA.csv",
];

const MANUAL_TERM_TO_FLOWS = {
  delivered: ["Delivered", "Delivered - TEST"],
  order_canceled: ["Order canceled"],
  order_placed: ["Order placed"],
  comment: ["comment & nps", "comment & nps - TEST"],
  when_shipment_updates_to_in_transit: ["When shipment updates to in transit"],
  delivery_arriving_soon: ["Delivery arriving soon"],
};

const MANUAL_TERM_ALIASES = {
  delivery_arriing_soon: "delivery_arriving_soon",
};

const FLOW_TO_MANUAL_TERM = (() => {
  const map = new Map();
  Object.entries(MANUAL_TERM_TO_FLOWS).forEach(([term, flows]) => {
    flows.forEach((flow) => map.set(flow, term));
  });
  return map;
})();

function canonicalizeManualTerm(value) {
  const raw = String(value || "").trim().toLowerCase().replace(/\s+/g, "_");
  if (!raw || raw === "(not_set)") return "";
  return MANUAL_TERM_ALIASES[raw] || raw;
}

const overviewMetrics = [
  { label: "发送数", key: "sendCount", type: "number", inverse: false },
  { label: "送达数", key: "deliveredCount", type: "number", inverse: false },
  { label: "打开率", key: "openRate", type: "rate", inverse: false },
  { label: "点击率", key: "clickRate", type: "rate", inverse: false },
  { label: "打开点击率", key: "ctor", type: "rate", inverse: false },
  { label: "退订率", key: "unsubscribeRate", type: "rate", inverse: true },
  { label: "Sessions (GA)", key: "sessions", type: "number", inverse: false },
  { label: "下单uv", key: "orderUv", type: "number", inverse: false },
  { label: "Revenue", key: "revenue", type: "currency", inverse: false },
  { label: "ARPU", key: "arpu", type: "currency", inverse: false },
  { label: "客净利", key: "revenuePerOrder", type: "currency", inverse: false },
];

const defaultColumns = [
  { key: "flow", label: "Flow", type: "text" },
  { key: "manualTerm", label: "Manual term", type: "text" },
  { key: "sendCount", label: "发送数", type: "number" },
  { key: "deliveredCount", label: "送达数", type: "number" },
  { key: "openRate", label: "打开率", type: "rate" },
  { key: "clickRate", label: "点击率", type: "rate" },
  { key: "ctor", label: "打开点击率", type: "rate" },
  { key: "unsubscribeRate", label: "退订率", type: "rate" },
  { key: "sessions", label: "Sessions", type: "number" },
  { key: "orderUv", label: "下单uv", type: "number" },
  { key: "revenue", label: "Revenue", type: "currency" },
  { key: "sessionsWow", label: "Sessions WoW %", type: "wowRateOnly" },
  { key: "revenueWow", label: "Revenue WoW %", type: "wowRateOnly" },
  { key: "orderUvWow", label: "下单uv WoW %", type: "wowRateOnly" },
  { key: "openRateWow", label: "打开率 WoW %", type: "wowRateOnly" },
  { key: "clickRateWow", label: "点击率 WoW %", type: "wowRateOnly" },
  { key: "ctorWow", label: "打开点击率 WoW %", type: "wowRateOnly" },
  { key: "unsubscribeRateWow", label: "退订率 WoW %", type: "wowRateOnly" },
];

const trendMetrics = [
  { label: "发送数趋势", key: "sendCount", type: "number" },
  { label: "打开率趋势", key: "openRate", type: "rate" },
  { label: "点击率趋势", key: "clickRate", type: "rate" },
  { label: "Sessions 趋势", key: "sessions", type: "number" },
  { label: "Revenue 趋势", key: "revenue", type: "currency" },
  { label: "下单uv 趋势", key: "orderUv", type: "number" },
];

const state = {
  store: { weeks: {}, weekOrder: [] },
  gaRows: [],
  gaByWeek: new Map(),
  selectedWeek: "",
  compareWeek: "",
  currentWeek: null,
  priorWeek: null,
  filteredRows: [],
  sortKey: "sendCount",
  sortDirection: "desc",
  flowFilterValue: "",
  searchValue: "",
  statusMessage: "Loading repository data...",
  columns: loadColumnPrefs(),
  showColumnPanel: false,
  dragColumnKey: "",
};

document.addEventListener("DOMContentLoaded", async () => {
  bindEvents();
  initTooltip();
  await initializeRepositoryData();
});

function bindEvents() {
  document.getElementById("weekSelect").addEventListener("change", (event) => {
    state.selectedWeek = event.target.value;
    autoSelectCompareWeek();
    hydrateCurrentWeek();
  });

  document.getElementById("compareWeekSelect").addEventListener("change", (event) => {
    state.compareWeek = event.target.value;
    hydrateCurrentWeek();
  });

  document.getElementById("flowSearch").addEventListener("input", (event) => {
    state.searchValue = event.target.value.trim().toLowerCase();
    applyFiltersAndRender();
  });

  document.getElementById("flowFilter").addEventListener("change", (event) => {
    state.flowFilterValue = event.target.value;
    applyFiltersAndRender();
  });

  document.getElementById("exportWeekButton").addEventListener("click", exportCurrentWeek);
  document.getElementById("toggleColumnsButton").addEventListener("click", toggleColumnPanel);
  document.getElementById("resetColumnsButton").addEventListener("click", resetColumns);
}

async function initializeRepositoryData() {
  try {
    const manifest = await fetchJson(MANIFEST_PATH);
    const entries = normalizeManifestEntries(manifest);

    if (!entries.length) {
      state.statusMessage = `No weeks found in ${MANIFEST_PATH}.`;
      render();
      return;
    }

    const weekRanges = computeWeekRanges(entries.map((entry) => entry.weekLabel));
    const weeklyCsvs = [];
    for (const entry of entries) {
      const csvText = await fetchText(entry.path);
      const rows = parseCsv(csvText);
      weeklyCsvs.push({ entry, rows, csvText });
    }

    let gaPath = "";
    let gaRows = [];
    for (const candidate of GA_FILE_CANDIDATES) {
      try {
        const response = await fetch(candidate, { cache: "no-store" });
        if (response.ok) {
          const text = await response.text();
          gaRows = parseGaCsv(text);
          gaPath = candidate;
          break;
        }
      } catch {
        // ignore
      }
    }
    state.gaRows = gaRows;
    state.gaByWeek = bucketGaByWeek(gaRows, weekRanges);

    for (const { entry, rows, csvText } of weeklyCsvs) {
      state.store.weeks[entry.weekLabel] = buildWeekData({
        rows,
        csvText,
        weekLabel: entry.weekLabel,
        fileName: entry.fileName,
        weekRange: weekRanges.get(entry.weekLabel) || null,
        gaForWeek: state.gaByWeek.get(entry.weekLabel) || new Map(),
      });
    }

    state.store.weekOrder = sortWeekLabels(Object.keys(state.store.weeks));
    state.selectedWeek = state.store.weekOrder[state.store.weekOrder.length - 1] || "";
    autoSelectCompareWeek();
    const gaInfo = gaPath
      ? `, ${gaRows.length} GA rows from ${gaPath}`
      : `, GA file not found (Sessions / Revenue / Order UV will be empty)`;
    state.statusMessage = `Loaded ${state.store.weekOrder.length} week(s) from ${MANIFEST_PATH}${gaInfo}.`;
    if (state.selectedWeek) {
      document.getElementById("dataFreshness").textContent = `数据截至: ${state.selectedWeek}`;
    }
    hydrateCurrentWeek();
  } catch (error) {
    state.statusMessage = `Failed to load repository data: ${error.message}`;
    render();
  }
}

function autoSelectCompareWeek() {
  const index = state.store.weekOrder.indexOf(state.selectedWeek);
  const priorLabel = index > 0 ? state.store.weekOrder[index - 1] : "";
  if (!state.compareWeek || state.compareWeek === state.selectedWeek || !state.store.weeks[state.compareWeek]) {
    state.compareWeek = priorLabel;
  }
}

async function fetchJson(path) {
  const response = await fetch(path, { cache: "no-store" });
  if (!response.ok) throw new Error(`${path} returned ${response.status}`);
  return response.json();
}

async function fetchText(path) {
  const response = await fetch(path, { cache: "no-store" });
  if (!response.ok) throw new Error(`${path} returned ${response.status}`);
  return response.text();
}

function normalizeManifestEntries(manifest) {
  const rawWeeks = Array.isArray(manifest)
    ? manifest
    : Array.isArray(manifest?.weeks)
      ? manifest.weeks
      : [];

  return rawWeeks.map((entry) => normalizeManifestEntry(entry)).filter(Boolean);
}

function normalizeManifestEntry(entry) {
  if (typeof entry === "string") {
    const weekLabel = normalizeWeekLabel(entry);
    if (!weekLabel) return null;
    return { weekLabel, fileName: `${weekLabel}.csv`, path: `./data/${weekLabel}.csv` };
  }

  if (!entry || typeof entry !== "object") return null;

  const weekLabel = normalizeWeekLabel(
    entry.week || entry.label || entry.name || entry.file || entry.filename || ""
  );
  const fileName = normalizeFileName(entry.file || entry.filename || `${weekLabel}.csv`);
  if (!weekLabel || !fileName) return null;

  return {
    weekLabel,
    fileName,
    path: fileName.startsWith("./") ? fileName : `./data/${fileName}`,
  };
}

function normalizeFileName(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (raw.endsWith(".csv")) return raw.split("/").pop();
  return `${raw.split("/").pop()}.csv`;
}

function buildWeekData({ rows, csvText, weekLabel, fileName, weekRange, gaForWeek }) {
  const mappedRows = rows.map(mapRow).filter((row) => row.flow || row.trigger);
  const summary = mappedRows.find((row) => row.flow === "Summary") || null;
  let detailRows = mappedRows.filter((row) => row.flow && row.flow !== "Summary");
  const allocationNotes = [];
  const gaMap = gaForWeek instanceof Map ? gaForWeek : new Map();

  detailRows = detailRows.map((row) => {
    const manualTerm = FLOW_TO_MANUAL_TERM.get(row.flow) || "";
    const ga = manualTerm ? gaMap.get(manualTerm) : null;
    const sessions = ga ? ga.sessions : null;
    const revenue = ga ? ga.revenue : null;
    const orderUv = ga ? ga.orderCount : null;
    return {
      ...row,
      manualTerm,
      sessions,
      revenue,
      orderUv,
      arpu: safeDivide(revenue, row.deliveredCount),
      revenuePerOrder: safeDivide(revenue, orderUv),
    };
  });

  if (summary) {
    let summarySessions = 0;
    let summaryRevenue = 0;
    let summaryOrderUv = 0;
    let hasGaForWeek = false;
    gaMap.forEach((agg) => {
      summarySessions += toNumber(agg.sessions) || 0;
      summaryRevenue += toNumber(agg.revenue) || 0;
      summaryOrderUv += toNumber(agg.orderCount) || 0;
      hasGaForWeek = true;
    });
    if (hasGaForWeek) {
      summary.sessions = summarySessions;
      summary.revenue = summaryRevenue;
      summary.orderUv = summaryOrderUv;
      summary.arpu = safeDivide(summaryRevenue, summary.deliveredCount);
      summary.revenuePerOrder = safeDivide(summaryRevenue, summaryOrderUv);
    } else {
      summary.sessions = null;
    }
  }

  if (gaMap.size === 0) {
    allocationNotes.push("fulfillment-GA has no rows for this week's date range; Sessions / Revenue / Order UV are empty.");
  } else {
    allocationNotes.push("Sessions / Revenue / Order UV are joined from fulfillment-GA on Manual term. TEST flows share the same GA aggregate as their production counterparts (intentional duplicate attribution).");
  }

  return {
    weekLabel,
    fileName,
    rawCsv: csvText,
    weekRange,
    summary,
    detailRows,
    rows: detailRows,
    allocationNotes,
  };
}

function hydrateCurrentWeek() {
  state.currentWeek = state.store.weeks[state.selectedWeek] || null;
  state.priorWeek = state.compareWeek ? state.store.weeks[state.compareWeek] || null : null;
  populateWeekSelectors();
  populateFlowFilter();
  applyFiltersAndRender();
}

function populateWeekSelectors() {
  const currentSelect = document.getElementById("weekSelect");
  const compareSelect = document.getElementById("compareWeekSelect");
  const weeks = state.store.weekOrder;

  if (!weeks.length) {
    currentSelect.innerHTML = '<option value="">No week loaded</option>';
    compareSelect.innerHTML = '<option value="">No week loaded</option>';
    return;
  }

  currentSelect.innerHTML = weeks
    .map((label) => `<option value="${escapeAttribute(label)}">${escapeHtml(label)}</option>`)
    .join("");
  currentSelect.value = state.selectedWeek;

  compareSelect.innerHTML =
    '<option value="">No comparison</option>' +
    weeks
      .filter((label) => label !== state.selectedWeek)
      .map((label) => `<option value="${escapeAttribute(label)}">${escapeHtml(label)}</option>`)
      .join("");

  compareSelect.value =
    state.compareWeek && state.compareWeek !== state.selectedWeek ? state.compareWeek : "";
}

function populateFlowFilter() {
  const rows = state.currentWeek?.detailRows || [];
  const flows = [...new Set(rows.map((row) => row.flow).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b)
  );
  const select = document.getElementById("flowFilter");

  select.innerHTML =
    '<option value="">All flows</option>' +
    flows.map((flow) => `<option value="${escapeAttribute(flow)}">${escapeHtml(flow)}</option>`).join("");

  if (state.flowFilterValue && !flows.includes(state.flowFilterValue)) {
    state.flowFilterValue = "";
  }
  select.value = state.flowFilterValue;
}

function applyFiltersAndRender() {
  const detailRows = state.currentWeek?.detailRows || [];
  const priorMap = new Map((state.priorWeek?.detailRows || []).map((row) => [row.flow, row]));

  state.filteredRows = sortRows(
    detailRows
      .filter((row) => {
        const flowSearchMatch = !state.searchValue || row.flow.toLowerCase().includes(state.searchValue);
        const flowFilterMatch = !state.flowFilterValue || row.flow === state.flowFilterValue;
        return flowSearchMatch && flowFilterMatch;
      })
      .map((row) => enrichFlowWow(row, priorMap.get(row.flow) || null, Boolean(state.priorWeek))),
    state.sortKey,
    state.sortDirection
  );

  render();
}

function enrichFlowWow(currentRow, priorRow, hasPriorWeek) {
  return {
    ...currentRow,
    sessionsWow: buildWowObject(currentRow.sessions, priorRow?.sessions, "number", false, hasPriorWeek, Boolean(priorRow)),
    revenueWow: buildWowObject(currentRow.revenue, priorRow?.revenue, "currency", false, hasPriorWeek, Boolean(priorRow)),
    orderUvWow: buildWowObject(currentRow.orderUv, priorRow?.orderUv, "number", false, hasPriorWeek, Boolean(priorRow)),
    openRateWow: buildWowObject(currentRow.openRate, priorRow?.openRate, "rate", false, hasPriorWeek, Boolean(priorRow)),
    clickRateWow: buildWowObject(currentRow.clickRate, priorRow?.clickRate, "rate", false, hasPriorWeek, Boolean(priorRow)),
    ctorWow: buildWowObject(currentRow.ctor, priorRow?.ctor, "rate", false, hasPriorWeek, Boolean(priorRow)),
    unsubscribeRateWow: buildWowObject(currentRow.unsubscribeRate, priorRow?.unsubscribeRate, "rate", true, hasPriorWeek, Boolean(priorRow)),
  };
}

function render() {
  renderMeta();
  renderHeaderSummary();
  renderOverview();
  renderFlowTable();
  renderContributionCharts();
  renderEfficiencyCharts();
  renderTrendCharts();
  renderColumnPanel();
}

function renderMeta() {
  const notes = state.currentWeek?.allocationNotes || [];
  const exportButton = document.getElementById("exportWeekButton");
  exportButton.disabled = !state.currentWeek;
  document.getElementById("fileMeta").textContent = notes.length
    ? `${state.statusMessage} ${notes.join(" ")}`
    : state.statusMessage;
}

function renderHeaderSummary() {
  const summaryEl = document.getElementById("headerSummary");
  const current = state.currentWeek?.weekLabel || "--";
  const compare = state.priorWeek?.weekLabel || "--";
  summaryEl.textContent = `Current Period: ${current} ｜ Compare Period: ${compare} ｜ Source: Repository CSV`;
}

function renderOverview() {
  const container = document.getElementById("overviewGrid");
  const context = document.getElementById("overviewContext");

  if (!state.currentWeek?.summary) {
    context.textContent = "Select current week and comparison week to view KPI performance.";
    container.innerHTML = '<div class="empty-state">Repository data is not available yet.</div>';
    return;
  }

  context.textContent = state.priorWeek
    ? `Current: ${state.currentWeek.weekLabel} · Compare: ${state.priorWeek.weekLabel}`
    : `Current: ${state.currentWeek.weekLabel}`;

  container.innerHTML = overviewMetrics
    .map((metric) => {
      const currentValue = state.currentWeek.summary[metric.key];
      const priorValue = state.priorWeek?.summary ? state.priorWeek.summary[metric.key] : null;
      const wow = buildWowObject(
        currentValue,
        priorValue,
        metric.type,
        metric.inverse,
        Boolean(state.priorWeek),
        Boolean(state.priorWeek?.summary)
      );

      return `
        <article class="metric-card">
          <p class="metric-label">${escapeHtml(metric.label)}</p>
          <p class="metric-value">${escapeHtml(formatValue(currentValue, metric.type))}</p>
          <div class="metric-wow ${wow.className}">
            <span class="metric-wow-arrow">${escapeHtml(wow.arrow)}</span>
            <span class="metric-wow-value">${escapeHtml(wow.pctOnlyLabel)}</span>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderFlowTable() {
  const thead = document.querySelector("#flowTable thead");
  const tbody = document.querySelector("#flowTable tbody");
  const context = document.getElementById("flowContext");
  const visibleColumns = getVisibleColumns();

  if (!state.currentWeek?.detailRows?.length) {
    context.textContent = "Current period flow-level performance compared with the selected comparison period.";
    thead.innerHTML = "";
    tbody.innerHTML = '<tr><td class="empty-cell">Repository-backed flow data is not available yet.</td></tr>';
    return;
  }

  context.textContent = state.priorWeek
    ? `Current: ${state.currentWeek.weekLabel} · Compare: ${state.priorWeek.weekLabel}`
    : `Current: ${state.currentWeek.weekLabel}`;

  thead.innerHTML = `<tr>${visibleColumns
    .map((column) => {
      const active = state.sortKey === column.key;
      const arrow = active ? (state.sortDirection === "asc" ? " ↑" : " ↓") : "";
      return `<th><button type="button" data-sort-key="${column.key}">${escapeHtml(column.label)}${arrow}</button></th>`;
    })
    .join("")}</tr>`;

  thead.querySelectorAll("button[data-sort-key]").forEach((button) => {
    button.addEventListener("click", () => handleSort(button.dataset.sortKey));
  });

  if (!state.filteredRows.length) {
    tbody.innerHTML = `<tr><td class="empty-cell" colspan="${visibleColumns.length}">No flows match the current search or flow filter.</td></tr>`;
    return;
  }

  tbody.innerHTML = state.filteredRows
    .map(
      (row) => `
      <tr>
        ${visibleColumns.map((column) => renderTableCell(row, column)).join("")}
      </tr>
    `
    )
    .join("");
}

function renderTableCell(row, column) {
  if (column.type === "wowRateOnly") {
    const wow = row[column.key];
    return `
      <td>
        <div class="wow-rate-only ${wow.className}">
          <span class="wow-rate-arrow">${escapeHtml(wow.arrow)}</span>
          <span class="wow-rate-value">${escapeHtml(wow.pctOnlyLabel)}</span>
        </div>
      </td>
    `;
  }

  const className = column.key === "flow" ? "flow-name" : "";
  return `<td class="${className}">${escapeHtml(formatValue(row[column.key], column.type))}</td>`;
}

function renderContributionCharts() {
  renderBarChart(document.getElementById("clickShareChart"), buildShareData("clickCount"));
  renderBarChart(document.getElementById("revenueShareChart"), buildShareData("revenue"));
  renderBarChart(document.getElementById("orderShareChart"), buildShareData("orderUv"));
}

function renderEfficiencyCharts() {
  renderScatterChart({
    element: document.getElementById("efficiencyChartOne"),
    rows: state.filteredRows,
    xKey: "openRate",
    yKey: "ctor",
    sizeKey: "sendCount",
    xLabel: "打开率",
    yLabel: "打开点击率",
    yType: "rate",
  });

  renderScatterChart({
    element: document.getElementById("efficiencyChartTwo"),
    rows: state.filteredRows,
    xKey: "clickRate",
    yKey: "revenuePerOrder",
    sizeKey: "deliveredCount",
    xLabel: "点击率",
    yLabel: "客净利",
    yType: "currency",
  });
}

function renderTrendCharts() {
  const container = document.getElementById("trendGrid");
  const allWeeks = state.store.weekOrder
    .map((label) => ({ label, summary: state.store.weeks[label]?.summary || null }))
    .filter((item) => item.summary);
  const weekSeries = allWeeks.slice(-4);

  if (!weekSeries.length) {
    container.innerHTML = '<div class="chart-card empty-chart">Repository trend data is not available yet.</div>';
    return;
  }

  const overlayColors = ["#c96442", "#248a3d", "#87867f"];

  container.innerHTML = trendMetrics
    .map((metric) => {
      const flowNames = new Set();
      weekSeries.forEach((w) => {
        const weekObj = state.store.weeks[w.label];
        if (weekObj?.rows) weekObj.rows.forEach((r) => flowNames.add(r.flow));
      });

      const flowTotals = [...flowNames].map((flow) => {
        const total = weekSeries.reduce((sum, w) => {
          const weekObj = state.store.weeks[w.label];
          const row = weekObj?.rows?.find((r) => r.flow === flow);
          return sum + (row ? toNumber(row[metric.key]) || 0 : 0);
        }, 0);
        return { flow, total };
      }).sort((a, b) => b.total - a.total);

      const topFlows = flowTotals.slice(0, 3);
      const overlays = topFlows.map((tf, i) => ({
        label: tf.flow,
        color: overlayColors[i % overlayColors.length],
        points: weekSeries.map((w) => {
          const weekObj = state.store.weeks[w.label];
          const row = weekObj?.rows?.find((r) => r.flow === tf.flow);
          return { label: w.label, value: row ? toNumber(row[metric.key]) || 0 : 0 };
        }),
      }));

      return `
      <div class="chart-card">
        <h3>${escapeHtml(metric.label)}</h3>
        <div class="line-chart">${renderLineSvg(weekSeries, metric, overlays)}</div>
      </div>
    `;
    })
    .join("");
}

function buildShareData(valueKey) {
  if (!state.filteredRows.length) return [];
  const total = state.filteredRows.reduce((sum, row) => sum + (toNumber(row[valueKey]) || 0), 0);

  return state.filteredRows
    .map((row) => ({ label: row.flow, share: safeDivide(row[valueKey], total) || 0 }))
    .sort((a, b) => b.share - a.share);
}

function renderBarChart(container, data) {
  if (!data.length) {
    container.classList.add("empty-chart");
    container.innerHTML = "No data available for the current filter.";
    return;
  }

  container.classList.remove("empty-chart");
  container.innerHTML = data
    .map(
      (item) => `
      <div class="bar-row chart-tip" data-label="${escapeAttribute(item.label)}" data-value="${escapeHtml(formatPercent(item.share))}">
        <div class="bar-label" title="${escapeAttribute(item.label)}">${escapeHtml(item.label)}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${Math.max(item.share * 100, 0)}%"></div></div>
        <div class="bar-value">${escapeHtml(formatPercent(item.share))}</div>
      </div>
    `
    )
    .join("");
}

function renderScatterChart({ element, rows, xKey, yKey, sizeKey, xLabel, yLabel, yType }) {
  if (!rows.length) {
    element.classList.add("empty-chart");
    element.innerHTML = "No data available for the current filter.";
    return;
  }

  const cleanRows = rows.filter((row) => row[xKey] !== null && row[yKey] !== null && row[sizeKey] !== null);
  if (!cleanRows.length) {
    element.classList.add("empty-chart");
    element.innerHTML = "No data available for the current filter.";
    return;
  }

  const width = 620;
  const height = 360;
  const padding = { top: 20, right: 20, bottom: 54, left: 58 };
  const innerWidth = width - padding.left - padding.right;
  const innerHeight = height - padding.top - padding.bottom;
  const xMax = Math.max(...cleanRows.map((row) => toNumber(row[xKey]) || 0), 0.01);
  const yMax = Math.max(...cleanRows.map((row) => toNumber(row[yKey]) || 0), 0.01);
  const sizeMax = Math.max(...cleanRows.map((row) => toNumber(row[sizeKey]) || 0), 1);

  const scaleX = (value) => padding.left + (safeDivide(value, xMax) || 0) * innerWidth;
  const scaleY = (value) => padding.top + innerHeight - (safeDivide(value, yMax) || 0) * innerHeight;
  const scaleR = (value) => 8 + Math.sqrt((toNumber(value) || 0) / sizeMax) * 28;

  const gridLines = [0.25, 0.5, 0.75, 1]
    .map((factor) => {
      const x = padding.left + innerWidth * factor;
      const y = padding.top + innerHeight - innerHeight * factor;
      return `
      <line class="grid-line" x1="${x}" y1="${padding.top}" x2="${x}" y2="${padding.top + innerHeight}"></line>
      <line class="grid-line" x1="${padding.left}" y1="${y}" x2="${padding.left + innerWidth}" y2="${y}"></line>
    `;
    })
    .join("");

  const bubbles = cleanRows
    .map((row) => {
      const cx = scaleX(row[xKey]);
      const cy = scaleY(row[yKey]);
      const r = scaleR(row[sizeKey]);
      const tipLabel = escapeHtml(row.flow);
      const tipValue = escapeHtml(`${xLabel}: ${formatValue(row[xKey], "rate")} | ${yLabel}: ${formatValue(row[yKey], yType)}`);
      return `
      <g class="chart-tip" data-label="${tipLabel}" data-value="${tipValue}">
        <circle class="bubble" cx="${cx}" cy="${cy}" r="${r}"></circle>
        <text class="bubble-text" x="${cx}" y="${cy + 4}" text-anchor="middle">${escapeHtml(shortFlow(row.flow))}</text>
      </g>
    `;
    })
    .join("");

  const xTicks = [0, 0.25, 0.5, 0.75, 1]
    .map((factor) => {
      const value = xMax * factor;
      const x = padding.left + innerWidth * factor;
      return `<text class="axis-label" x="${x}" y="${height - 24}" text-anchor="middle">${escapeHtml(
        formatValue(value, "rate")
      )}</text>`;
    })
    .join("");

  const yTicks = [0, 0.25, 0.5, 0.75, 1]
    .map((factor) => {
      const value = yMax * factor;
      const y = padding.top + innerHeight - innerHeight * factor;
      return `<text class="axis-label" x="${padding.left - 12}" y="${y + 4}" text-anchor="end">${escapeHtml(
        formatValue(value, yType)
      )}</text>`;
    })
    .join("");

  element.classList.remove("empty-chart");
  element.innerHTML = `
    <svg class="scatter-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeAttribute(
      xLabel
    )} and ${escapeAttribute(yLabel)} bubble chart">
      ${gridLines}
      <line class="grid-line" x1="${padding.left}" y1="${padding.top + innerHeight}" x2="${padding.left + innerWidth}" y2="${padding.top + innerHeight}"></line>
      <line class="grid-line" x1="${padding.left}" y1="${padding.top}" x2="${padding.left}" y2="${padding.top + innerHeight}"></line>
      ${bubbles}
      ${xTicks}
      ${yTicks}
      <text class="axis-label" x="${padding.left + innerWidth / 2}" y="${height - 4}" text-anchor="middle">${escapeHtml(
        xLabel
      )}</text>
      <text class="axis-label" x="18" y="${padding.top + innerHeight / 2}" text-anchor="middle" transform="rotate(-90 18 ${
        padding.top + innerHeight / 2
      })">${escapeHtml(yLabel)}</text>
    </svg>
    <p class="legend-note">Hover bubbles for metric details.</p>
  `;
}

function renderLineSvg(series, metric, overlays) {
  const points = series
    .map((item) => ({ label: item.label, value: item.summary[metric.key] }))
    .filter((item) => item.value !== null && item.value !== undefined);

  if (!points.length) return '<div class="empty-chart">No data available.</div>';

  const width = 820;
  const height = 260;
  const padding = { top: 20, right: 60, bottom: 52, left: 58 };
  const innerWidth = width - padding.left - padding.right;
  const innerHeight = height - padding.top - padding.bottom;
  const maxValue = Math.max(...points.map((point) => toNumber(point.value) || 0), 0.01);

  const scaleX = (index) =>
    padding.left + (points.length === 1 ? innerWidth / 2 : (index / (points.length - 1)) * innerWidth);
  const scaleY = (value) => padding.top + innerHeight - (safeDivide(value, maxValue) || 0) * innerHeight;

  const path = points
    .map((point, index) => `${index === 0 ? "M" : "L"}${scaleX(index)} ${scaleY(point.value)}`)
    .join(" ");

  const gridLines = [0.25, 0.5, 0.75, 1]
    .map((factor) => {
      const y = padding.top + innerHeight - innerHeight * factor;
      return `<line class="grid-line" x1="${padding.left}" y1="${y}" x2="${padding.left + innerWidth}" y2="${y}"></line>`;
    })
    .join("");

  const dots = points
    .map((point, index) => {
      const x = scaleX(index);
      const y = scaleY(point.value);
      return `
      <g class="chart-tip" data-label="${escapeHtml(point.label)}" data-value="${escapeHtml(formatValue(point.value, metric.type))}">
        <circle cx="${x}" cy="${y}" r="16" fill="transparent"></circle>
        <circle class="line-dot" cx="${x}" cy="${y}" r="4"></circle>
        <text class="line-value" x="${x}" y="${y - 10}" text-anchor="middle">${escapeHtml(
        formatValue(point.value, metric.type)
      )}</text>
      </g>
    `;
    })
    .join("");

  const xTicks = points
    .map((point, index) =>
      `<text class="tick-label" x="${scaleX(index)}" y="${height - 18}" text-anchor="middle">${escapeHtml(
        point.label
      )}</text>`
    )
    .join("");

  const yTicks = [0, 0.25, 0.5, 0.75, 1]
    .map((factor) => {
      const value = maxValue * factor;
      const y = padding.top + innerHeight - innerHeight * factor;
      return `<text class="tick-label" x="${padding.left - 10}" y="${y + 4}" text-anchor="end">${escapeHtml(
        formatValue(value, metric.type)
      )}</text>`;
    })
    .join("");

  const overlayPaths = (overlays || []).map((ov) => {
    const ovPath = ov.points
      .map((p, i) => `${i === 0 ? "M" : "L"}${scaleX(i)} ${scaleY(p.value)}`)
      .join(" ");
    return `<path class="line-path-overlay" d="${ovPath}" style="stroke:${ov.color};stroke-width:1.5;fill:none;opacity:0.7"></path>`;
  }).join("");

  const overlayLegend = (overlays || []).length ? `<div class="chart-legend">${(overlays || []).map((ov) =>
    `<span class="legend-item"><span class="legend-swatch" style="background:${ov.color}"></span>${escapeHtml(ov.label)}</span>`
  ).join("")}</div>` : "";

  return `
    <svg class="line-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeAttribute(
      metric.label
    )} trend chart">
      ${gridLines}
      <line class="grid-line" x1="${padding.left}" y1="${padding.top + innerHeight}" x2="${padding.left + innerWidth}" y2="${padding.top + innerHeight}"></line>
      <line class="grid-line" x1="${padding.left}" y1="${padding.top}" x2="${padding.left}" y2="${padding.top + innerHeight}"></line>
      ${overlayPaths}
      <path class="line-path" d="${path}"></path>
      ${dots}
      ${xTicks}
      ${yTicks}
    </svg>
    ${overlayLegend}
  `;
}

function toggleColumnPanel() {
  state.showColumnPanel = !state.showColumnPanel;
  renderColumnPanel();
}

function resetColumns() {
  state.columns = defaultColumns.map((column) => ({ ...column, visible: true }));
  persistColumnPrefs();
  renderColumnPanel();
  renderFlowTable();
}

function renderColumnPanel() {
  const panel = document.getElementById("columnPanel");
  if (!panel) return;

  if (!state.showColumnPanel) {
    panel.innerHTML = "";
    panel.classList.add("hidden");
    return;
  }

  panel.classList.remove("hidden");
  panel.innerHTML = `
    <div class="column-panel-inner">
      <div class="column-panel-head">
        <h3>Customize Columns</h3>
        <p>Drag to reorder. Uncheck to hide.</p>
      </div>
      <div class="column-list">
        ${state.columns
          .map(
            (column) => `
          <label class="column-item" draggable="true" data-column-key="${escapeAttribute(column.key)}">
            <span class="drag-handle">⋮⋮</span>
            <input type="checkbox" data-column-toggle="${escapeAttribute(column.key)}" ${
              column.visible !== false ? "checked" : ""
            }>
            <span>${escapeHtml(column.label)}</span>
          </label>
        `
          )
          .join("")}
      </div>
    </div>
  `;

  panel.querySelectorAll("[data-column-toggle]").forEach((input) => {
    input.addEventListener("change", (event) => {
      const key = event.target.dataset.columnToggle;
      const column = state.columns.find((item) => item.key === key);
      if (column) {
        column.visible = event.target.checked;
        persistColumnPrefs();
        renderFlowTable();
      }
    });
  });

  panel.querySelectorAll(".column-item").forEach((item) => {
    item.addEventListener("dragstart", handleColumnDragStart);
    item.addEventListener("dragover", handleColumnDragOver);
    item.addEventListener("drop", handleColumnDrop);
    item.addEventListener("dragend", handleColumnDragEnd);
  });
}

function handleColumnDragStart(event) {
  const key = event.currentTarget.dataset.columnKey;
  state.dragColumnKey = key;
  event.dataTransfer.effectAllowed = "move";
}

function handleColumnDragOver(event) {
  event.preventDefault();
  event.dataTransfer.dropEffect = "move";
}

function handleColumnDrop(event) {
  event.preventDefault();
  const targetKey = event.currentTarget.dataset.columnKey;
  const draggedKey = state.dragColumnKey;
  if (!draggedKey || !targetKey || draggedKey === targetKey) return;

  const draggedIndex = state.columns.findIndex((item) => item.key === draggedKey);
  const targetIndex = state.columns.findIndex((item) => item.key === targetKey);
  if (draggedIndex === -1 || targetIndex === -1) return;

  const [dragged] = state.columns.splice(draggedIndex, 1);
  state.columns.splice(targetIndex, 0, dragged);
  persistColumnPrefs();
  renderColumnPanel();
  renderFlowTable();
}

function handleColumnDragEnd() {
  state.dragColumnKey = "";
}

function getVisibleColumns() {
  return state.columns.filter((column) => column.visible !== false);
}

function loadColumnPrefs() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultColumns.map((column) => ({ ...column, visible: true }));

    const parsed = JSON.parse(raw);
    const merged = defaultColumns.map((base) => {
      const matched = parsed.find((item) => item.key === base.key);
      return matched ? { ...base, visible: matched.visible !== false } : { ...base, visible: true };
    });

    const ordered = [];
    parsed.forEach((saved) => {
      const found = merged.find((item) => item.key === saved.key);
      if (found && !ordered.some((item) => item.key === found.key)) ordered.push(found);
    });
    merged.forEach((item) => {
      if (!ordered.some((entry) => entry.key === item.key)) ordered.push(item);
    });

    return ordered;
  } catch {
    return defaultColumns.map((column) => ({ ...column, visible: true }));
  }
}

function persistColumnPrefs() {
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify(
      state.columns.map((column) => ({
        key: column.key,
        visible: column.visible !== false,
      }))
    )
  );
}

function exportCurrentWeek() {
  if (!state.currentWeek?.rawCsv) return;

  const blob = new Blob([state.currentWeek.rawCsv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = state.currentWeek.fileName || `${state.currentWeek.weekLabel}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function handleSort(sortKey) {
  if (state.sortKey === sortKey) {
    state.sortDirection = state.sortDirection === "asc" ? "desc" : "asc";
  } else {
    state.sortKey = sortKey;
    state.sortDirection = "desc";
  }
  applyFiltersAndRender();
}

function sortRows(rows, sortKey, direction) {
  const column =
    state.columns.find((item) => item.key === sortKey) || defaultColumns.find((item) => item.key === sortKey);
  const type = column?.type || "text";
  const sorted = [...rows].sort((a, b) => compareValues(getSortableValue(a[sortKey], type), getSortableValue(b[sortKey], type), type));
  return direction === "asc" ? sorted : sorted.reverse();
}

function getSortableValue(value, type) {
  if (type === "wowRateOnly") {
    return value?.pctSortValue ?? Number.NEGATIVE_INFINITY;
  }
  return value;
}

function compareValues(a, b, type) {
  if (type === "text") {
    return String(a || "").localeCompare(String(b || ""));
  }
  return (toNumber(a) || 0) - (toNumber(b) || 0);
}

function buildWowObject(currentValue, priorValue, type, inverse, hasPriorWeek, hasPriorMatch) {
  if (!hasPriorWeek) {
    return {
      absLabel: "—",
      pctLabel: "—",
      pctOnlyLabel: "—",
      className: "wow-neutral",
      sortValue: Number.NEGATIVE_INFINITY,
      pctSortValue: Number.NEGATIVE_INFINITY,
      arrow: "→",
    };
  }

  if (!hasPriorMatch || priorValue === null || priorValue === undefined) {
    return {
      absLabel: "new",
      pctLabel: "—",
      pctOnlyLabel: "—",
      className: "wow-neutral",
      sortValue: Number.POSITIVE_INFINITY,
      pctSortValue: Number.POSITIVE_INFINITY,
      arrow: "→",
    };
  }

  const current = toNumber(currentValue);
  const prior = toNumber(priorValue);

  if (!Number.isFinite(current) || !Number.isFinite(prior)) {
    return {
      absLabel: "—",
      pctLabel: "—",
      pctOnlyLabel: "—",
      className: "wow-neutral",
      sortValue: Number.NEGATIVE_INFINITY,
      pctSortValue: Number.NEGATIVE_INFINITY,
      arrow: "→",
    };
  }

  const abs = current - prior;
  const pct = prior === 0 ? null : abs / prior;

  const good = inverse ? abs < 0 : abs > 0;
  const bad = inverse ? abs > 0 : abs < 0;
  const className = good ? "wow-good" : bad ? "wow-bad" : "wow-neutral";

  return {
    absLabel: formatDelta(abs, type),
    pctLabel: pct === null ? "—" : formatSignedPercent(pct),
    pctOnlyLabel: pct === null ? "—" : formatSignedPercent(pct),
    className,
    sortValue: abs,
    pctSortValue: pct ?? Number.NEGATIVE_INFINITY,
    arrow: abs > 0 ? "↑" : abs < 0 ? "↓" : "→",
  };
}

function formatDelta(value, type) {
  const sign = value > 0 ? "+" : value < 0 ? "-" : "";
  const absValue = Math.abs(value);
  if (type === "currency") return `${sign}${formatCurrency(absValue)}`;
  if (type === "rate") return `${sign}${formatPercent(absValue)}`;
  return `${sign}${formatNumber(absValue)}`;
}

function formatSignedPercent(value) {
  const sign = value > 0 ? "+" : value < 0 ? "-" : "";
  return `${sign}${formatPercent(Math.abs(value))}`;
}

function mapRow(raw) {
  const sendCount = parseNumber(raw["Emails sent"]);
  const deliveredCount = parseNumber(raw["Emails delivered"]);
  const openCount = parseNumber(raw["Emails opened"]);
  const clickCount = parseNumber(raw["Emails clicked"]);
  const unsubscribeCount = parseNumber(raw["Unsubscribed"]);
  const orderUv = parseNumber(raw["Email attributed orders"]);
  const revenue = parseNumber(raw["Email attributed revenue"]);

  return {
    flow: normalizeText(raw["Flow"]),
    trigger: normalizeText(raw["Trigger"]),
    sendCount,
    deliveredCount,
    openCount,
    clickCount,
    unsubscribeCount,
    openRate: parsePercent(raw["Open rate"]),
    clickRate: parsePercent(raw["Click rate"]),
    unsubscribeRate: parsePercent(raw["Unsubscribe rate"]),
    orderUv,
    revenue,
    clickCount,
    ctor: safeDivide(clickCount, openCount),
    arpu: safeDivide(revenue, deliveredCount),
    revenuePerOrder: safeDivide(revenue, orderUv),
  };
}

function parseGaCsv(text) {
  const cleanLines = text
    .split(/\r?\n/)
    .filter((line) => line.trim() !== "")
    .filter((line) => !line.trim().startsWith("#"));

  const csvText = cleanLines.join("\n");
  const records = parseCsv(csvText);

  const rows = [];
  records.forEach((raw) => {
    const dateText = normalizeText(pickField(raw, ["Date", "date"]));
    const txIdRaw = normalizeText(pickField(raw, ["Transaction ID", "transaction ID", "transaction id", "TransactionID"]));
    const manualTermRaw = normalizeText(pickField(raw, [
      "Manual term", "manual term", "Session manual term",
      "Session Manual term", "session manual term",
    ]));
    const sessionsRaw = pickField(raw, ["Sessions", "sessions", "会话数"]);
    const revenueRaw = pickField(raw, ["Total revenue", "total revenue", "Revenue", "总收入"]);

    const dateObj = parseFlexibleDate(dateText);
    if (!dateObj) return;
    if (/grand total/i.test(manualTermRaw) || /grand total/i.test(dateText)) return;

    const manualTerm = canonicalizeManualTerm(manualTermRaw);
    if (!manualTerm) return;
    if (!Object.prototype.hasOwnProperty.call(MANUAL_TERM_TO_FLOWS, manualTerm)) return;

    const transactionId = /^\(not\s*set\)$/i.test(txIdRaw) ? "" : txIdRaw;
    rows.push({
      dateObj,
      transactionId,
      manualTerm,
      sessions: toNumber(sessionsRaw) || 0,
      revenue: toNumber(revenueRaw) || 0,
    });
  });
  return rows;
}

function bucketGaByWeek(gaRows, weekRanges) {
  const result = new Map();
  weekRanges.forEach((_range, label) => result.set(label, new Map()));

  const ordered = [...weekRanges.entries()];
  gaRows.forEach((row) => {
    const match = ordered.find(([, range]) => range && row.dateObj >= range.start && row.dateObj <= range.end);
    if (!match) return;
    const [label] = match;
    const map = result.get(label);
    if (!map.has(row.manualTerm)) {
      map.set(row.manualTerm, { sessions: 0, revenue: 0, orderCount: 0 });
    }
    const agg = map.get(row.manualTerm);
    agg.sessions += row.sessions;
    agg.revenue += row.revenue;
    if (row.transactionId) agg.orderCount += 1;
  });
  return result;
}

function computeWeekRanges(weekLabels) {
  const result = new Map();
  const explicit = [];
  const dayOnly = [];
  weekLabels.forEach((label, index) => {
    const range = parseExplicitWeekLabel(label);
    if (range) {
      explicit.push({ index, label, range });
      result.set(label, range);
    } else {
      dayOnly.push({ index, label });
    }
  });

  if (!dayOnly.length) return result;

  const referenceLabel = explicit[0];
  if (!referenceLabel) {
    weekLabels.forEach((label) => result.set(label, null));
    return result;
  }

  // For each day-only label, walk backwards from the first explicit week.
  const referenceStart = referenceLabel.range.start;
  const offset = referenceLabel.index;
  dayOnly
    .sort((a, b) => a.index - b.index)
    .forEach(({ index, label }) => {
      const weeksBefore = offset - index;
      const start = addDays(referenceStart, -7 * weeksBefore);
      const end = addDays(start, 6);
      const days = parseDayOnlyLabel(label);
      // Sanity check: confirm Sun day-of-month matches
      if (days && start.getDate() !== days.startDay) {
        // Fall back to inferred date even if Sun mismatch — record it but proceed.
      }
      result.set(label, { start, end });
    });

  return result;
}

function parseExplicitWeekLabel(label) {
  const value = String(label || "").trim();
  const match = value.match(/^week(\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})$/i);
  if (!match) return null;
  const [, m1, d1, m2, d2] = match.map(Number);
  const year = inferYearFromContext();
  const start = buildLocalDate(year, m1, d1);
  let end = buildLocalDate(year, m2, d2);
  // Cross-year wrap (e.g. week12.30-1.5)
  if (end < start) end = buildLocalDate(year + 1, m2, d2);
  if (!start || !end) return null;
  return { start, end };
}

function parseDayOnlyLabel(label) {
  const match = String(label || "").trim().match(/^week(\d{1,2})-(\d{1,2})$/i);
  if (!match) return null;
  return { startDay: Number(match[1]), endDay: Number(match[2]) };
}

function inferYearFromContext() {
  return new Date().getFullYear();
}

function buildLocalDate(year, month, day) {
  if (!year || !month || !day) return null;
  const date = new Date(year, month - 1, day);
  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }
  return date;
}

function addDays(date, amount) {
  const next = new Date(date);
  next.setDate(next.getDate() + amount);
  return new Date(next.getFullYear(), next.getMonth(), next.getDate());
}

function parseFlexibleDate(value) {
  if (!value) return null;
  const raw = String(value).trim();

  const compact = raw.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (compact) {
    return buildLocalDate(Number(compact[1]), Number(compact[2]), Number(compact[3]));
  }

  const tryDate = new Date(raw);
  if (!Number.isNaN(tryDate.getTime())) {
    return new Date(tryDate.getFullYear(), tryDate.getMonth(), tryDate.getDate());
  }

  const m = raw.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
  if (m) {
    const month = Number(m[1]);
    const day = Number(m[2]);
    const year = Number(m[3].length === 2 ? `20${m[3]}` : m[3]);
    return buildLocalDate(year, month, day);
  }

  return null;
}

function pickField(record, keys) {
  for (const key of keys) {
    if (record[key] !== undefined && record[key] !== null && String(record[key]).trim() !== "") {
      return record[key];
    }
  }
  return "";
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

function normalizeWeekLabel(value) {
  const raw = String(value || "").trim().replace(/\.csv$/i, "").replace(/\s+/g, "");
  if (!raw) return "";
  const matched = raw.match(/week(\d+)-(\d+)/i);
  if (matched) return `week${Number(matched[1])}-${Number(matched[2])}`;
  return raw;
}

function sortWeekLabels(labels) {
  return [...new Set(labels)].sort((a, b) => {
    const parsedA = parseWeekSortInfo(a);
    const parsedB = parseWeekSortInfo(b);
    if (parsedA && parsedB) return parsedA.start - parsedB.start || parsedA.end - parsedB.end;
    if (parsedA) return -1;
    if (parsedB) return 1;
    return String(a).localeCompare(String(b));
  });
}

function parseWeekSortInfo(label) {
  const matched = String(label || "").match(/week(\d+)-(\d+)/i);
  if (!matched) return null;
  return { start: Number(matched[1]), end: Number(matched[2]) };
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const normalized = String(value).trim().replace(/[$,]/g, "");
  if (!normalized) return null;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function parsePercent(value) {
  if (value === null || value === undefined || value === "") return null;
  const normalized = String(value).trim().replace("%", "");
  if (!normalized) return null;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed / 100 : null;
}

function safeDivide(numerator, denominator) {
  const num = toNumber(numerator);
  const den = toNumber(denominator);
  if (!Number.isFinite(num) || !Number.isFinite(den) || den === 0) return null;
  return num / den;
}

function toNumber(value) {
  return typeof value === "number" ? value : parseNumber(value);
}

function normalizeText(value) {
  return String(value ?? "").trim();
}

function formatValue(value, type) {
  if (value === null || value === undefined || value === "") return "-";
  if (type === "rate" || type === "wowRateOnly") return formatPercent(value);
  if (type === "currency") return formatCurrency(value);
  if (type === "text") return String(value);
  return formatNumber(value);
}

function formatNumber(value) {
  const number = toNumber(value);
  if (!Number.isFinite(number)) return "-";
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(number);
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

function shortFlow(flow) {
  const parts = String(flow || "").split(/\s+/).filter(Boolean);
  return parts.length ? parts.slice(0, 2).map((part) => part[0]).join("").toUpperCase() : "";
}

function initTooltip() {
  const tip = document.createElement("div");
  tip.className = "chart-tooltip";
  document.body.appendChild(tip);

  document.addEventListener("mouseover", (e) => {
    const t = e.target.closest(".chart-tip");
    if (!t) return;
    tip.textContent = `${t.dataset.label}: ${t.dataset.value}`;
    tip.style.display = "block";
    const r = t.getBoundingClientRect();
    tip.style.left = `${r.left + r.width / 2 - tip.offsetWidth / 2}px`;
    tip.style.top = `${r.top - tip.offsetHeight - 8 + window.scrollY}px`;
  });

  document.addEventListener("mouseout", (e) => {
    if (e.target.closest(".chart-tip")) tip.style.display = "none";
  });
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replace(/`/g, "&#96;");
}
