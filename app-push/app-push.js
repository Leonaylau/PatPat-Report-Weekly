const state = {
  rawRows: [],
  classifiedRows: [],
  overviewRows: [],
  filteredRows: [],
  excludedRows: [],
  currentStart: "",
  currentEnd: "",
  compareStart: "",
  compareEnd: "",
  pushCategoryFilter: "all",
  pushBucketFilter: "all",
  pushSearch: "",
  zeroRowFilter: "show",
};

const overviewMetrics = [
  { label: "Sessions", key: "sessions", type: "number", inverse: false },
  { label: "Users / UV", key: "users", type: "number", inverse: false },
  { label: "Purchasers", key: "purchasers", type: "number", inverse: false },
  { label: "Revenue", key: "revenue", type: "currency", inverse: false },
  { label: "AOV (客单价)", key: "aov", type: "currency", inverse: false },
  { label: "CVR (转化率)", key: "cvr", type: "rate", inverse: false },
];

const pushCategoryLabelMap = {
  all: "All",
  manual: "Manual Push",
  automation: "Automation Push",
};

const pushBucketLabelMap = {
  all: "All Push Names",
  current: "Current Push",
  past: "Past Push",
  automation: "Automation Push",
  future: "Future Push",
  invalid: "Invalid Manual Name",
};

document.addEventListener("DOMContentLoaded", async () => {
  bindEvents();
  initTooltip();
  await initializePushData();
});

function bindEvents() {
  document.getElementById("applyDateRange")?.addEventListener("click", () => {
    syncDateInputsToState();
    applyFiltersAndRender();
  });

  document.getElementById("useLatestWeek")?.addEventListener("click", () => {
    syncDateInputsToState();
    const start = parseInputDate(state.currentStart);
    const end = parseInputDate(state.currentEnd);
    if (!start || !end || start > end) return;
    const days = Math.round((end - start) / (1000 * 60 * 60 * 24));
    const compareEnd = addDays(start, -1);
    const compareStart = addDays(compareEnd, -days);
    state.compareStart = formatDateInput(compareStart);
    state.compareEnd = formatDateInput(compareEnd);
    syncStateToDateInputs();
    applyFiltersAndRender();
  });

  document.getElementById("pushCategoryFilter")?.addEventListener("change", (e) => {
    state.pushCategoryFilter = e.target.value;
    applyFiltersAndRender();
  });

  document.getElementById("pushBucketFilter")?.addEventListener("change", (e) => {
    state.pushBucketFilter = e.target.value;
    applyFiltersAndRender();
  });

  document.getElementById("pushSearch")?.addEventListener("input", (e) => {
    state.pushSearch = e.target.value.trim().toLowerCase();
    applyFiltersAndRender();
  });

  document.getElementById("zeroRowFilter")?.addEventListener("change", (e) => {
    state.zeroRowFilter = e.target.value;
    applyFiltersAndRender();
  });
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
      setLatestCompleteWeekRange();
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
  const pushName = normalizePushName(normalizeText(raw["Session manual term"]));
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

function normalizePushName(name) {
  return name.replace(/^(push04)0(1[0-2])/i, "$1$2");
}

function syncDateInputsToState() {
  state.currentStart = document.getElementById("currentStart")?.value || "";
  state.currentEnd = document.getElementById("currentEnd")?.value || "";
  state.compareStart = document.getElementById("compareStart")?.value || "";
  state.compareEnd = document.getElementById("compareEnd")?.value || "";
}

function syncStateToDateInputs() {
  const currentStart = document.getElementById("currentStart");
  const currentEnd = document.getElementById("currentEnd");
  const compareStart = document.getElementById("compareStart");
  const compareEnd = document.getElementById("compareEnd");

  if (currentStart) currentStart.value = state.currentStart || "";
  if (currentEnd) currentEnd.value = state.currentEnd || "";
  if (compareStart) compareStart.value = state.compareStart || "";
  if (compareEnd) compareEnd.value = state.compareEnd || "";
}

function setLatestCompleteWeekRange() {
  const latestRange = getLatestCompletedWeekRange(state.rawRows);
  if (!latestRange) return;

  state.currentStart = formatDateInput(latestRange.currentStart);
  state.currentEnd = formatDateInput(latestRange.currentEnd);
  state.compareStart = formatDateInput(latestRange.compareStart);
  state.compareEnd = formatDateInput(latestRange.compareEnd);
}

function getLatestCompletedWeekRange(rows) {
  const dates = (rows || [])
    .map((row) => row.dateObj)
    .filter(Boolean)
    .sort((a, b) => a - b);

  if (!dates.length) return null;

  const latestDataDate = dates[dates.length - 1];
  const currentEnd = getPreviousOrSameSaturday(latestDataDate);
  const currentStart = addDays(currentEnd, -6);
  const compareEnd = addDays(currentStart, -1);
  const compareStart = addDays(compareEnd, -6);

  return { currentStart, currentEnd, compareStart, compareEnd };
}

function applyFiltersAndRender() {
  const currentStart = parseInputDate(state.currentStart);
  const currentEnd = parseInputDate(state.currentEnd);

  if (!currentStart || !currentEnd || currentStart > currentEnd) {
    renderEmptyStates("Please select a valid current period.");
    return;
  }

  state.classifiedRows = state.rawRows.map((row) => classifyRow(row, currentStart, currentEnd));
  state.excludedRows = state.classifiedRows.filter((row) => row.pushBucket === "future" || row.pushBucket === "invalid");

  state.overviewRows = state.classifiedRows.filter((row) => {
    const inDateRange = row.dateObj >= currentStart && row.dateObj <= currentEnd;
    return inDateRange && row.pushBucket !== "future" && row.pushBucket !== "invalid";
  });

  state.filteredRows = state.classifiedRows.filter((row) => isRowIncluded(row, currentStart, currentEnd));

  renderAll();
}

function classifyRow(row, currentStart, currentEnd) {
  const pushCategory = classifyPushCategory(row.pushName);
  const resolvedManualDate = resolvePushNameDate(row.pushName, currentStart, currentEnd);
  const pushBucket = classifyPushBucket(pushCategory, resolvedManualDate, currentStart, currentEnd);

  return {
    ...row,
    pushCategory,
    resolvedManualDate,
    pushBucket,
  };
}

function classifyPushCategory(pushName) {
  const name = String(pushName || "").trim().toLowerCase();

  const automationPushNames = [
    "pushwelcom01",
    "cartship",
    "cart30m",
    "cart2h",
    "cart24h",
    "quickship",
    "view3",
    "checkout",
    "checkout1",
    "cart15off",
  ];

  if (automationPushNames.includes(name)) {
    return "automation";
  }

  if (/pu+sh/i.test(name)) {
    return "manual";
  }

  return "automation";
}

function resolvePushNameDate(pushName, currentStart, currentEnd) {
  const name = String(pushName || "").trim().toLowerCase();
  const match = name.match(/pu+sh\s*(\d{2})(\d{2})/i);
  if (!match) return null;

  const month = Number(match[1]);
  const day = Number(match[2]);
  if (!month || !day || month > 12 || day > 31) return null;

  const yearCandidates = new Set([
    currentStart.getFullYear() - 1,
    currentStart.getFullYear(),
    currentEnd.getFullYear(),
    currentEnd.getFullYear() + 1,
  ]);

  const candidates = [...yearCandidates]
    .map((year) => buildSafeDate(year, month, day))
    .filter(Boolean)
    .sort((a, b) => a - b);

  if (!candidates.length) return null;

  const inRange = candidates.find((candidate) => candidate >= currentStart && candidate <= currentEnd);
  if (inRange) return inRange;

  const pastCandidates = candidates.filter((candidate) => candidate < currentStart);
  if (pastCandidates.length) return pastCandidates[pastCandidates.length - 1];

  return candidates[0];
}

function classifyPushBucket(pushCategory, resolvedManualDate, currentStart, currentEnd) {
  if (pushCategory === "automation") return "automation";
  if (!resolvedManualDate) return "invalid";
  if (resolvedManualDate >= currentStart && resolvedManualDate <= currentEnd) return "current";
  if (resolvedManualDate < currentStart) return "past";
  return "future";
}

function isRowIncluded(row, startDate, endDate) {
  const inDateRange = row.dateObj >= startDate && row.dateObj <= endDate;
  if (!inDateRange) return false;
  if (!matchesPushFilters(row)) return false;
  if (row.pushBucket === "future" || row.pushBucket === "invalid") return false;
  return true;
}

function matchesPushFilters(row) {
  const categoryMatch =
    state.pushCategoryFilter === "all" ||
    (state.pushCategoryFilter === "manual" && row.pushCategory === "manual") ||
    (state.pushCategoryFilter === "automation" && row.pushCategory === "automation");

  const bucketMatch =
    state.pushBucketFilter === "all" ||
    (state.pushBucketFilter === "current" && row.pushBucket === "current") ||
    (state.pushBucketFilter === "past" && row.pushBucket === "past") ||
    (state.pushBucketFilter === "automation" && row.pushBucket === "automation") ||
    (state.pushBucketFilter === "future" && row.pushBucket === "future") ||
    (state.pushBucketFilter === "invalid" && row.pushBucket === "invalid");

  const searchMatch = !state.pushSearch || row.pushName.toLowerCase().includes(state.pushSearch);
  const zeroRowMatch = state.zeroRowFilter === "show" || (toNumber(row.purchasers) || 0) !== 0;

  return categoryMatch && bucketMatch && searchMatch && zeroRowMatch;
}

function renderAll() {
  renderHeaderSummary();
  renderOverview();
  renderPushTable();
  renderContributionCharts();
  renderTrendCharts();
  renderAutomationWeeklyDetail();
  renderExcludedTable();
}

function renderHeaderSummary() {
  document.getElementById("headerSummary").textContent =
    `Current Period: ${state.currentStart || "--"} ~ ${state.currentEnd || "--"} ｜ Compare Period: ${state.compareStart || "--"} ~ ${state.compareEnd || "--"} ｜ Push Category: ${pushCategoryLabelMap[state.pushCategoryFilter]} ｜ Push Name Group: ${pushBucketLabelMap[state.pushBucketFilter]}`;
}

function renderOverview() {
  const container = document.getElementById("overviewGrid");
  const context = document.getElementById("overviewContext");

  if (!state.overviewRows.length) {
    context.textContent = "Current period KPI summary and period-over-period movement.";
    container.innerHTML = '<div class="empty-state">No app-push rows in the selected date range.</div>';
    return;
  }

  const currentAgg = aggregateRows(state.overviewRows);

  const compareStart = parseInputDate(state.compareStart);
  const compareEnd = parseInputDate(state.compareEnd);
  const compareRows = (compareStart && compareEnd)
    ? state.classifiedRows.filter((row) => {
        const inRange = row.dateObj >= compareStart && row.dateObj <= compareEnd;
        return inRange && row.pushBucket !== "future" && row.pushBucket !== "invalid";
      })
    : [];
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
    .map((row) => classifyRow(row, compareStart, compareEnd))
    .filter((row) => isRowIncluded(row, compareStart, compareEnd));
}

function aggregateRows(rows) {
  const agg = rows.reduce(
    (acc, row) => {
      acc.sessions += toNumber(row.sessions) || 0;
      acc.users += toNumber(row.users) || 0;
      acc.purchasers += toNumber(row.purchasers) || 0;
      acc.revenue += toNumber(row.revenue) || 0;
      return acc;
    },
    { sessions: 0, users: 0, purchasers: 0, revenue: 0 }
  );
  agg.aov = agg.purchasers ? agg.revenue / agg.purchasers : 0;
  agg.cvr = agg.users ? agg.purchasers / agg.users : 0;
  return agg;
}

function renderPushTable() {
  const thead = document.querySelector("#pushTable thead");
  const tbody = document.querySelector("#pushTable tbody");
  const context = document.getElementById("pushTableContext");
  const pushOverview = document.getElementById("pushOverviewGrid");

  context.textContent =
    `Current: ${state.currentStart} ~ ${state.currentEnd} · Push Category: ${pushCategoryLabelMap[state.pushCategoryFilter]} · Push Name Group: ${pushBucketLabelMap[state.pushBucketFilter]}`;

  const grouped = groupByPushName(state.filteredRows);

  const currentAgg = aggregateRows(state.filteredRows);
  const compareRows = getCompareRows();
  const compareAgg = aggregateRows(compareRows);

  if (!state.filteredRows.length) {
    pushOverview.innerHTML = "";
  } else {
    pushOverview.innerHTML = overviewMetrics.map((metric) => {
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

  thead.innerHTML = `
    <tr>
      <th>Push Name</th>
      <th>Push Category</th>
      <th>Push Name Group</th>
      <th>Resolved Push Date</th>
      <th>Sessions</th>
      <th>Users / UV</th>
      <th>Purchasers</th>
      <th>Revenue</th>
      <th>AOV (客单价)</th>
      <th>CVR (转化率)</th>
    </tr>
  `;

  if (!grouped.length) {
    tbody.innerHTML = '<tr><td class="empty-cell" colspan="10">No push rows match the current filters.</td></tr>';
    return;
  }

  tbody.innerHTML = grouped.map((row) => {
    const aov = (toNumber(row.purchasers) || 0) ? (toNumber(row.revenue) || 0) / (toNumber(row.purchasers) || 1) : 0;
    const cvr = (toNumber(row.users) || 0) ? (toNumber(row.purchasers) || 0) / (toNumber(row.users) || 1) : 0;
    return `
    <tr>
      <td>${escapeHtml(row.pushName)}</td>
      <td>${renderPill(row.pushCategory, "category")}</td>
      <td>${renderPill(row.pushBucket, "bucket")}</td>
      <td>${escapeHtml(row.resolvedManualDate ? formatDateInput(row.resolvedManualDate) : "-")}</td>
      <td>${escapeHtml(formatNumber(row.sessions))}</td>
      <td>${escapeHtml(formatNumber(row.users))}</td>
      <td>${escapeHtml(formatNumber(row.purchasers))}</td>
      <td>${escapeHtml(formatCurrency(row.revenue))}</td>
      <td>${escapeHtml(formatCurrency(aov))}</td>
      <td>${escapeHtml(formatPercent(cvr))}</td>
    </tr>
  `}).join("");
}

function groupByPushName(rows) {
  const map = new Map();

  rows.forEach((row) => {
    const key = row.pushName;
    if (!map.has(key)) {
      map.set(key, {
        pushName: row.pushName,
        pushCategory: row.pushCategory,
        pushBucket: row.pushBucket,
        resolvedManualDate: row.resolvedManualDate,
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

  return [...map.values()].sort((a, b) => {
    const revenueGap = (toNumber(b.revenue) || 0) - (toNumber(a.revenue) || 0);
    if (revenueGap !== 0) return revenueGap;
    return (toNumber(b.sessions) || 0) - (toNumber(a.sessions) || 0);
  });
}

function renderPill(value, type) {
  const labelMap = type === "category" ? pushCategoryLabelMap : pushBucketLabelMap;
  const className = `push-type-pill ${type === "category" ? `pill-category-${value}` : `pill-bucket-${value}`}`;
  return `<span class="${className}">${escapeHtml(labelMap[value] || value)}</span>`;
}

function renderContributionCharts() {
  const grouped = groupByPushBucket(state.filteredRows);

  renderBarChart(
    document.getElementById("sessionsShareChart"),
    buildShareData(grouped, "sessions")
  );

  renderBarChart(
    document.getElementById("revenueShareChart"),
    buildShareData(grouped, "revenue")
  );
}

function groupByPushBucket(rows) {
  const base = {
    current: { label: "Current Push", sessions: 0, revenue: 0 },
    past: { label: "Past Push", sessions: 0, revenue: 0 },
    automation: { label: "Automation Push", sessions: 0, revenue: 0 },
  };

  rows.forEach((row) => {
    if (!base[row.pushBucket]) return;
    base[row.pushBucket].sessions += toNumber(row.sessions) || 0;
    base[row.pushBucket].revenue += toNumber(row.revenue) || 0;
  });

  return Object.values(base).filter((row) => row.sessions !== 0 || row.revenue !== 0);
}

function buildShareData(rows, key) {
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
    <div class="bar-row chart-tip" data-label="${escapeHtml(item.label)}" data-value="${escapeHtml(formatPercent(item.share))}">
      <div class="bar-label">${escapeHtml(item.label)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(item.share * 100, 0)}%"></div></div>
      <div class="bar-value">${escapeHtml(formatPercent(item.share))}</div>
    </div>
  `).join("");
}

function renderTrendCharts() {
  const container = document.getElementById("trendGrid");
  const filteredAllTimeRows = getAllTimeRowsForTrend();
  const weeklySeries = buildWeeklySeries(filteredAllTimeRows);

  if (!weeklySeries.length) {
    container.innerHTML = '<div class="chart-card empty-chart">No weekly trend data available for the current filters.</div>';
    return;
  }

  const metrics = [
    { label: "Weekly Purchasers", key: "purchasers", type: "number" },
    { label: "Weekly Revenue", key: "revenue", type: "currency" },
    { label: "Weekly Sessions", key: "sessions", type: "number" },
    { label: "Weekly UV", key: "users", type: "number" },
  ];

  container.innerHTML = metrics.map((metric) => `
    <div class="chart-card">
      <h3>${escapeHtml(metric.label)}</h3>
      <p class="chart-caption">Grouped by Sunday to Saturday across all available weeks.</p>
      <div class="line-chart">${renderLineSvg(weeklySeries, metric)}</div>
    </div>
  `).join("");
}

function getAllTimeRowsForTrend() {
  return state.classifiedRows.filter((row) => {
    if (row.pushBucket === "future" || row.pushBucket === "invalid") return false;
    return matchesPushFilters(row);
  });
}

function buildWeeklySeries(rows) {
  if (!rows.length) return [];

  const sortedDates = rows
    .map((row) => row.dateObj)
    .filter(Boolean)
    .sort((a, b) => a - b);

  const firstWeekStart = getSunday(sortedDates[0]);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const lastCompleteSaturday = getPreviousOrSameSaturday(today);
  const dataLastSaturday = getSaturday(sortedDates[sortedDates.length - 1]);
  const lastWeekEnd = lastCompleteSaturday < dataLastSaturday
    ? lastCompleteSaturday
    : dataLastSaturday;
  const weekMap = new Map();

  const cursor = new Date(firstWeekStart);
  while (cursor <= lastWeekEnd) {
    const weekStart = new Date(cursor);
    const weekEnd = addDays(weekStart, 6);
    const key = formatDateInput(weekStart);
    weekMap.set(key, {
      label: `${formatMonthDay(weekStart)}-${formatMonthDay(weekEnd)}`,
      weekStart,
      weekEnd,
      sessions: 0,
      users: 0,
      purchasers: 0,
      revenue: 0,
    });
    cursor.setDate(cursor.getDate() + 7);
  }

  rows.forEach((row) => {
    const weekStart = getSunday(row.dateObj);
    const key = formatDateInput(weekStart);
    if (!weekMap.has(key)) return;
    const item = weekMap.get(key);
    item.sessions += toNumber(row.sessions) || 0;
    item.users += toNumber(row.users) || 0;
    item.purchasers += toNumber(row.purchasers) || 0;
    item.revenue += toNumber(row.revenue) || 0;
  });

  return [...weekMap.values()];
}

function renderLineSvg(series, metric) {
  const points = series.map((item) => ({
    label: item.label,
    value: item[metric.key],
  }));

  if (!points.length) return '<div class="empty-chart">No data available.</div>';

  const width = 620;
  const height = 260;
  const padding = { top: 20, right: 20, bottom: 52, left: 58 };
  const innerWidth = width - padding.left - padding.right;
  const innerHeight = height - padding.top - padding.bottom;
  const maxValue = Math.max(...points.map((point) => toNumber(point.value) || 0), 0.01);
  const tickIndexes = buildTickIndexes(points.length, 8);

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
    return `<g class="chart-tip" data-label="${escapeHtml(point.label)}" data-value="${escapeHtml(formatValue(point.value, metric.type))}"><circle cx="${x}" cy="${y}" r="16" fill="transparent"/><circle class="line-dot" cx="${x}" cy="${y}" r="4"/></g>`;
  }).join("");

  const xTicks = tickIndexes.map((index) =>
    `<text class="tick-label" x="${scaleX(index)}" y="${height - 18}" text-anchor="middle">${escapeHtml(points[index].label)}</text>`
  ).join("");

  const yTopLabel = `<text class="tick-label" x="${padding.left}" y="${padding.top - 4}" text-anchor="start">${escapeHtml(formatValue(maxValue, metric.type))}</text>`;

  return `
    <svg class="line-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeHtml(metric.label)}">
      ${gridLines}
      <line class="grid-line" x1="${padding.left}" y1="${padding.top + innerHeight}" x2="${padding.left + innerWidth}" y2="${padding.top + innerHeight}"></line>
      <line class="grid-line" x1="${padding.left}" y1="${padding.top}" x2="${padding.left}" y2="${padding.top + innerHeight}"></line>
      ${yTopLabel}
      <path class="line-path" d="${path}"></path>
      ${dots}
      ${xTicks}
    </svg>
  `;
}

function buildTickIndexes(length, maxTicks) {
  if (length <= maxTicks) return [...Array(length).keys()];
  const step = Math.ceil(length / maxTicks);
  const indexes = [];
  for (let index = 0; index < length; index += step) {
    indexes.push(index);
  }
  if (indexes[indexes.length - 1] !== length - 1) indexes.push(length - 1);
  return indexes;
}

function renderAutomationWeeklyDetail() {
  const container = document.getElementById("automationWeeklyDetail");
  const autoRows = state.classifiedRows.filter((row) => row.pushCategory === "automation");

  if (!autoRows.length) {
    container.innerHTML = '<div class="empty-state">No automation push data available.</div>';
    return;
  }

  const pushNames = [...new Set(autoRows.map((r) => r.pushName))].sort();

  const sortedDates = autoRows
    .map((r) => r.dateObj)
    .filter(Boolean)
    .sort((a, b) => a - b);

  const firstWeekStart = getSunday(sortedDates[0]);
  const lastWeekEnd = getSaturday(sortedDates[sortedDates.length - 1]);

  const weekKeys = [];
  const cursor = new Date(firstWeekStart);
  while (cursor <= lastWeekEnd) {
    weekKeys.push({
      key: formatDateInput(new Date(cursor)),
      start: new Date(cursor),
      end: addDays(new Date(cursor), 6),
      label: `${formatMonthDay(new Date(cursor))}-${formatMonthDay(addDays(new Date(cursor), 6))}`,
    });
    cursor.setDate(cursor.getDate() + 7);
  }

  container.innerHTML = pushNames.map((pushName) => {
    const pushRows = autoRows.filter((r) => r.pushName === pushName);

    const weekData = weekKeys.map((week) => {
      const weekRows = pushRows.filter(
        (r) => r.dateObj >= week.start && r.dateObj <= week.end
      );
      return {
        label: week.label,
        sessions: weekRows.reduce((s, r) => s + (toNumber(r.sessions) || 0), 0),
        users: weekRows.reduce((s, r) => s + (toNumber(r.users) || 0), 0),
        purchasers: weekRows.reduce((s, r) => s + (toNumber(r.purchasers) || 0), 0),
        revenue: weekRows.reduce((s, r) => s + (toNumber(r.revenue) || 0), 0),
      };
    });

    const tableRows = weekData.map((week, i) => {
      const prev = i > 0 ? weekData[i - 1] : null;
      const aov = week.purchasers ? week.revenue / week.purchasers : 0;
      const cvr = week.users ? week.purchasers / week.users : 0;

      const wowCell = (cur, pri) => {
        if (!prev || !pri) return '<td class="wow-neutral">—</td>';
        const pct = (cur - pri) / pri;
        const cls = pct > 0 ? "wow-good" : pct < 0 ? "wow-bad" : "wow-neutral";
        return `<td class="${cls}">${escapeHtml(formatSignedPercent(pct))}</td>`;
      };

      return `<tr>
        <td>${escapeHtml(week.label)}</td>
        <td>${escapeHtml(formatNumber(week.sessions))}</td>
        ${wowCell(week.sessions, prev?.sessions)}
        <td>${escapeHtml(formatNumber(week.users))}</td>
        ${wowCell(week.users, prev?.users)}
        <td>${escapeHtml(formatNumber(week.purchasers))}</td>
        ${wowCell(week.purchasers, prev?.purchasers)}
        <td>${escapeHtml(formatCurrency(week.revenue))}</td>
        ${wowCell(week.revenue, prev?.revenue)}
        <td>${escapeHtml(formatCurrency(aov))}</td>
        <td>${escapeHtml(formatPercent(cvr))}</td>
      </tr>`;
    }).join("");

    return `
      <div class="automation-detail-block">
        <h3 class="automation-detail-title">${escapeHtml(pushName)}</h3>
        <div class="table-wrap">
          <table class="automation-table">
            <thead>
              <tr>
                <th>Week</th>
                <th>Sessions</th><th>WoW</th>
                <th>Users</th><th>WoW</th>
                <th>Purchasers</th><th>WoW</th>
                <th>Revenue</th><th>WoW</th>
                <th>AOV</th>
                <th>CVR</th>
              </tr>
            </thead>
            <tbody>${tableRows}</tbody>
          </table>
        </div>
      </div>
    `;
  }).join("");
}

function renderExcludedTable() {
  const thead = document.querySelector("#excludedTable thead");
  const tbody = document.querySelector("#excludedTable tbody");

  thead.innerHTML = `
    <tr>
      <th>Date</th>
      <th>Push Name</th>
      <th>Push Category</th>
      <th>Push Name Group</th>
      <th>Resolved Push Date</th>
      <th>Reason</th>
    </tr>
  `;

  if (!state.excludedRows.length) {
    tbody.innerHTML = '<tr><td class="empty-cell" colspan="6">No future / invalid manual push names in the current dataset.</td></tr>';
    return;
  }

  tbody.innerHTML = state.excludedRows.map((row) => `
    <tr>
      <td>${escapeHtml(row.date)}</td>
      <td>${escapeHtml(row.pushName)}</td>
      <td>${renderPill(row.pushCategory, "category")}</td>
      <td>${renderPill(row.pushBucket, "bucket")}</td>
      <td>${escapeHtml(row.resolvedManualDate ? formatDateInput(row.resolvedManualDate) : "-")}</td>
      <td>${escapeHtml(row.pushBucket === "future" ? "Manual push name date is later than the selected current period." : "Manual push name could not be parsed into a valid month.day date.")}</td>
    </tr>
  `).join("");
}

function buildWowObject(currentValue, priorValue, type, inverse) {
  const current = toNumber(currentValue) || 0;
  const prior = toNumber(priorValue) || 0;

  if (prior === 0) {
    return {
      pctOnlyLabel: current === 0 ? "0.00%" : "—",
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
  if (!Number.isNaN(tryDate.getTime())) {
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
  if (Number.isNaN(date.getTime())) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function buildSafeDate(year, month, day) {
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

function getSunday(date) {
  const copy = new Date(date);
  copy.setDate(copy.getDate() - copy.getDay());
  return new Date(copy.getFullYear(), copy.getMonth(), copy.getDate());
}

function getSaturday(date) {
  return addDays(getSunday(date), 6);
}

function getPreviousOrSameSaturday(date) {
  const day = date.getDay();
  const distance = (day + 1) % 7;
  return addDays(date, -distance);
}

function formatDateInput(date) {
  if (!date) return "";
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function formatMonthDay(date) {
  if (!date) return "";
  return `${date.getMonth() + 1}.${date.getDate()}`;
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
  if (type === "rate") return formatPercent(value);
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
  document.getElementById("automationWeeklyDetail").innerHTML = `<div class="empty-state">${escapeHtml(message)}</div>`;
  document.querySelector("#excludedTable thead").innerHTML = "";
  document.querySelector("#excludedTable tbody").innerHTML = `<tr><td class="empty-cell">${escapeHtml(message)}</td></tr>`;
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
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
