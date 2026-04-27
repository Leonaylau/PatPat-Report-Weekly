const state = {
  rawRows: [],
  classifiedRows: [],
  overviewRows: [],
  filteredRows: [],
  excludedRows: [],
  productsByTxId: new Map(),
  productCatalog: [],
  currentStart: "",
  currentEnd: "",
  compareStart: "",
  compareEnd: "",
  pushCategoryFilter: "all",
  productTitleFilter: [],
  productSearch: "",
};

const overviewMetrics = [
  { label: "Sessions", key: "sessions", type: "number", inverse: false },
  { label: "Total Purchasers", key: "purchasers", type: "number", inverse: false },
  { label: "Total Revenue", key: "revenue", type: "currency", inverse: false },
  { label: "AOV (客单价)", key: "aov", type: "currency", inverse: false },
  { label: "CVR (Purchasers / Sessions)", key: "cvr", type: "rate", inverse: false },
];

const trendMetrics = [
  { label: "Weekly Purchasers", key: "purchasers", type: "number" },
  { label: "Weekly Revenue", key: "revenue", type: "currency" },
  { label: "Weekly Sessions", key: "sessions", type: "number" },
];

const pushCategoryLabelMap = {
  all: "All",
  manual: "Manual Push (手动PUSH)",
  automation: "Automation Push (自动化PUSH)",
};

const pushBucketLabelMap = {
  all: "All Push Names",
  current: "Current Push",
  past: "Past Push",
  automation: "Automation Push",
  future: "Future Push",
  invalid: "Invalid Manual Name",
};

const SERIES_COLORS = {
  automation: "#c96442",
  manual: "#3b6cb1",
};

const NOT_SET_VALUES = new Set(["", "(not set)"]);

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

  document.getElementById("productTitleFilter")?.addEventListener("change", (e) => {
    const select = e.target;
    state.productTitleFilter = Array.from(select.selectedOptions).map((opt) => opt.value);
    applyFiltersAndRender();
  });

  document.getElementById("productTitleClear")?.addEventListener("click", () => {
    state.productTitleFilter = [];
    syncProductFilterUi();
    applyFiltersAndRender();
  });

  document.getElementById("productSearch")?.addEventListener("input", (e) => {
    state.productSearch = e.target.value.trim().toLowerCase();
    renderProductSection();
  });
}

async function initializePushData() {
  try {
    const mainCandidates = ["./data/current.csv", "./data/raw/current.csv"];
    const productCandidates = [
      "./data/orders_with_date_jan_apr_2026.csv",
      "./data/products.csv",
      "./data/raw/orders_with_date_jan_apr_2026.csv",
      "./data/raw/products.csv",
    ];

    const mainResult = await fetchFirstAvailable(mainCandidates);
    if (!mainResult) throw new Error("No app-push CSV found in ./data/");

    const productsResult = await fetchFirstAvailable(productCandidates);

    const parsed = parsePushCsv(mainResult.text);
    state.rawRows = parsed.rows;

    if (productsResult) {
      const productData = parseProductsCsv(productsResult.text);
      state.productsByTxId = productData.byTxId;
      state.productCatalog = productData.catalog;
    } else {
      state.productsByTxId = new Map();
      state.productCatalog = [];
    }

    attachProductsToRows(state.rawRows, state.productsByTxId);
    populateProductTitleFilterUi(state.productCatalog);

    if (state.rawRows.length) {
      setLatestCompleteWeekRange();
      syncStateToDateInputs();
      const latestDate = state.rawRows.reduce((max, r) => (r.dateObj > max ? r.dateObj : max), state.rawRows[0].dateObj);
      document.getElementById("dataFreshness").textContent = `数据截至: ${formatDateInput(latestDate)}`;
    }

    const productCountText = productsResult
      ? ` · ${state.productCatalog.length} product titles from ${productsResult.path}`
      : " · products file not found (Product titles will display as Unknown)";
    document.getElementById("fileMeta").textContent = `Loaded ${state.rawRows.length} app-push rows from ${mainResult.path}${productCountText}`;
    applyFiltersAndRender();
  } catch (error) {
    document.getElementById("fileMeta").textContent = `Failed to load app-push data: ${error.message}`;
    renderEmptyStates(error.message);
  }
}

async function fetchFirstAvailable(paths) {
  for (const path of paths) {
    try {
      const response = await fetch(path, { cache: "no-store" });
      if (response.ok) {
        const text = await response.text();
        return { text, path };
      }
    } catch {
      // ignore
    }
  }
  return null;
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
    .filter((row) => row.dateObj && row.pushName && !row.isGrandTotal);

  return { rows };
}

function parseProductsCsv(text) {
  const cleanLines = text
    .split(/\r?\n/)
    .filter((line) => line.trim() !== "")
    .filter((line) => !line.trim().startsWith("#"));

  const csvText = cleanLines.join("\n");
  const records = parseCsv(csvText);

  const byTxId = new Map();
  const titleSet = new Set();

  records.forEach((raw) => {
    const txId = normalizeText(pickField(raw, ["Transaction ID", "transaction ID", "TransactionID", "transaction id"]));
    const title = normalizeText(pickField(raw, ["Product title", "Product Title", "product title"]));
    if (!txId || isNotSet(txId)) return;
    if (!title || isNotSet(title)) return;

    const ordersRaw = pickField(raw, ["Orders", "orders", "Order"]);
    const orders = parseNumber(ordersRaw);
    const ordersValue = orders === null ? 1 : orders;

    if (!byTxId.has(txId)) byTxId.set(txId, []);
    const list = byTxId.get(txId);
    const existing = list.find((entry) => entry.productTitle === title);
    if (existing) {
      existing.orders += ordersValue;
    } else {
      list.push({ productTitle: title, orders: ordersValue });
    }
    titleSet.add(title);
  });

  return {
    byTxId,
    catalog: [...titleSet].sort((a, b) => a.localeCompare(b)),
  };
}

function pickField(record, keys) {
  for (const key of keys) {
    if (record[key] !== undefined && record[key] !== null && String(record[key]).trim() !== "") {
      return record[key];
    }
  }
  return "";
}

function isNotSet(value) {
  return NOT_SET_VALUES.has(String(value || "").trim().toLowerCase());
}

function mapPushRow(raw) {
  const dateText = normalizeText(raw["Date"]);
  const transactionIdRaw = normalizeText(pickField(raw, ["Transaction ID", "transaction ID", "TransactionID", "transaction id"]));
  const sourceMediumRaw = normalizeText(pickField(raw, ["Source / medium", "Source/medium", "source / medium"]));
  const manualTermRaw = normalizeText(pickField(raw, ["Manual term", "manual term"]));
  const sessionManualTermRaw = normalizeText(pickField(raw, ["Session Manual term", "Session manual term", "session manual term"]));
  const sessionsRaw = pickField(raw, ["Sessions", "sessions"]);
  const revenueRaw = pickField(raw, ["Total revenue", "total revenue"]);

  const dateObj = parseFlexibleDate(dateText);

  const isGrandTotal = (manualTermRaw + sessionManualTermRaw + transactionIdRaw + dateText)
    .toLowerCase()
    .includes("grand total");

  const transactionId = isNotSet(transactionIdRaw) ? "" : transactionIdRaw;

  // Prefer Manual term as the campaign name; fall back to Session manual term.
  let pushName = isNotSet(manualTermRaw) ? "" : manualTermRaw;
  if (!pushName) pushName = isNotSet(sessionManualTermRaw) ? "" : sessionManualTermRaw;
  if (!pushName) pushName = "(not set)";
  pushName = normalizePushName(pushName);

  const sourceMedium = sourceMediumRaw;
  const sourceCategory = sourceMediumToCategory(sourceMedium);

  return {
    date: dateText,
    dateObj,
    pushName,
    sessionManualTerm: sessionManualTermRaw,
    manualTerm: manualTermRaw,
    sourceMedium,
    sourceCategory,
    transactionId,
    sessions: parseNumber(sessionsRaw),
    purchasers: transactionId ? 1 : 0,
    revenue: parseNumber(revenueRaw),
    isGrandTotal,
    productTitles: [],
    productTitlesLabel: "Unknown",
    productOrdersByTitle: new Map(),
  };
}

function sourceMediumToCategory(sourceMedium) {
  const value = String(sourceMedium || "").trim().toLowerCase();
  if (!value) return null;
  if (value.startsWith("automation")) return "automation";
  if (value.startsWith("manual")) return "manual";
  return null;
}

function attachProductsToRows(rows, productsByTxId) {
  rows.forEach((row) => {
    const txId = row.transactionId;
    if (!txId) {
      row.productTitles = [];
      row.productTitlesLabel = "Unknown";
      row.productOrdersByTitle = new Map();
      return;
    }
    const entries = productsByTxId.get(txId);
    if (!entries || !entries.length) {
      row.productTitles = [];
      row.productTitlesLabel = "Unknown";
      row.productOrdersByTitle = new Map();
      return;
    }
    row.productTitles = entries.map((entry) => entry.productTitle);
    row.productTitlesLabel = row.productTitles.join(" | ");
    row.productOrdersByTitle = new Map(entries.map((entry) => [entry.productTitle, entry.orders]));
  });
}

function populateProductTitleFilterUi(catalog) {
  const select = document.getElementById("productTitleFilter");
  if (!select) return;
  select.innerHTML = catalog.map((title) => `<option value="${escapeHtml(title)}">${escapeHtml(title)}</option>`).join("");
  syncProductFilterUi();
}

function syncProductFilterUi() {
  const select = document.getElementById("productTitleFilter");
  if (!select) return;
  const wanted = new Set(state.productTitleFilter);
  Array.from(select.options).forEach((opt) => {
    opt.selected = wanted.has(opt.value);
  });
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
    return inDateRange && row.pushBucket !== "future" && row.pushBucket !== "invalid" && matchesProductFilter(row);
  });

  state.filteredRows = state.classifiedRows.filter((row) => isRowIncluded(row, currentStart, currentEnd));

  renderAll();
}

function classifyRow(row, currentStart, currentEnd) {
  const pushCategory = row.sourceCategory || classifyPushCategoryFromName(row.pushName);
  const resolvedManualDate = pushCategory === "manual" ? resolvePushNameDate(row.pushName, currentStart, currentEnd) : null;
  const pushBucket = classifyPushBucket(pushCategory, resolvedManualDate, currentStart, currentEnd);

  return {
    ...row,
    pushCategory,
    resolvedManualDate,
    pushBucket,
  };
}

function classifyPushCategoryFromName(pushName) {
  const name = String(pushName || "").trim().toLowerCase();
  if (/pu+sh/i.test(name)) return "manual";
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
    state.pushCategoryFilter === row.pushCategory;

  return categoryMatch && matchesProductFilter(row);
}

function matchesProductFilter(row) {
  if (!state.productTitleFilter.length) return true;
  if (!row.productTitles.length) return false;
  return row.productTitles.some((t) => state.productTitleFilter.includes(t));
}

function renderAll() {
  renderHeaderSummary();
  renderOverview();
  renderCategoryComparison();
  renderProductSection();
  renderContributionCharts();
  renderTrendCharts();
  renderDayOfWeekAnalysis();
  renderAutomationWeeklyDetail();
  renderExcludedTable();
}

function renderHeaderSummary() {
  const productLabel = state.productTitleFilter.length
    ? `${state.productTitleFilter.length} selected`
    : "All Products";
  document.getElementById("headerSummary").textContent =
    `Current Period: ${state.currentStart || "--"} ~ ${state.currentEnd || "--"} ｜ Compare Period: ${state.compareStart || "--"} ~ ${state.compareEnd || "--"} ｜ Push Category: ${pushCategoryLabelMap[state.pushCategoryFilter]} ｜ Product Title: ${productLabel}`;
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
        return inRange && row.pushBucket !== "future" && row.pushBucket !== "invalid" && matchesProductFilter(row);
      })
    : [];
  const compareAgg = aggregateRows(compareRows);

  context.textContent =
    `Current: ${state.currentStart} ~ ${state.currentEnd} · Compare: ${state.compareStart} ~ ${state.compareEnd}`;

  container.innerHTML = overviewMetrics.map((metric) => {
    const wow = buildWowObject(currentAgg[metric.key], compareAgg[metric.key], metric.type, metric.inverse);
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

function aggregateRows(rows) {
  const agg = rows.reduce(
    (acc, row) => {
      acc.sessions += toNumber(row.sessions) || 0;
      acc.purchasers += toNumber(row.purchasers) || 0;
      acc.revenue += toNumber(row.revenue) || 0;
      return acc;
    },
    { sessions: 0, purchasers: 0, revenue: 0 }
  );
  agg.aov = agg.purchasers ? agg.revenue / agg.purchasers : 0;
  agg.cvr = agg.sessions ? agg.purchasers / agg.sessions : 0;
  return agg;
}

function renderCategoryComparison() {
  const cardContainer = document.getElementById("categoryCardGrid");
  const trendContainer = document.getElementById("categoryTrendGrid");
  const context = document.getElementById("categoryContext");

  context.textContent =
    `Current: ${state.currentStart} ~ ${state.currentEnd} · Push Category Filter: ${pushCategoryLabelMap[state.pushCategoryFilter]}`;

  const inPeriodRows = state.filteredRows;

  const allGroups = [
    { key: "automation", label: "Automation Push (自动化PUSH)", color: SERIES_COLORS.automation },
    { key: "manual", label: "Manual Push (手动PUSH)", color: SERIES_COLORS.manual },
  ];
  const groups = state.pushCategoryFilter === "all"
    ? allGroups
    : allGroups.filter((g) => g.key === state.pushCategoryFilter);

  if (!inPeriodRows.length) {
    cardContainer.innerHTML = '<div class="empty-state">No app-push rows in the selected date range.</div>';
    trendContainer.innerHTML = '<div class="chart-card empty-chart">No weekly trend data available.</div>';
    return;
  }

  cardContainer.innerHTML = groups.map((group) => {
    const groupRows = inPeriodRows.filter((row) => row.pushCategory === group.key);
    const agg = aggregateRows(groupRows);
    const cards = overviewMetrics.map((metric) => `
      <article class="metric-card">
        <p class="metric-label">${escapeHtml(metric.label)}</p>
        <p class="metric-value">${escapeHtml(formatValue(agg[metric.key], metric.type))}</p>
      </article>
    `).join("");
    return `
      <div class="category-block">
        <div class="category-block-head">
          <span class="category-swatch" style="background:${group.color}"></span>
          <h3 class="category-block-title">${escapeHtml(group.label)}</h3>
          <span class="category-block-meta">${formatNumber(groupRows.length)} rows</span>
        </div>
        <div class="overview-grid">${cards}</div>
      </div>
    `;
  }).join("");

  const allTimeRows = state.classifiedRows.filter((row) => {
    if (row.pushBucket === "future" || row.pushBucket === "invalid") return false;
    return matchesProductFilter(row);
  });

  const weekFrame = buildWeekFrame(allTimeRows);
  if (!weekFrame.length) {
    trendContainer.innerHTML = '<div class="chart-card empty-chart">No weekly trend data available.</div>';
    return;
  }

  const seriesByGroup = groups.map((group) => ({
    label: group.label,
    color: group.color,
    weekly: aggregateWeekly(weekFrame, allTimeRows.filter((row) => row.pushCategory === group.key)),
  }));

  const visibleWeeks = Math.min(weekFrame.length, 4);
  const sliceStart = weekFrame.length - visibleWeeks;
  const xLabels = weekFrame.slice(sliceStart).map((w) => w.label);

  trendContainer.innerHTML = trendMetrics.map((metric) => {
    const series = seriesByGroup.map((group) => ({
      label: group.label,
      color: group.color,
      points: group.weekly.slice(sliceStart).map((w, i) => ({ label: xLabels[i], value: w[metric.key] })),
    }));
    return `
      <div class="chart-card">
        <h3>${escapeHtml(metric.label)}</h3>
        <p class="chart-caption">Sunday to Saturday · last ${visibleWeeks} week(s).</p>
        <div class="legend-row">${series.map((s) => `<span class="legend-item"><span class="legend-swatch" style="background:${s.color}"></span>${escapeHtml(s.label)}</span>`).join("")}</div>
        <div class="line-chart">${renderMultiLineSvg(series, metric)}</div>
      </div>
    `;
  }).join("");
}

function buildWeekFrame(rows) {
  if (!rows.length) return [];
  const dates = rows.map((row) => row.dateObj).filter(Boolean).sort((a, b) => a - b);
  const firstWeekStart = getSunday(dates[0]);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const lastCompleteSaturday = getPreviousOrSameSaturday(today);
  const dataLastSaturday = getSaturday(dates[dates.length - 1]);
  const lastWeekEnd = lastCompleteSaturday < dataLastSaturday ? lastCompleteSaturday : dataLastSaturday;

  const frame = [];
  const cursor = new Date(firstWeekStart);
  while (cursor <= lastWeekEnd) {
    const weekStart = new Date(cursor);
    const weekEnd = addDays(weekStart, 6);
    frame.push({
      key: formatDateInput(weekStart),
      weekStart,
      weekEnd,
      label: `${formatMonthDay(weekStart)}-${formatMonthDay(weekEnd)}`,
    });
    cursor.setDate(cursor.getDate() + 7);
  }
  return frame;
}

function aggregateWeekly(frame, rows) {
  const indexByKey = new Map(frame.map((w, i) => [w.key, i]));
  const buckets = frame.map((w) => ({
    label: w.label,
    weekStart: w.weekStart,
    weekEnd: w.weekEnd,
    sessions: 0,
    purchasers: 0,
    revenue: 0,
  }));
  rows.forEach((row) => {
    const sunday = getSunday(row.dateObj);
    const key = formatDateInput(sunday);
    const idx = indexByKey.get(key);
    if (idx === undefined) return;
    const b = buckets[idx];
    b.sessions += toNumber(row.sessions) || 0;
    b.purchasers += toNumber(row.purchasers) || 0;
    b.revenue += toNumber(row.revenue) || 0;
  });
  return buckets;
}

function renderProductSection() {
  const tableBody = document.querySelector("#productTable tbody");
  const tableHead = document.querySelector("#productTable thead");
  const rankingContainer = document.getElementById("productRankingChart");
  const context = document.getElementById("productContext");

  const rolled = rollupByProduct(state.filteredRows);
  const search = state.productSearch;
  const visible = search ? rolled.filter((p) => p.productTitle.toLowerCase().includes(search)) : rolled;

  context.textContent = `Current: ${state.currentStart} ~ ${state.currentEnd} · ${visible.length} product titles in scope (per-product totals may exceed grand totals because revenue / sessions are not split across multi-product transactions).`;

  tableHead.innerHTML = `
    <tr>
      <th>Product Title</th>
      <th>Sessions</th>
      <th>Purchasers</th>
      <th>Revenue</th>
      <th>AOV</th>
      <th>CVR (Purchasers / Sessions)</th>
    </tr>
  `;

  if (!visible.length) {
    tableBody.innerHTML = '<tr><td class="empty-cell" colspan="6">No products match the current filters.</td></tr>';
    rankingContainer.classList.add("empty-chart");
    rankingContainer.innerHTML = "No product revenue available.";
    return;
  }

  tableBody.innerHTML = visible.map((p) => `
    <tr>
      <td>${escapeHtml(p.productTitle)}</td>
      <td>${escapeHtml(formatNumber(p.sessions))}</td>
      <td>${escapeHtml(formatNumber(p.purchasers))}</td>
      <td>${escapeHtml(formatCurrency(p.revenue))}</td>
      <td>${escapeHtml(formatCurrency(p.aov))}</td>
      <td>${escapeHtml(formatPercent(p.cvr))}</td>
    </tr>
  `).join("");

  const top = visible.slice().sort((a, b) => b.revenue - a.revenue).slice(0, 10);
  renderRankingChart(rankingContainer, top.map((p) => ({ pushName: p.productTitle, revenue: p.revenue })));
}

function rollupByProduct(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const titles = row.productTitles.length ? row.productTitles : ["Unknown"];
    titles.forEach((title) => {
      if (!map.has(title)) {
        map.set(title, { productTitle: title, sessions: 0, purchasers: 0, revenue: 0 });
      }
      const item = map.get(title);
      item.sessions += toNumber(row.sessions) || 0;
      item.revenue += toNumber(row.revenue) || 0;
      const orders = row.productOrdersByTitle.get(title);
      if (typeof orders === "number" && Number.isFinite(orders)) {
        item.purchasers += orders;
      } else if (title === "Unknown") {
        item.purchasers += toNumber(row.purchasers) || 0;
      }
    });
  });
  return [...map.values()]
    .map((p) => ({
      ...p,
      aov: p.purchasers ? p.revenue / p.purchasers : 0,
      cvr: p.sessions ? p.purchasers / p.sessions : 0,
    }))
    .sort((a, b) => b.revenue - a.revenue);
}

function renderContributionCharts() {
  const grouped = groupByPushBucket(state.filteredRows);
  renderBarChart(document.getElementById("sessionsShareChart"), buildShareData(grouped, "sessions"));
  renderBarChart(document.getElementById("revenueShareChart"), buildShareData(grouped, "revenue"));
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

function renderRankingChart(container, items) {
  if (!items.length) {
    container.classList.add("empty-chart");
    container.innerHTML = "No data for ranking.";
    return;
  }
  const maxRevenue = Math.max(...items.map((r) => toNumber(r.revenue) || 0), 0.01);
  container.classList.remove("empty-chart");
  container.innerHTML = items.map((item, i) => {
    const rev = toNumber(item.revenue) || 0;
    const pct = (rev / maxRevenue) * 100;
    return `
    <div class="bar-row chart-tip" data-label="${escapeHtml(item.pushName)}" data-value="${escapeHtml(formatCurrency(rev))}">
      <div class="bar-label">${i + 1}. ${escapeHtml(item.pushName)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(pct, 0)}%"></div></div>
      <div class="bar-value">${escapeHtml(formatCurrency(rev))}</div>
    </div>`;
  }).join("");
}

function renderTrendCharts() {
  const container = document.getElementById("trendGrid");
  const filteredAllTimeRows = getAllTimeRowsForTrend();
  const frame = buildWeekFrame(filteredAllTimeRows);
  const weekly = aggregateWeekly(frame, filteredAllTimeRows);
  const visibleWeeks = Math.min(weekly.length, 4);
  const sliceStart = weekly.length - visibleWeeks;
  const weeklySeries = weekly.slice(sliceStart);

  if (!weeklySeries.length) {
    container.innerHTML = '<div class="chart-card empty-chart">No weekly trend data available for the current filters.</div>';
    return;
  }

  container.innerHTML = trendMetrics.map((metric) => `
    <div class="chart-card">
      <h3>${escapeHtml(metric.label)}</h3>
      <p class="chart-caption">Grouped by Sunday to Saturday across the latest ${visibleWeeks} week(s).</p>
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

function renderLineSvg(series, metric) {
  const points = series.map((item) => ({ label: item.label, value: item[metric.key] }));
  if (!points.length) return '<div class="empty-chart">No data available.</div>';
  return renderMultiLineSvg([{ label: metric.label, color: SERIES_COLORS.automation, points }], metric);
}

function renderMultiLineSvg(seriesList, metric) {
  const seriesWithData = seriesList.filter((s) => s.points && s.points.length);
  if (!seriesWithData.length) return '<div class="empty-chart">No data available.</div>';

  const length = seriesWithData[0].points.length;
  const width = 820;
  const height = 260;
  const padding = { top: 20, right: 60, bottom: 52, left: 58 };
  const innerWidth = width - padding.left - padding.right;
  const innerHeight = height - padding.top - padding.bottom;
  const allValues = seriesWithData.flatMap((s) => s.points.map((p) => toNumber(p.value) || 0));
  const maxValue = Math.max(...allValues, 0.01);
  const tickIndexes = buildTickIndexes(length, 8);

  const scaleX = (index) =>
    padding.left + (length === 1 ? innerWidth / 2 : (index / (length - 1)) * innerWidth);
  const scaleY = (value) =>
    padding.top + innerHeight - ((toNumber(value) || 0) / maxValue) * innerHeight;

  const gridLines = [0.25, 0.5, 0.75, 1].map((factor) => {
    const y = padding.top + innerHeight - innerHeight * factor;
    return `<line class="grid-line" x1="${padding.left}" y1="${y}" x2="${padding.left + innerWidth}" y2="${y}"></line>`;
  }).join("");

  const seriesSvg = seriesWithData.map((s) => {
    const path = s.points
      .map((point, index) => `${index === 0 ? "M" : "L"}${scaleX(index)} ${scaleY(point.value)}`)
      .join(" ");
    const dots = s.points.map((point, index) => {
      const x = scaleX(index);
      const y = scaleY(point.value);
      return `<g class="chart-tip" data-label="${escapeHtml(`${s.label} · ${point.label}`)}" data-value="${escapeHtml(formatValue(point.value, metric.type))}"><circle cx="${x}" cy="${y}" r="16" fill="transparent"/><circle class="line-dot" cx="${x}" cy="${y}" r="4" style="fill:${s.color}"/></g>`;
    }).join("");
    return `<path class="line-path" d="${path}" style="stroke:${s.color}"></path>${dots}`;
  }).join("");

  const xTicks = tickIndexes.map((index) =>
    `<text class="tick-label" x="${scaleX(index)}" y="${height - 18}" text-anchor="middle">${escapeHtml(seriesWithData[0].points[index].label)}</text>`
  ).join("");

  const yTopLabel = `<text class="tick-label" x="${padding.left}" y="${padding.top - 4}" text-anchor="start">${escapeHtml(formatValue(maxValue, metric.type))}</text>`;

  return `
    <svg class="line-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="${escapeHtml(metric.label)}">
      ${gridLines}
      <line class="grid-line" x1="${padding.left}" y1="${padding.top + innerHeight}" x2="${padding.left + innerWidth}" y2="${padding.top + innerHeight}"></line>
      <line class="grid-line" x1="${padding.left}" y1="${padding.top}" x2="${padding.left}" y2="${padding.top + innerHeight}"></line>
      ${yTopLabel}
      ${seriesSvg}
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

function renderDayOfWeekAnalysis() {
  const container = document.getElementById("dayOfWeekGrid");
  const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

  if (!state.filteredRows.length) {
    container.innerHTML = '<div class="empty-state">No data for day-of-week analysis.</div>';
    return;
  }

  const buckets = Array.from({ length: 7 }, () => ({ sessions: 0, purchasers: 0, revenue: 0 }));
  state.filteredRows.forEach((row) => {
    const d = row.dateObj.getDay();
    buckets[d].sessions += toNumber(row.sessions) || 0;
    buckets[d].purchasers += toNumber(row.purchasers) || 0;
    buckets[d].revenue += toNumber(row.revenue) || 0;
  });

  const bestDay = buckets.reduce((bi, b, i, a) => b.revenue > a[bi].revenue ? i : bi, 0);

  container.innerHTML = buckets.map((b, i) => {
    const aov = b.purchasers ? b.revenue / b.purchasers : 0;
    const cvr = b.sessions ? b.purchasers / b.sessions : 0;
    const highlight = i === bestDay ? " dow-best" : "";
    return `
      <article class="metric-card${highlight}">
        <p class="metric-label">${dayNames[i]}</p>
        <p class="metric-value">${escapeHtml(formatCurrency(b.revenue))}</p>
        <p class="metric-compare">${escapeHtml(formatNumber(b.sessions))} sessions · ${escapeHtml(formatNumber(b.purchasers))} purch</p>
        <p class="metric-compare">AOV ${escapeHtml(formatCurrency(aov))} · CVR ${escapeHtml(formatPercent(cvr))}</p>
      </article>
    `;
  }).join("");
}

function renderAutomationWeeklyDetail() {
  const container = document.getElementById("automationWeeklyDetail");
  const autoRows = state.classifiedRows.filter((row) => row.pushCategory === "automation" && matchesProductFilter(row));

  if (!autoRows.length) {
    container.innerHTML = '<div class="empty-state">No automation push data available.</div>';
    return;
  }

  const pushNames = [...new Set(autoRows.map((r) => r.pushName))].sort();

  const sortedDates = autoRows.map((r) => r.dateObj).filter(Boolean).sort((a, b) => a - b);
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
      const weekRows = pushRows.filter((r) => r.dateObj >= week.start && r.dateObj <= week.end);
      return {
        label: week.label,
        sessions: weekRows.reduce((s, r) => s + (toNumber(r.sessions) || 0), 0),
        purchasers: weekRows.reduce((s, r) => s + (toNumber(r.purchasers) || 0), 0),
        revenue: weekRows.reduce((s, r) => s + (toNumber(r.revenue) || 0), 0),
      };
    });

    const bestIdx = weekData.reduce((bi, w, i, a) => w.revenue > a[bi].revenue ? i : bi, 0);
    const worstIdx = weekData.reduce((wi, w, i, a) => w.revenue < a[wi].revenue ? i : wi, 0);
    const hasData = weekData.some((w) => w.revenue > 0);

    const tableRows = weekData.map((week, i) => {
      const prev = i > 0 ? weekData[i - 1] : null;
      const aov = week.purchasers ? week.revenue / week.purchasers : 0;
      const cvr = week.sessions ? week.purchasers / week.sessions : 0;
      const rowCls = hasData && i === bestIdx ? ' class="week-best"' : hasData && i === worstIdx && bestIdx !== worstIdx ? ' class="week-worst"' : "";

      const wowCell = (cur, pri) => {
        if (!prev || !pri) return '<td class="wow-neutral">—</td>';
        const pct = (cur - pri) / pri;
        const cls = pct > 0 ? "wow-good" : pct < 0 ? "wow-bad" : "wow-neutral";
        return `<td class="${cls}">${escapeHtml(formatSignedPercent(pct))}</td>`;
      };

      return `<tr${rowCls}>
        <td>${escapeHtml(week.label)}</td>
        <td>${escapeHtml(formatNumber(week.sessions))}</td>
        ${wowCell(week.sessions, prev?.sessions)}
        <td>${escapeHtml(formatNumber(week.purchasers))}</td>
        ${wowCell(week.purchasers, prev?.purchasers)}
        <td>${escapeHtml(formatCurrency(week.revenue))}</td>
        ${wowCell(week.revenue, prev?.revenue)}
        <td>${escapeHtml(formatCurrency(aov))}</td>
        <td>${escapeHtml(formatPercent(cvr))}</td>
      </tr>`;
    }).join("");

    const totals = weekData.reduce((t, w) => {
      t.sessions += w.sessions;
      t.purchasers += w.purchasers; t.revenue += w.revenue;
      return t;
    }, { sessions: 0, purchasers: 0, revenue: 0 });
    const totalAov = totals.purchasers ? totals.revenue / totals.purchasers : 0;
    const totalCvr = totals.sessions ? totals.purchasers / totals.sessions : 0;
    const summaryRow = `<tr class="summary-row">
      <td>Total</td>
      <td>${escapeHtml(formatNumber(totals.sessions))}</td><td></td>
      <td>${escapeHtml(formatNumber(totals.purchasers))}</td><td></td>
      <td>${escapeHtml(formatCurrency(totals.revenue))}</td><td></td>
      <td>${escapeHtml(formatCurrency(totalAov))}</td>
      <td>${escapeHtml(formatPercent(totalCvr))}</td>
    </tr>`;

    return `
      <div class="automation-detail-block">
        <h3 class="automation-detail-title">${escapeHtml(pushName)}</h3>
        <div class="table-wrap">
          <table class="automation-table">
            <thead>
              <tr>
                <th>Week</th>
                <th>Sessions</th><th>WoW</th>
                <th>Purchasers</th><th>WoW</th>
                <th>Revenue</th><th>WoW</th>
                <th>AOV</th>
                <th>CVR</th>
              </tr>
            </thead>
            <tbody>${tableRows}${summaryRow}</tbody>
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

  const sample = state.excludedRows.slice(0, 200);
  tbody.innerHTML = sample.map((row) => `
    <tr>
      <td>${escapeHtml(row.date)}</td>
      <td>${escapeHtml(row.pushName)}</td>
      <td>${renderPill(row.pushCategory, "category")}</td>
      <td>${renderPill(row.pushBucket, "bucket")}</td>
      <td>${escapeHtml(row.resolvedManualDate ? formatDateInput(row.resolvedManualDate) : "-")}</td>
      <td>${escapeHtml(row.pushBucket === "future" ? "Manual push name date is later than the selected current period." : "Manual push name could not be parsed into a valid month.day date.")}</td>
    </tr>
  `).join("") + (state.excludedRows.length > sample.length ? `<tr><td class="empty-cell" colspan="6">… ${state.excludedRows.length - sample.length} more excluded rows truncated for display.</td></tr>` : "");
}

function renderPill(value, type) {
  const labelMap = type === "category" ? pushCategoryLabelMap : pushBucketLabelMap;
  const className = `push-type-pill ${type === "category" ? `pill-category-${value}` : `pill-bucket-${value}`}`;
  return `<span class="${className}">${escapeHtml(labelMap[value] || value)}</span>`;
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
  const catCard = document.getElementById("categoryCardGrid");
  if (catCard) catCard.innerHTML = `<div class="empty-state">${escapeHtml(message)}</div>`;
  const catTrend = document.getElementById("categoryTrendGrid");
  if (catTrend) catTrend.innerHTML = `<div class="chart-card empty-chart">${escapeHtml(message)}</div>`;
  const productHead = document.querySelector("#productTable thead");
  if (productHead) productHead.innerHTML = "";
  const productBody = document.querySelector("#productTable tbody");
  if (productBody) productBody.innerHTML = `<tr><td class="empty-cell">${escapeHtml(message)}</td></tr>`;
  const productRanking = document.getElementById("productRankingChart");
  if (productRanking) productRanking.innerHTML = escapeHtml(message);
  document.getElementById("sessionsShareChart").innerHTML = message;
  document.getElementById("revenueShareChart").innerHTML = message;
  document.getElementById("trendGrid").innerHTML = `<div class="chart-card empty-chart">${escapeHtml(message)}</div>`;
  document.getElementById("dayOfWeekGrid").innerHTML = `<div class="empty-state">${escapeHtml(message)}</div>`;
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
