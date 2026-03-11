const METRIC_CONFIG = [
  {
    key: "clicks",
    label: "点击量",
    aliases: ["点击量", "clicks", "click"]
  },
  {
    key: "dpv",
    label: "总 DPV",
    aliases: ["总 dpv", "总dpv", "dpv", "total dpv"]
  },
  {
    key: "atc",
    label: "ATC 总计",
    aliases: ["atc 总计", "atc总计", "atc", "add to cart"]
  },
  {
    key: "revenue",
    label: "购买总额",
    aliases: ["购买总额", "购买金额", "purchase amount", "revenue", "sales amount"]
  },
  {
    key: "units",
    label: "商品销量总计",
    aliases: ["商品销量总计", "商品销量", "销量总计", "units sold", "units"]
  },
  {
    key: "bonus",
    label: "品牌引流奖励计划",
    aliases: ["品牌引流奖励计划", "品牌引流奖励", "奖励计划", "bonus"]
  }
];

const FIELD_CONFIG = [
  {
    key: "adGroup",
    label: "广告组",
    aliases: ["广告组", "ad group", "adgroup"]
  },
  {
    key: "channel",
    label: "Channel",
    aliases: ["channel", "渠道"]
  },
  {
    key: "publisher",
    label: "出版商",
    aliases: ["出版商", "publisher", "账号"]
  }
];

const DEFAULT_SOURCE = {
  name: "红人数据追踪.xlsx",
  path: "./excel/红人数据追踪.xlsx"
};

const dashboardState = {
  sourceFile: { ...DEFAULT_SOURCE },
  sheets: [],
  charts: new Map()
};

const nodes = {
  sourceFileName: document.getElementById("sourceFileName"),
  sourceStatus: document.getElementById("sourceStatus"),
  reloadDataBtn: document.getElementById("reloadDataBtn"),
  sheetBoard: document.getElementById("sheetBoard"),
  loadingPanel: document.getElementById("loadingPanel")
};

bindIfPresent(nodes.reloadDataBtn, "click", () => {
  loadDashboardFromRepo();
});

loadDashboardFromRepo();

async function loadDashboardFromRepo() {
  showLoading("正在加载红人数据...");
  destroyAllCharts();

  if (typeof window.XLSX === "undefined") {
    showError("Excel 解析库加载失败，请刷新页面后重试。");
    return;
  }

  if (typeof window.Chart !== "function") {
    showError("图表库加载失败，请刷新页面后重试。");
    return;
  }

  try {
    const preferredSource = await resolveSourceFileInfo();
    let workbookBundle = null;

    try {
      workbookBundle = await loadWorkbookBySource(preferredSource);
    } catch (primaryError) {
      if (preferredSource.path !== DEFAULT_SOURCE.path) {
        workbookBundle = await loadWorkbookBySource(DEFAULT_SOURCE);
      } else {
        throw primaryError;
      }
    }

    dashboardState.sourceFile = workbookBundle.source;
    setText(nodes.sourceFileName, workbookBundle.source.name);

    const sheetModels = buildSheetModels(workbookBundle.workbook);
    if (!sheetModels.length) {
      throw new Error("文件内没有可用的 Sheet 数据。");
    }

    dashboardState.sheets = sheetModels;
    renderSheetBoard();
    showBoard();
    setStatus(`已加载 ${workbookBundle.source.name}｜共 ${sheetModels.length} 个 Sheet`);
  } catch (error) {
    showError(`加载失败：${error.message}`);
  }
}

async function resolveSourceFileInfo() {
  try {
    const response = await fetch("./excel/manifest.json", { cache: "no-store" });
    if (!response.ok) {
      return { ...DEFAULT_SOURCE };
    }

    const files = await parseManifestResponse(response);
    if (!files.length) {
      return { ...DEFAULT_SOURCE };
    }

    const preferred = pickPreferredSource(files);
    if (!preferred) {
      return { ...DEFAULT_SOURCE };
    }

    return preferred;
  } catch (_) {
    return { ...DEFAULT_SOURCE };
  }
}

function pickPreferredSource(files) {
  const normalizedFiles = files.map((item) => ({
    name: item.name || getFileBaseName(item.path),
    path: normalizeFilePath(item.path)
  }));

  const rules = [
    (text) => text.includes("红人数据追踪"),
    (text) => text.includes("红人") && text.includes("追踪"),
    (text) => text.includes("红人"),
    (text) => text.includes("influencer")
  ];

  for (const rule of rules) {
    const found = normalizedFiles.find((item) => rule(`${item.name} ${item.path}`.toLowerCase()));
    if (found) {
      return found;
    }
  }

  return normalizedFiles[0] || null;
}

async function loadWorkbookBySource(source) {
  const normalizedPath = normalizeFilePath(source.path);
  const fetched = await fetchFileResponse(normalizedPath);
  const workbook = await readWorkbookFromResponse(fetched.response, fetched.path);

  return {
    source: {
      name: source.name || getFileBaseName(normalizedPath),
      path: fetched.path
    },
    workbook
  };
}

async function fetchFileResponse(path) {
  const candidates = Array.from(new Set([path, encodeURI(path)]));
  let lastError = null;

  for (const candidate of candidates) {
    try {
      const response = await fetch(candidate, { cache: "no-store" });
      if (response.ok) {
        return { response, path: candidate };
      }
      lastError = new Error(`HTTP ${response.status}`);
    } catch (error) {
      lastError = error;
    }
  }

  throw new Error(lastError?.message || "读取仓库文件失败");
}

async function readWorkbookFromResponse(response, filePath) {
  const lowerPath = String(filePath || "").toLowerCase();
  if (lowerPath.endsWith(".csv")) {
    const text = await response.text();
    return XLSX.read(text, { type: "string" });
  }

  const arrayBuffer = await response.arrayBuffer();
  return XLSX.read(arrayBuffer, { type: "array" });
}

function buildSheetModels(workbook) {
  return workbook.SheetNames.map((sheetName, index) => {
    const sheet = workbook.Sheets[sheetName];
    const parsed = parseSheetRows(sheet);
    if (!parsed.rows.length) {
      return null;
    }

    const columnMap = resolveColumns(parsed.headers);
    const normalizedRows = normalizeRows(parsed.rows, parsed.headers, columnMap);
    const aggregatedRows = aggregateByAdGroup(normalizedRows);

    if (!aggregatedRows.length) {
      return null;
    }

    return {
      id: `sheet-${index + 1}`,
      name: sheetName,
      rowCount: aggregatedRows.length,
      rows: aggregatedRows,
      lineMetric: METRIC_CONFIG[0].key,
      pieMetric: METRIC_CONFIG[1].key
    };
  }).filter(Boolean);
}

function resolveColumns(headers) {
  const map = {};
  const headerMap = new Map(headers.map((header) => [normalizeHeader(header), header]));

  [...METRIC_CONFIG, ...FIELD_CONFIG].forEach((item) => {
    let matched = "";

    for (const alias of item.aliases) {
      const exact = headerMap.get(normalizeHeader(alias));
      if (exact) {
        matched = exact;
        break;
      }
    }

    if (!matched) {
      const fuzzy = headers.find((header) => {
        const normalized = normalizeHeader(header);
        return item.aliases.some((alias) => normalized.includes(normalizeHeader(alias)));
      });
      matched = fuzzy || "";
    }

    map[item.key] = matched;
  });

  if (!map.adGroup) {
    map.adGroup = headers[0] || "";
  }

  return map;
}

function normalizeRows(rows, headers, columnMap) {
  return rows.map((row, index) => {
    const fallbackName = firstNonEmptyCell(row, headers) || `未命名广告组 ${index + 1}`;
    const adGroup = columnMap.adGroup ? String(row[columnMap.adGroup] || "").trim() || fallbackName : fallbackName;
    const channel = columnMap.channel ? String(row[columnMap.channel] || "").trim() : "";
    const publisher = columnMap.publisher ? String(row[columnMap.publisher] || "").trim() : "";

    const metrics = {};
    METRIC_CONFIG.forEach((item) => {
      const header = columnMap[item.key];
      metrics[item.key] = header ? toNumber(row[header]) : 0;
    });

    return {
      adGroup,
      channel,
      publisher,
      ...metrics
    };
  });
}

function firstNonEmptyCell(row, headers) {
  for (const header of headers) {
    const value = String(row[header] ?? "").trim();
    if (value) {
      return value;
    }
  }
  return "";
}

function aggregateByAdGroup(rows) {
  const grouped = new Map();

  rows.forEach((row, index) => {
    const key = row.adGroup || `未命名广告组 ${index + 1}`;
    const current = grouped.get(key) || {
      adGroup: key,
      channel: row.channel || "",
      publisher: row.publisher || "",
      clicks: 0,
      dpv: 0,
      atc: 0,
      revenue: 0,
      units: 0,
      bonus: 0
    };

    METRIC_CONFIG.forEach((item) => {
      current[item.key] += toNumber(row[item.key]);
    });

    if (!current.channel && row.channel) {
      current.channel = row.channel;
    }
    if (!current.publisher && row.publisher) {
      current.publisher = row.publisher;
    }

    grouped.set(key, current);
  });

  return Array.from(grouped.values()).sort((a, b) => a.adGroup.localeCompare(b.adGroup, "zh-CN"));
}

function renderSheetBoard() {
  if (!nodes.sheetBoard) {
    return;
  }

  nodes.sheetBoard.innerHTML = dashboardState.sheets.map((sheet) => buildSheetCardMarkup(sheet)).join("");

  dashboardState.sheets.forEach((sheet) => {
    const card = getSheetCard(sheet.id);
    if (!card) {
      return;
    }

    const lineSelect = card.querySelector(".line-metric-select");
    const pieSelect = card.querySelector(".pie-metric-select");

    if (lineSelect) {
      lineSelect.value = sheet.lineMetric;
      lineSelect.addEventListener("change", () => {
        sheet.lineMetric = lineSelect.value;
        renderSheetModule(sheet);
      });
    }

    if (pieSelect) {
      pieSelect.value = sheet.pieMetric;
      pieSelect.addEventListener("change", () => {
        sheet.pieMetric = pieSelect.value;
        renderSheetModule(sheet);
      });
    }

    renderSheetModule(sheet);
  });
}

function buildSheetCardMarkup(sheet) {
  const totals = calculateTotals(sheet.rows);
  const metricOptions = METRIC_CONFIG.map(
    (item) => `<option value="${item.key}">${escapeHtml(item.label)}</option>`
  ).join("");

  const summaryChips = METRIC_CONFIG.map((item) => {
    return `
      <article class="summary-chip">
        <p class="chip-label">${escapeHtml(item.label)}</p>
        <p class="chip-value">${formatMetric(item.key, totals[item.key])}</p>
      </article>
    `;
  }).join("");

  return `
    <article class="sheet-card" data-sheet-id="${escapeHtml(sheet.id)}">
      <header class="sheet-head">
        <div>
          <h2>${escapeHtml(sheet.name)}</h2>
          <p class="sheet-meta">广告组数量：${sheet.rowCount}</p>
        </div>
      </header>

      <section class="sheet-summary">
        ${summaryChips}
      </section>

      <section class="sheet-controls">
        <label>
          折线图变量
          <select class="line-metric-select">${metricOptions}</select>
        </label>
        <label>
          饼状图变量
          <select class="pie-metric-select">${metricOptions}</select>
        </label>
      </section>

      <section class="sheet-chart-grid">
        <article class="chart-shell">
          <h3 class="line-title">折线图</h3>
          <p class="chart-note">横坐标：广告组名称，纵坐标：变量数值</p>
          <canvas class="line-canvas"></canvas>
        </article>
        <article class="chart-shell">
          <h3 class="pie-title">饼状图</h3>
          <p class="chart-note">按广告组占比显示（Top 8 + 其他）</p>
          <canvas class="pie-canvas"></canvas>
        </article>
      </section>

      <section class="sheet-table">
        <div class="table-header-line">
          <h3>广告组明细</h3>
          <p class="table-note"></p>
        </div>

        <div class="mini-table-wrap">
          <table class="mini-table">
            ${tableHeadMarkup()}
            <tbody class="top-body"></tbody>
          </table>
        </div>

        <details class="rest-block hidden">
          <summary class="rest-summary"></summary>
          <div class="rest-table-scroll">
            <table class="mini-table">
              ${tableHeadMarkup()}
              <tbody class="rest-body"></tbody>
            </table>
          </div>
        </details>
      </section>
    </article>
  `;
}

function tableHeadMarkup() {
  return `
    <thead>
      <tr>
        <th>广告组</th>
        <th>Channel</th>
        <th>出版商</th>
        <th>点击量</th>
        <th>总 DPV</th>
        <th>ATC 总计</th>
        <th>购买总额</th>
        <th>商品销量总计</th>
        <th>品牌引流奖励计划</th>
      </tr>
    </thead>
  `;
}

function renderSheetModule(sheet) {
  const card = getSheetCard(sheet.id);
  if (!card) {
    return;
  }

  const sortedRows = sortRowsByMetric(sheet.rows, sheet.lineMetric);
  const lineTitle = card.querySelector(".line-title");
  const pieTitle = card.querySelector(".pie-title");
  const tableNote = card.querySelector(".table-note");

  if (lineTitle) {
    lineTitle.textContent = `${sheet.name}｜折线图变量：${getMetricLabel(sheet.lineMetric)}`;
  }

  if (pieTitle) {
    pieTitle.textContent = `${sheet.name}｜饼图变量：${getMetricLabel(sheet.pieMetric)}`;
  }

  if (tableNote) {
    tableNote.textContent = `默认展示前 5 行，其余折叠。当前排序变量：${getMetricLabel(sheet.lineMetric)}`;
  }

  renderLineChart(sheet, sortedRows, card.querySelector(".line-canvas"));
  renderPieChart(sheet, sortedRows, card.querySelector(".pie-canvas"));
  renderSheetTable(card, sortedRows);
}

function renderLineChart(sheet, rows, canvas) {
  const chartKey = `${sheet.id}:line`;
  destroyChart(chartKey);

  if (!canvas || typeof window.Chart !== "function") {
    return;
  }

  const chart = new Chart(canvas, {
    type: "line",
    data: {
      labels: rows.map((row) => row.adGroup),
      datasets: [
        {
          label: getMetricLabel(sheet.lineMetric),
          data: rows.map((row) => row[sheet.lineMetric]),
          borderColor: "#5ce9ff",
          backgroundColor: "rgba(92, 233, 255, 0.24)",
          borderWidth: 2,
          pointRadius: 2,
          pointHoverRadius: 3,
          tension: 0.25,
          fill: true
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          labels: {
            color: "#b3eeff"
          }
        }
      },
      scales: {
        x: {
          ticks: {
            color: "#7bc8e7",
            maxRotation: 45,
            minRotation: 25
          },
          grid: {
            color: "rgba(103, 193, 249, 0.18)"
          }
        },
        y: {
          beginAtZero: true,
          ticks: {
            color: "#7bc8e7"
          },
          grid: {
            color: "rgba(103, 193, 249, 0.18)"
          }
        }
      }
    }
  });

  dashboardState.charts.set(chartKey, chart);
}

function renderPieChart(sheet, rows, canvas) {
  const chartKey = `${sheet.id}:pie`;
  destroyChart(chartKey);

  if (!canvas || typeof window.Chart !== "function") {
    return;
  }

  const ranked = rows
    .map((row) => ({ label: row.adGroup, value: toNumber(row[sheet.pieMetric]) }))
    .sort((a, b) => b.value - a.value);

  const top = ranked.slice(0, 8);
  const rest = ranked.slice(8);
  const restValue = rest.reduce((sum, item) => sum + item.value, 0);

  const labels = top.map((item) => item.label);
  const values = top.map((item) => item.value);

  if (restValue > 0) {
    labels.push("其他");
    values.push(restValue);
  }

  if (!values.length) {
    labels.push("无数据");
    values.push(0);
  }

  const chart = new Chart(canvas, {
    type: "doughnut",
    data: {
      labels,
      datasets: [
        {
          data: values,
          borderWidth: 1,
          borderColor: "rgba(9, 18, 65, 0.85)",
          backgroundColor: [
            "#5ce9ff",
            "#8d76ff",
            "#4fb7ff",
            "#6de2b4",
            "#ff87d0",
            "#5d87ff",
            "#59cff2",
            "#ab9cff",
            "#4b638d"
          ]
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            color: "#b3eeff",
            boxWidth: 10,
            boxHeight: 10
          }
        }
      }
    }
  });

  dashboardState.charts.set(chartKey, chart);
}

function renderSheetTable(card, rows) {
  const topBody = card.querySelector(".top-body");
  const restBody = card.querySelector(".rest-body");
  const restBlock = card.querySelector(".rest-block");
  const restSummary = card.querySelector(".rest-summary");

  if (!topBody || !restBody || !restBlock || !restSummary) {
    return;
  }

  const topRows = rows.slice(0, 5);
  const restRows = rows.slice(5);

  topBody.innerHTML = topRows.map((row) => tableRowMarkup(row)).join("");

  if (!restRows.length) {
    restBody.innerHTML = "";
    restBlock.classList.add("hidden");
    restBlock.open = false;
    return;
  }

  restBody.innerHTML = restRows.map((row) => tableRowMarkup(row)).join("");
  restSummary.textContent = `展开剩余 ${restRows.length} 行（下拉滚动）`;
  restBlock.classList.remove("hidden");
  restBlock.open = false;
}

function tableRowMarkup(row) {
  return `
    <tr>
      <td>${escapeHtml(row.adGroup)}</td>
      <td>${escapeHtml(row.channel)}</td>
      <td>${escapeHtml(row.publisher)}</td>
      <td>${formatMetric("clicks", row.clicks)}</td>
      <td>${formatMetric("dpv", row.dpv)}</td>
      <td>${formatMetric("atc", row.atc)}</td>
      <td>${formatMetric("revenue", row.revenue)}</td>
      <td>${formatMetric("units", row.units)}</td>
      <td>${formatMetric("bonus", row.bonus)}</td>
    </tr>
  `;
}

function getSheetCard(sheetId) {
  return nodes.sheetBoard?.querySelector(`[data-sheet-id="${cssEscape(sheetId)}"]`) || null;
}

function sortRowsByMetric(rows, metricKey) {
  return rows
    .slice()
    .sort((a, b) => toNumber(b[metricKey]) - toNumber(a[metricKey]) || a.adGroup.localeCompare(b.adGroup, "zh-CN"));
}

function calculateTotals(rows) {
  return rows.reduce(
    (sum, row) => {
      METRIC_CONFIG.forEach((item) => {
        sum[item.key] += toNumber(row[item.key]);
      });
      return sum;
    },
    {
      clicks: 0,
      dpv: 0,
      atc: 0,
      revenue: 0,
      units: 0,
      bonus: 0
    }
  );
}

function getMetricLabel(metricKey) {
  const matched = METRIC_CONFIG.find((item) => item.key === metricKey);
  return matched ? matched.label : metricKey;
}

function metricFractionDigits(metricKey) {
  return metricKey === "revenue" || metricKey === "bonus" ? 2 : 0;
}

function formatMetric(metricKey, value) {
  return formatNumber(value, metricFractionDigits(metricKey));
}

function destroyAllCharts() {
  dashboardState.charts.forEach((chart) => {
    try {
      chart.destroy();
    } catch (_) {
      // Ignore chart destroy errors.
    }
  });
  dashboardState.charts.clear();
}

function destroyChart(chartKey) {
  const chart = dashboardState.charts.get(chartKey);
  if (!chart) {
    return;
  }
  try {
    chart.destroy();
  } catch (_) {
    // Ignore chart destroy errors.
  }
  dashboardState.charts.delete(chartKey);
}

function showLoading(message) {
  if (nodes.loadingPanel) {
    nodes.loadingPanel.textContent = message;
    nodes.loadingPanel.classList.remove("hidden");
  }
  nodes.sheetBoard?.classList.add("hidden");
  setStatus(message);
}

function showBoard() {
  nodes.loadingPanel?.classList.add("hidden");
  nodes.sheetBoard?.classList.remove("hidden");
}

function showError(message) {
  if (nodes.loadingPanel) {
    nodes.loadingPanel.textContent = message;
    nodes.loadingPanel.classList.remove("hidden");
  }
  nodes.sheetBoard?.classList.add("hidden");
  setStatus(message, true);
}

function setStatus(message, isError) {
  if (!nodes.sourceStatus) {
    return;
  }
  nodes.sourceStatus.textContent = message;
  nodes.sourceStatus.style.color = isError ? "#ff9dc3" : "";
}

function parseSheetRows(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  if (!matrix.length) {
    return { headers: [], rows: [] };
  }

  const headerIndex = detectHeaderRowIndex(matrix);
  const headerRow = Array.isArray(matrix[headerIndex]) ? matrix[headerIndex] : [];
  const headers = sanitizeHeaders(headerRow);

  const rows = matrix
    .slice(headerIndex + 1)
    .filter((row) => row.some((cell) => String(cell ?? "").trim() !== ""))
    .map((row) => {
      const item = {};
      headers.forEach((header, index) => {
        item[header] = row[index] ?? "";
      });
      return item;
    });

  return { headers, rows };
}

function detectHeaderRowIndex(matrix) {
  const maxScan = Math.min(matrix.length, 30);
  let bestIndex = 0;
  let bestScore = -1;

  for (let i = 0; i < maxScan; i += 1) {
    const row = Array.isArray(matrix[i]) ? matrix[i] : [];
    const values = row.map((cell) => String(cell ?? "").trim()).filter(Boolean);
    if (values.length < 2) {
      continue;
    }

    const textCount = values.filter((value) => Number.isNaN(Number(value))).length;
    const score = values.length * 2 + textCount;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  return bestIndex;
}

function sanitizeHeaders(row) {
  const used = new Set();
  return row.map((cell, index) => {
    const base = String(cell ?? "").trim() || `字段${index + 1}`;
    let name = base;
    let suffix = 2;
    while (used.has(name)) {
      name = `${base}_${suffix}`;
      suffix += 1;
    }
    used.add(name);
    return name;
  });
}

async function parseManifestResponse(response) {
  const text = await response.text();
  return parseManifestText(text);
}

function parseManifestText(text) {
  const source = String(text || "").trim();
  if (!source) {
    return [];
  }

  try {
    return normalizeManifestFiles(JSON.parse(source));
  } catch (_) {
    const chunks = extractJsonObjects(source);
    const merged = [];
    chunks.forEach((chunk) => {
      try {
        merged.push(...normalizeManifestFiles(JSON.parse(chunk)));
      } catch (_) {
        // Ignore malformed chunks.
      }
    });
    return dedupeManifestFiles(merged);
  }
}

function extractJsonObjects(text) {
  const parts = [];
  let depth = 0;
  let start = -1;
  let inString = false;
  let escaped = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];

    if (escaped) {
      escaped = false;
      continue;
    }

    if (char === "\\") {
      escaped = true;
      continue;
    }

    if (char === '"') {
      inString = !inString;
      continue;
    }

    if (inString) {
      continue;
    }

    if (char === "{") {
      if (depth === 0) {
        start = i;
      }
      depth += 1;
      continue;
    }

    if (char === "}") {
      depth -= 1;
      if (depth === 0 && start >= 0) {
        parts.push(text.slice(start, i + 1));
        start = -1;
      }
    }
  }

  return parts;
}

function normalizeManifestFiles(manifest) {
  const source = Array.isArray(manifest) ? manifest : Array.isArray(manifest?.files) ? manifest.files : [];
  const files = source
    .map((item) => {
      if (typeof item === "string") {
        return {
          name: getFileBaseName(item),
          path: normalizeFilePath(item)
        };
      }

      return {
        name: item?.name || getFileBaseName(item?.path || ""),
        path: normalizeFilePath(item?.path || "")
      };
    })
    .filter((item) => item.path);

  return dedupeManifestFiles(files);
}

function dedupeManifestFiles(files) {
  const map = new Map();
  files.forEach((item) => {
    if (!item.path) {
      return;
    }
    map.set(item.path, item);
  });
  return Array.from(map.values());
}

function normalizeFilePath(pathValue) {
  const path = String(pathValue || "").trim();
  if (!path) {
    return "";
  }
  if (/^https?:\/\//i.test(path) || path.startsWith("./") || path.startsWith("../") || path.startsWith("/")) {
    return path;
  }
  return `./${path}`;
}

function getFileBaseName(pathValue) {
  const normalized = String(pathValue || "").replace(/\\/g, "/");
  const rawName = normalized.split("/").pop() || pathValue;
  try {
    return decodeURIComponent(rawName);
  } catch (_) {
    return rawName;
  }
}

function normalizeHeader(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()（）_\-]/g, "");
}

function toNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  const normalized = String(value ?? "")
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "");
  const parsed = parseFloat(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatNumber(value, fractionDigits) {
  return new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits
  }).format(Number.isFinite(value) ? value : 0);
}

function setText(node, value) {
  if (!node) {
    return;
  }
  node.textContent = value;
}

function bindIfPresent(node, eventName, handler) {
  if (!node) {
    return;
  }
  node.addEventListener(eventName, handler);
}

function cssEscape(value) {
  if (window.CSS && typeof window.CSS.escape === "function") {
    return window.CSS.escape(String(value));
  }
  return String(value).replace(/[^a-zA-Z0-9_-]/g, "_");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
