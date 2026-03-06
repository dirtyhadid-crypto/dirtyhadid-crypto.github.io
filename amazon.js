const amazonState = {
  fileTitle: "",
  rawRows: [],
  headers: [],
  mapping: null,
  mappedRows: [],
  sort: {
    key: "date",
    direction: "desc"
  },
  charts: {
    trend: null,
    change: null
  }
};

const amazonNodes = {
  fileInput: document.getElementById("fileInput"),
  fileName: document.getElementById("fileName"),
  presetPanel: document.getElementById("presetPanel"),
  presetFileSelect: document.getElementById("presetFileSelect"),
  loadPresetBtn: document.getElementById("loadPresetBtn"),
  uploadStatus: document.getElementById("uploadStatus"),
  mappingPanel: document.getElementById("mappingPanel"),
  filterPanel: document.getElementById("filterPanel"),
  kpiPanel: document.getElementById("kpiPanel"),
  chartPanel: document.getElementById("chartPanel"),
  tablePanel: document.getElementById("tablePanel"),
  emptyState: document.getElementById("emptyState"),
  applyMappingBtn: document.getElementById("applyMappingBtn"),
  dateColumn: document.getElementById("dateColumn"),
  asinColumn: document.getElementById("asinColumn"),
  salesColumn: document.getElementById("salesColumn"),
  cvrColumn: document.getElementById("cvrColumn"),
  adOrdersColumn: document.getElementById("adOrdersColumn"),
  adClicksColumn: document.getElementById("adClicksColumn"),
  dpvColumn: document.getElementById("dpvColumn"),
  asinFallback: document.getElementById("asinFallback"),
  asinFilter: document.getElementById("asinFilter"),
  rangeSelect: document.getElementById("rangeSelect"),
  tableBody: document.getElementById("amazonTableBody"),
  tableSummary: document.getElementById("amazonTableSummary"),
  trendChartTitle: document.getElementById("trendChartTitle"),
  trendChartNote: document.getElementById("trendChartNote"),
  changeChartTitle: document.getElementById("changeChartTitle"),
  changeChartNote: document.getElementById("changeChartNote")
};

const AMAZON_KEYWORDS = {
  date: ["日期", "销售时间", "销售日期", "date", "day", "dt", "time"],
  asin: ["asin", "子asin", "parent asin", "sku"],
  sales: ["销量", "销售件数", "销售量", "units", "orders", "order units", "qty"],
  cvr: ["广告cvr", "广告转化率", "ad cvr", "ad conversion", "转化率", "cvr", "conversion"],
  adOrders: ["广告订单", "广告订单量", "ad orders", "attributed orders", "ad order", "购买次数", "总购买次数", "order count"],
  adClicks: ["广告点击", "广告点击量", "ad clicks", "clicks", "click"],
  dpv: ["dpv", "总dpv", "总 dpv", "访问", "流量"]
};

bindIfPresent(amazonNodes.fileInput, "click", clearFileInputValue);
bindIfPresent(amazonNodes.fileInput, "change", handleAmazonFileUpload);
bindIfPresent(amazonNodes.loadPresetBtn, "click", loadPresetFromManifest);
bindIfPresent(amazonNodes.applyMappingBtn, "click", applyAmazonMapping);
bindIfPresent(amazonNodes.asinFilter, "change", renderAmazonDashboard);
bindIfPresent(amazonNodes.rangeSelect, "change", renderAmazonDashboard);

initTableSorting();
initPresetFiles();

async function initPresetFiles() {
  if (!amazonNodes.presetPanel || !amazonNodes.presetFileSelect) {
    return;
  }

  try {
    const response = await fetch("./excel/manifest.json", { cache: "no-store" });
    if (!response.ok) {
      return;
    }

    const files = await parseManifestResponse(response);
    if (!files.length) {
      return;
    }

    amazonNodes.presetFileSelect.innerHTML = "";
    files.forEach((item) => {
      const option = document.createElement("option");
      option.value = item.path;
      option.textContent = item.name;
      amazonNodes.presetFileSelect.appendChild(option);
    });

    amazonNodes.presetPanel.classList.remove("hidden");
  } catch (_) {
    // Ignore manifest loading failure in static pages.
  }
}

async function loadPresetFromManifest() {
  if (typeof window.XLSX === "undefined") {
    setUploadStatus("Excel 解析库加载失败，请刷新页面后重试。", true);
    return;
  }

  const filePath = amazonNodes.presetFileSelect?.value || "";
  if (!filePath) {
    return;
  }

  setUploadStatus("正在加载仓库文件...");

  try {
    const response = await fetch(filePath, { cache: "no-store" });
    if (!response.ok) {
      throw new Error("读取仓库文件失败");
    }

    const fileName = getFileBaseName(filePath);
    if (amazonNodes.fileName) {
      amazonNodes.fileName.textContent = `已加载仓库文件: ${fileName}`;
    }

    const workbook = await readWorkbookFromResponse(response, filePath);
    processWorkbook(workbook, fileName);
  } catch (error) {
    setUploadStatus(`加载失败：${error.message}`, true);
    window.alert(`加载失败：${error.message}`);
  }
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
        // Ignore broken chunks.
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

    if (char === "\"") {
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
          path: item
        };
      }
      return {
        name: item?.name || getFileBaseName(item?.path || ""),
        path: item?.path || ""
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

async function readWorkbookFromResponse(response, filePath) {
  const lowerPath = String(filePath || "").toLowerCase();
  if (lowerPath.endsWith(".csv")) {
    const text = await response.text();
    return XLSX.read(text, { type: "string" });
  }
  const arrayBuffer = await response.arrayBuffer();
  return XLSX.read(arrayBuffer, { type: "array" });
}

function getFileBaseName(filePath) {
  const normalized = String(filePath || "").replace(/\\/g, "/");
  return normalized.split("/").pop() || filePath;
}

function handleAmazonFileUpload(event) {
  if (typeof window.XLSX === "undefined") {
    setUploadStatus("Excel 解析库加载失败，请刷新页面后重试。", true);
    window.alert("Excel 解析库加载失败，请刷新页面后重试。");
    return;
  }

  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  if (amazonNodes.fileName) {
    amazonNodes.fileName.textContent = `已上传: ${file.name}`;
  }
  amazonState.fileTitle = stripFileExtension(file.name);
  setUploadStatus("正在读取文件...");

  const reader = new FileReader();
  reader.onload = function onload(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      processWorkbook(workbook, file.name);
    } catch (error) {
      setUploadStatus(`读取失败：${error.message}`, true);
      window.alert(`读取失败：${error.message}`);
    }
  };

  reader.readAsArrayBuffer(file);
}

function processWorkbook(workbook, fileName) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const parsed = parseSheetRows(sheet);
  const rows = parsed.rows;

  if (!rows.length) {
    throw new Error("文件没有数据");
  }

  amazonState.fileTitle = stripFileExtension(fileName);
  amazonState.rawRows = rows;
  amazonState.headers = parsed.headers;
  initMappingSelectors(amazonState.headers);
  if (amazonNodes.mappingPanel) {
    amazonNodes.mappingPanel.classList.remove("hidden");
  }
  showSetupOnly();

  const autoApplied = tryAutoApplyMapping();
  if (autoApplied) {
    setUploadStatus(`已自动生成看板：${rows.length} 行数据。`);
  } else {
    setUploadStatus(`已读取 ${rows.length} 行，请检查映射后点击“应用映射并生成看板”。`);
  }
}

function initMappingSelectors(headers) {
  const selectors = [
    amazonNodes.dateColumn,
    amazonNodes.asinColumn,
    amazonNodes.salesColumn,
    amazonNodes.cvrColumn,
    amazonNodes.adOrdersColumn,
    amazonNodes.adClicksColumn,
    amazonNodes.dpvColumn
  ].filter(Boolean);

  if (!selectors.length) {
    return;
  }

  selectors.forEach((selectNode) => {
    selectNode.innerHTML = "";

    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = "不使用 / 未选择";
    selectNode.appendChild(emptyOption);

    headers.forEach((header) => {
      const option = document.createElement("option");
      option.value = header;
      option.textContent = header;
      selectNode.appendChild(option);
    });
  });

  if (amazonNodes.dateColumn) {
    amazonNodes.dateColumn.value = guessHeader(headers, AMAZON_KEYWORDS.date);
  }
  if (amazonNodes.asinColumn) {
    amazonNodes.asinColumn.value = guessHeader(headers, AMAZON_KEYWORDS.asin);
  }
  if (amazonNodes.salesColumn) {
    amazonNodes.salesColumn.value = guessHeader(headers, AMAZON_KEYWORDS.sales);
  }
  if (amazonNodes.cvrColumn) {
    amazonNodes.cvrColumn.value = guessHeader(headers, AMAZON_KEYWORDS.cvr);
  }
  if (amazonNodes.adOrdersColumn) {
    amazonNodes.adOrdersColumn.value = guessHeader(headers, AMAZON_KEYWORDS.adOrders);
  }
  if (amazonNodes.adClicksColumn) {
    amazonNodes.adClicksColumn.value = guessHeader(headers, AMAZON_KEYWORDS.adClicks);
  }
  if (amazonNodes.dpvColumn) {
    amazonNodes.dpvColumn.value = guessHeader(headers, AMAZON_KEYWORDS.dpv);
  }
  if (amazonNodes.asinFallback) {
    amazonNodes.asinFallback.value = extractAsinFromText(amazonState.fileTitle) || "";
  }
}

function guessHeader(headers, keywords) {
  for (const keyword of keywords) {
    const matched = headers.find((header) => {
      const name = normalizeHeader(header);
      return name.includes(normalizeHeader(keyword));
    });
    if (matched) {
      return matched;
    }
  }
  return "";
}

function applyAmazonMapping() {
  applyAmazonMappingInternal(getMappingFromSelectors(), { showAlert: true, source: "manual" });
}

function tryAutoApplyMapping() {
  return applyAmazonMappingInternal(getMappingFromSelectors(), { showAlert: false, source: "auto" });
}

function getMappingFromSelectors() {
  return {
    date: amazonNodes.dateColumn?.value || "",
    asin: amazonNodes.asinColumn?.value || "",
    sales: amazonNodes.salesColumn?.value || "",
    cvr: amazonNodes.cvrColumn?.value || "",
    adOrders: amazonNodes.adOrdersColumn?.value || "",
    adClicks: amazonNodes.adClicksColumn?.value || "",
    dpv: amazonNodes.dpvColumn?.value || "",
    asinFallback: String(amazonNodes.asinFallback?.value || "").trim() || extractAsinFromText(amazonState.fileTitle) || "单ASIN"
  };
}

function applyAmazonMappingInternal(mapping, options) {
  const showAlert = options?.showAlert ?? true;
  const source = options?.source || "manual";

  if (!mapping.date || !mapping.sales) {
    if (showAlert) {
      window.alert("请至少选择日期列、销量列。ASIN列可留空。");
    }
    return false;
  }

  const normalizedRows = normalizeAmazonRows(amazonState.rawRows, mapping);
  if (!normalizedRows.length) {
    if (showAlert) {
      window.alert("没有解析到有效数据，请检查日期/ASIN/销量字段。");
    }
    return false;
  }

  amazonState.mapping = mapping;
  amazonState.mappedRows = normalizedRows;
  initAsinFilter(normalizedRows);

  showDashboard();
  renderAmazonDashboard();

  if (source === "manual") {
    setUploadStatus(`看板已更新：${normalizedRows.length} 条有效记录。`);
  }

  return true;
}

function normalizeAmazonRows(rows, mapping) {
  const grouped = new Map();

  rows.forEach((row) => {
    const date = normalizeDateLabel(row[mapping.date]);
    const asinValue = mapping.asin ? row[mapping.asin] : "";
    const asin = String(asinValue || "").trim() || mapping.asinFallback || "单ASIN";
    if (!date || !asin) {
      return;
    }

    const sales = toNumber(row[mapping.sales]);
    const cvr = mapping.cvr ? toPercentValue(row[mapping.cvr]) : NaN;
    const adOrders = mapping.adOrders ? toNumber(row[mapping.adOrders]) : 0;
    const adClicks = mapping.adClicks ? toNumber(row[mapping.adClicks]) : 0;
    const dpv = mapping.dpv ? toNumber(row[mapping.dpv]) : 0;

    const key = `${asin}||${date}`;
    const current = grouped.get(key) || {
      date,
      asin,
      sales: 0,
      adOrders: 0,
      adClicks: 0,
      dpv: 0,
      cvrSum: 0,
      cvrCount: 0
    };

    current.sales += sales;
    current.adOrders += adOrders;
    current.adClicks += adClicks;
    current.dpv += dpv;

    if (Number.isFinite(cvr)) {
      current.cvrSum += cvr;
      current.cvrCount += 1;
    }

    grouped.set(key, current);
  });

  return Array.from(grouped.values())
    .map((item) => ({
      date: item.date,
      asin: item.asin,
      sales: item.sales,
      cvr: resolveCvr(item),
      adOrders: item.adOrders,
      adClicks: item.adClicks,
      dpv: item.dpv,
      cvrSum: item.cvrSum,
      cvrCount: item.cvrCount
    }))
    .sort((a, b) => a.date.localeCompare(b.date) || a.asin.localeCompare(b.asin, "zh-CN"));
}

function resolveCvr(item) {
  if (item.adClicks > 0) {
    return (item.adOrders / item.adClicks) * 100;
  }
  if (item.cvrCount > 0) {
    return item.cvrSum / item.cvrCount;
  }
  if (item.dpv > 0) {
    return (item.sales / item.dpv) * 100;
  }
  return NaN;
}

function initAsinFilter(rows) {
  const allAsins = Array.from(new Set(rows.map((row) => row.asin))).sort((a, b) => a.localeCompare(b, "zh-CN"));

  const previous = amazonNodes.asinFilter.value;
  amazonNodes.asinFilter.innerHTML = "";

  const allOption = document.createElement("option");
  allOption.value = "__ALL__";
  allOption.textContent = "全部ASIN（汇总）";
  amazonNodes.asinFilter.appendChild(allOption);

  allAsins.forEach((asin) => {
    const option = document.createElement("option");
    option.value = asin;
    option.textContent = asin;
    amazonNodes.asinFilter.appendChild(option);
  });

  if (allAsins.includes(previous)) {
    amazonNodes.asinFilter.value = previous;
  } else if (allAsins.length === 1) {
    amazonNodes.asinFilter.value = allAsins[0];
  } else {
    amazonNodes.asinFilter.value = allAsins[0] || "__ALL__";
  }
}

function getSelectedAsinLabel() {
  return amazonNodes.asinFilter.value === "__ALL__" ? "全部ASIN" : amazonNodes.asinFilter.value;
}

function getFilteredRows() {
  const asinFilter = amazonNodes.asinFilter.value;
  const range = amazonNodes.rangeSelect.value;

  let rows = [];
  if (asinFilter === "__ALL__") {
    rows = aggregateAcrossAsins(amazonState.mappedRows);
  } else {
    rows = amazonState.mappedRows.filter((row) => row.asin === asinFilter);
  }

  rows = rows.slice().sort((a, b) => a.date.localeCompare(b.date));

  if (range === "all") {
    return rows;
  }

  const count = parseInt(range, 10);
  if (!Number.isFinite(count)) {
    return rows;
  }

  return rows.slice(-count);
}

function aggregateAcrossAsins(rows) {
  const grouped = new Map();

  rows.forEach((row) => {
    const current = grouped.get(row.date) || {
      date: row.date,
      asin: "全部ASIN",
      sales: 0,
      adOrders: 0,
      adClicks: 0,
      dpv: 0,
      cvrSum: 0,
      cvrCount: 0
    };

    current.sales += row.sales;
    current.adOrders += row.adOrders;
    current.adClicks += row.adClicks;
    current.dpv += row.dpv;

    if (Number.isFinite(row.cvr)) {
      current.cvrSum += row.cvr;
      current.cvrCount += 1;
    }

    grouped.set(row.date, current);
  });

  return Array.from(grouped.values()).map((item) => ({
    date: item.date,
    asin: item.asin,
    sales: item.sales,
    cvr: resolveCvr(item),
    adOrders: item.adOrders,
    adClicks: item.adClicks,
    dpv: item.dpv,
    cvrSum: item.cvrSum,
    cvrCount: item.cvrCount
  }));
}

function buildDerivedRows(rows) {
  return rows.map((row, index) => {
    const previous = index > 0 ? rows[index - 1] : null;

    let salesChange = null;
    let salesChangeLabel = "-";

    if (previous) {
      if (previous.sales === 0) {
        if (row.sales === 0) {
          salesChange = 0;
          salesChangeLabel = "0.0%";
        } else {
          salesChangeLabel = "新增";
        }
      } else {
        salesChange = ((row.sales - previous.sales) / Math.abs(previous.sales)) * 100;
        salesChangeLabel = formatSignedPercent(salesChange);
      }
    }

    const cvrChange = previous && Number.isFinite(row.cvr) && Number.isFinite(previous.cvr) ? row.cvr - previous.cvr : null;

    return {
      ...row,
      salesChange,
      salesChangeLabel,
      cvrChange,
      cvrChangeLabel: Number.isFinite(cvrChange) ? formatPointDelta(cvrChange) : "-"
    };
  });
}

function renderAmazonDashboard() {
  if (!amazonState.mappedRows.length) {
    return;
  }

  const filteredRows = getFilteredRows();
  if (!filteredRows.length) {
    amazonNodes.tableBody.innerHTML = "";
    amazonNodes.tableSummary.textContent = "当前筛选没有数据。";
    return;
  }

  const derivedRows = buildDerivedRows(filteredRows);
  const asinLabel = getSelectedAsinLabel();

  renderKpiLine(derivedRows, asinLabel);
  renderTrendChart(derivedRows, asinLabel);
  renderChangeChart(derivedRows, asinLabel);
  renderTable(derivedRows, asinLabel);
}

function renderKpiLine(rows, asinLabel) {
  const latest = rows[rows.length - 1];
  const totalSales = rows.reduce((sum, row) => sum + row.sales, 0);
  const avgSales = totalSales / Math.max(rows.length, 1);

  setText("kpiAsin", asinLabel || "-");
  setText("kpiDays", `${rows.length} 天`);
  setText("kpiLatestSales", formatNumber(latest.sales, 0));
  setText("kpiTotalSales", formatNumber(totalSales, 0));
  setText("kpiAvgSales", formatNumber(avgSales, 1));
  setText("kpiLatestCvr", formatPercent(latest.cvr, 2));
}

function renderTrendChart(rows, asinLabel) {
  const labels = rows.map((row) => row.date);

  if (amazonState.charts.trend) {
    amazonState.charts.trend.destroy();
  }

  amazonNodes.trendChartTitle.textContent = `${asinLabel} 日销量与广告转化率`;
  amazonNodes.trendChartNote.textContent = `${amazonState.fileTitle}｜记录 ${rows.length} 天`;

  if (typeof window.Chart !== "function") {
    amazonNodes.trendChartNote.textContent = "图表库加载失败，仅展示顶部汇总与明细表。";
    return;
  }

  amazonState.charts.trend = new Chart(document.getElementById("amazonTrendChart"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "日销量",
          data: rows.map((row) => row.sales),
          borderColor: "#f0c66a",
          backgroundColor: "rgba(240, 198, 106, 0.22)",
          yAxisID: "ySales",
          tension: 0.3,
          pointRadius: 2,
          fill: true
        },
        {
          label: "广告转化率(%)",
          data: rows.map((row) => (Number.isFinite(row.cvr) ? row.cvr : null)),
          borderColor: "#bf8d35",
          backgroundColor: "rgba(191, 141, 53, 0.18)",
          yAxisID: "yCvr",
          tension: 0.28,
          pointRadius: 2,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          labels: {
            boxWidth: 10,
            boxHeight: 10,
            color: "#e9c988"
          }
        }
      },
      scales: {
        x: {
          ticks: { color: "#ba975f" },
          grid: { color: "rgba(233, 191, 93, 0.12)" }
        },
        ySales: {
          type: "linear",
          position: "left",
          beginAtZero: true,
          ticks: { color: "#ba975f" },
          grid: { color: "rgba(233, 191, 93, 0.12)" }
        },
        yCvr: {
          type: "linear",
          position: "right",
          beginAtZero: true,
          ticks: {
            color: "#ba975f",
            callback(value) {
              return `${value}%`;
            }
          },
          grid: { drawOnChartArea: false }
        }
      }
    }
  });
}

function renderChangeChart(rows, asinLabel) {
  const labels = rows.map((row) => row.date);

  if (amazonState.charts.change) {
    amazonState.charts.change.destroy();
  }

  amazonNodes.changeChartTitle.textContent = `${asinLabel} 环比波动`;
  amazonNodes.changeChartNote.textContent = "销量环比(%) 与 转化率变化(百分点)";

  if (typeof window.Chart !== "function") {
    amazonNodes.changeChartNote.textContent = "图表库加载失败，仅展示顶部汇总与明细表。";
    return;
  }

  amazonState.charts.change = new Chart(document.getElementById("amazonChangeChart"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "销量环比(%)",
          data: rows.map((row) => (Number.isFinite(row.salesChange) ? row.salesChange : null)),
          borderColor: "#f0c66a",
          backgroundColor: "rgba(240, 198, 106, 0.2)",
          tension: 0.28,
          pointRadius: 2,
          fill: false
        },
        {
          label: "转化率变化(pp)",
          data: rows.map((row) => (Number.isFinite(row.cvrChange) ? row.cvrChange : null)),
          borderColor: "#8e6628",
          backgroundColor: "rgba(142, 102, 40, 0.2)",
          tension: 0.28,
          pointRadius: 2,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          labels: {
            boxWidth: 10,
            boxHeight: 10,
            color: "#e9c988"
          }
        }
      },
      scales: {
        x: {
          ticks: { color: "#ba975f" },
          grid: { color: "rgba(233, 191, 93, 0.12)" }
        },
        y: {
          ticks: {
            color: "#ba975f",
            callback(value) {
              return `${value}`;
            }
          },
          grid: { color: "rgba(233, 191, 93, 0.12)" }
        }
      }
    }
  });
}

function renderTable(rows, asinLabel) {
  const sortedRows = sortRows(rows, amazonState.sort.key, amazonState.sort.direction);
  amazonNodes.tableBody.innerHTML = sortedRows
    .map(
      (row) => `
      <tr>
        <td>${row.date}</td>
        <td>${escapeHtml(row.asin)}</td>
        <td>${formatNumber(row.sales, 0)}</td>
        <td>${formatPercent(row.cvr, 2)}</td>
        <td>${row.salesChangeLabel}</td>
        <td>${row.cvrChangeLabel}</td>
      </tr>
    `
    )
    .join("");

  const latest = rows[rows.length - 1];
  amazonNodes.tableSummary.textContent = `${amazonState.fileTitle}｜${asinLabel}｜共 ${rows.length} 天｜最近销量 ${formatNumber(
    latest.sales,
    0
  )}｜最近广告转化率 ${formatPercent(latest.cvr, 2)}`;
}

function initTableSorting() {
  const headers = document.querySelectorAll("#tablePanel th.sortable");
  headers.forEach((header) => {
    header.addEventListener("click", () => {
      const key = header.dataset.sortKey;
      if (!key) {
        return;
      }

      if (amazonState.sort.key === key) {
        amazonState.sort.direction = amazonState.sort.direction === "asc" ? "desc" : "asc";
      } else {
        amazonState.sort.key = key;
        amazonState.sort.direction = isNumericKey(key) ? "desc" : "asc";
      }

      updateSortHeaderState();
      renderAmazonDashboard();
    });
  });

  updateSortHeaderState();
}

function bindIfPresent(node, eventName, handler) {
  if (!node) {
    return;
  }
  node.addEventListener(eventName, handler);
}

function clearFileInputValue(event) {
  if (!event?.target) {
    return;
  }
  event.target.value = "";
}

function sortRows(rows, key, direction) {
  const cloned = rows.slice();

  cloned.sort((a, b) => {
    let result = 0;

    if (isNumericKey(key)) {
      result = (toSortableNumber(a[key]) || 0) - (toSortableNumber(b[key]) || 0);
    } else if (key === "date") {
      result = a.date.localeCompare(b.date);
    } else {
      result = String(a[key] || "").localeCompare(String(b[key] || ""), "zh-CN");
    }

    return direction === "asc" ? result : -result;
  });

  return cloned;
}

function toSortableNumber(value) {
  return Number.isFinite(value) ? value : -Infinity;
}

function isNumericKey(key) {
  return ["sales", "cvr", "salesChange", "cvrChange"].includes(key);
}

function updateSortHeaderState() {
  const headers = document.querySelectorAll("#tablePanel th.sortable");
  headers.forEach((header) => {
    const key = header.dataset.sortKey;
    const baseText = header.textContent.replace(/\s*[↑↓]$/, "");

    if (key === amazonState.sort.key) {
      const arrow = amazonState.sort.direction === "asc" ? "↑" : "↓";
      header.textContent = `${baseText} ${arrow}`;
      header.classList.add("active-sort");
    } else {
      header.textContent = baseText;
      header.classList.remove("active-sort");
    }
  });
}

function showSetupOnly() {
  amazonNodes.emptyState.classList.remove("hidden");
  amazonNodes.filterPanel.classList.add("hidden");
  amazonNodes.kpiPanel.classList.add("hidden");
  amazonNodes.chartPanel.classList.add("hidden");
  amazonNodes.tablePanel.classList.add("hidden");
}

function showDashboard() {
  amazonNodes.emptyState.classList.add("hidden");
  amazonNodes.filterPanel.classList.remove("hidden");
  amazonNodes.kpiPanel.classList.remove("hidden");
  amazonNodes.chartPanel.classList.remove("hidden");
  amazonNodes.tablePanel.classList.remove("hidden");
}

function normalizeHeader(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()（）_\-]/g, "");
}

function normalizeDateLabel(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  let dateObject = null;

  if (typeof value === "number") {
    const excelDate = XLSX.SSF.parse_date_code(value);
    if (excelDate) {
      dateObject = new Date(excelDate.y, excelDate.m - 1, excelDate.d);
    }
  } else if (value instanceof Date) {
    dateObject = new Date(value.getFullYear(), value.getMonth(), value.getDate());
  } else {
    const cleaned = String(value)
      .trim()
      .replace(/[年月]/g, "-")
      .replace(/日/g, "")
      .replace(/\./g, "-")
      .replace(/\//g, "-");

    if (/^\d{8}$/.test(cleaned)) {
      const y = Number(cleaned.slice(0, 4));
      const m = Number(cleaned.slice(4, 6));
      const d = Number(cleaned.slice(6, 8));
      dateObject = new Date(y, m - 1, d);
    } else {
      const parsed = new Date(cleaned);
      if (!Number.isNaN(parsed.getTime())) {
        dateObject = new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
      }
    }
  }

  if (!dateObject || Number.isNaN(dateObject.getTime())) {
    return "";
  }

  const y = dateObject.getFullYear();
  const m = String(dateObject.getMonth() + 1).padStart(2, "0");
  const d = String(dateObject.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
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

function toPercentValue(value) {
  if (value === null || value === undefined || value === "") {
    return NaN;
  }

  const text = String(value).trim();
  const parsed = toNumber(text);
  if (!Number.isFinite(parsed)) {
    return NaN;
  }

  if (text.includes("%")) {
    return parsed;
  }

  if (Math.abs(parsed) <= 1) {
    return parsed * 100;
  }

  return parsed;
}

function formatNumber(value, fractionDigits) {
  return new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits
  }).format(Number.isFinite(value) ? value : 0);
}

function formatPercent(value, fractionDigits) {
  if (!Number.isFinite(value)) {
    return "-";
  }
  return `${formatNumber(value, fractionDigits)}%`;
}

function formatSignedPercent(value) {
  if (!Number.isFinite(value)) {
    return "-";
  }
  const sign = value > 0 ? "+" : "";
  return `${sign}${formatNumber(value, 1)}%`;
}

function formatPointDelta(value) {
  if (!Number.isFinite(value)) {
    return "-";
  }
  const sign = value > 0 ? "+" : "";
  return `${sign}${formatNumber(value, 2)} pp`;
}

function stripFileExtension(fileName) {
  const name = String(fileName || "未命名文件");
  const index = name.lastIndexOf(".");
  if (index <= 0) {
    return name;
  }
  return name.slice(0, index);
}

function setText(id, value) {
  const node = document.getElementById(id);
  if (!node) {
    return;
  }
  node.textContent = value;
}

function setUploadStatus(message, isError) {
  if (!amazonNodes.uploadStatus) {
    return;
  }
  amazonNodes.uploadStatus.textContent = message;
  amazonNodes.uploadStatus.style.color = isError ? "#da8f66" : "";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
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
  const maxScan = Math.min(matrix.length, 40);
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

function extractAsinFromText(text) {
  const source = String(text || "").toUpperCase();
  const match = source.match(/\bB0[A-Z0-9]{8}\b/);
  return match ? match[0] : "";
}
