const overviewState = {
  amazon: {
    fileTitle: "",
    rawRows: [],
    headers: [],
    mapping: null,
    rows: []
  },
  influencer: {
    fileTitle: "",
    rawRows: [],
    headers: [],
    mapping: null,
    metricLabel: "红人指标",
    rows: []
  },
  mergedRows: [],
  charts: {
    trend: null,
    scatter: null
  }
};

const overviewNodes = {
  amazonFileInput: document.getElementById("amazonFileInput"),
  amazonFileName: document.getElementById("amazonFileName"),
  influencerFileInput: document.getElementById("influencerFileInput"),
  influencerFileName: document.getElementById("influencerFileName"),

  amazonMappingPanel: document.getElementById("amazonMappingPanel"),
  influencerMappingPanel: document.getElementById("influencerMappingPanel"),

  amazonDateColumn: document.getElementById("amazonDateColumn"),
  amazonAsinColumn: document.getElementById("amazonAsinColumn"),
  amazonSalesColumn: document.getElementById("amazonSalesColumn"),
  amazonAsinFallback: document.getElementById("amazonAsinFallback"),

  influencerDateColumn: document.getElementById("influencerDateColumn"),
  influencerNameColumn: document.getElementById("influencerNameColumn"),
  influencerDpvColumn: document.getElementById("influencerDpvColumn"),

  applyAmazonMappingBtn: document.getElementById("applyAmazonMappingBtn"),
  applyInfluencerMappingBtn: document.getElementById("applyInfluencerMappingBtn"),

  filterPanel: document.getElementById("filterPanel"),
  asinFilter: document.getElementById("asinFilter"),
  influencerFilter: document.getElementById("influencerFilter"),
  rangeSelect: document.getElementById("rangeSelect"),

  kpiPanel: document.getElementById("kpiPanel"),
  formulaPanel: document.getElementById("formulaPanel"),
  chartPanel: document.getElementById("chartPanel"),
  tablePanel: document.getElementById("tablePanel"),
  emptyState: document.getElementById("emptyState"),

  trendChartTitle: document.getElementById("trendChartTitle"),
  trendChartNote: document.getElementById("trendChartNote"),
  scatterChartTitle: document.getElementById("scatterChartTitle"),
  scatterChartNote: document.getElementById("scatterChartNote"),
  metricHeaderLabel: document.getElementById("metricHeaderLabel"),
  metricLogHeaderLabel: document.getElementById("metricLogHeaderLabel"),

  tableSummary: document.getElementById("correlationTableSummary"),
  tableBody: document.getElementById("correlationTableBody")
};

const AMAZON_KEYWORDS = {
  date: ["日期", "销售时间", "销售日期", "date", "day", "dt", "time"],
  asin: ["asin", "子asin", "parent asin", "sku", "msku", "父asin"],
  sales: ["销量", "销售量", "销售件数", "units", "orders", "qty"]
};

const INFLUENCER_KEYWORDS = {
  date: ["日期", "date", "day", "dt", "时间"],
  name: ["红人", "publisher", "达人", "kol", "账号", "influencer", "出版商"],
  dpv: ["销量", "总销量", "销售量", "dpv", "总dpv", "总 dpv", "访问", "流量", "purchase", "购买次数"]
};

overviewNodes.amazonFileInput.addEventListener("change", (event) => handleDataFileUpload("amazon", event));
overviewNodes.influencerFileInput.addEventListener("change", (event) => handleDataFileUpload("influencer", event));
overviewNodes.applyAmazonMappingBtn.addEventListener("click", applyAmazonMapping);
overviewNodes.applyInfluencerMappingBtn.addEventListener("click", applyInfluencerMapping);
overviewNodes.asinFilter.addEventListener("change", renderCorrelationDashboard);
overviewNodes.influencerFilter.addEventListener("change", renderCorrelationDashboard);
overviewNodes.rangeSelect.addEventListener("change", renderCorrelationDashboard);

function handleDataFileUpload(type, event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  const fileNameNode = type === "amazon" ? overviewNodes.amazonFileName : overviewNodes.influencerFileName;
  fileNameNode.textContent = `已上传: ${file.name}`;

  const reader = new FileReader();
  reader.onload = function onload(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const parsed = parseSheetRows(sheet);
      const rows = parsed.rows;

      if (!rows.length) {
        throw new Error("文件没有数据");
      }

      const stateSection = overviewState[type];
      stateSection.fileTitle = stripFileExtension(file.name);
      stateSection.rawRows = rows;
      stateSection.headers = parsed.headers;

      if (type === "amazon") {
        initAmazonSelectors(stateSection.headers);
        overviewNodes.amazonMappingPanel.classList.remove("hidden");
      } else {
        initInfluencerSelectors(stateSection.headers);
        overviewNodes.influencerMappingPanel.classList.remove("hidden");
      }

      showSetupState();
    } catch (error) {
      window.alert(`读取失败：${error.message}`);
    }
  };

  reader.readAsArrayBuffer(file);
}

function initAmazonSelectors(headers) {
  buildSelectOptions(overviewNodes.amazonDateColumn, headers, true);
  buildSelectOptions(overviewNodes.amazonAsinColumn, headers, false);
  buildSelectOptions(overviewNodes.amazonSalesColumn, headers, true);

  overviewNodes.amazonDateColumn.value = guessHeader(headers, AMAZON_KEYWORDS.date);
  overviewNodes.amazonAsinColumn.value = guessHeader(headers, AMAZON_KEYWORDS.asin);
  overviewNodes.amazonSalesColumn.value = guessHeader(headers, AMAZON_KEYWORDS.sales);
  overviewNodes.amazonAsinFallback.value = extractAsinFromText(overviewState.amazon.fileTitle) || "";
}

function initInfluencerSelectors(headers) {
  buildSelectOptions(overviewNodes.influencerDateColumn, headers, true);
  buildSelectOptions(overviewNodes.influencerNameColumn, headers, false);
  buildSelectOptions(overviewNodes.influencerDpvColumn, headers, true);

  overviewNodes.influencerDateColumn.value = guessHeader(headers, INFLUENCER_KEYWORDS.date);
  overviewNodes.influencerNameColumn.value = guessHeader(headers, INFLUENCER_KEYWORDS.name);
  overviewNodes.influencerDpvColumn.value = guessHeader(headers, INFLUENCER_KEYWORDS.dpv);
}

function buildSelectOptions(selectNode, headers, required) {
  selectNode.innerHTML = "";

  const emptyOption = document.createElement("option");
  emptyOption.value = "";
  emptyOption.textContent = required ? "请选择列" : "不使用 / 未选择";
  selectNode.appendChild(emptyOption);

  headers.forEach((header) => {
    const option = document.createElement("option");
    option.value = header;
    option.textContent = header;
    selectNode.appendChild(option);
  });
}

function guessHeader(headers, keywords) {
  for (const keyword of keywords) {
    const matched = headers.find((header) => {
      const normalized = normalizeHeader(header);
      return normalized.includes(normalizeHeader(keyword));
    });
    if (matched) {
      return matched;
    }
  }
  return "";
}

function applyAmazonMapping() {
  const mapping = {
    date: overviewNodes.amazonDateColumn.value,
    asin: overviewNodes.amazonAsinColumn.value,
    sales: overviewNodes.amazonSalesColumn.value,
    asinFallback: String(overviewNodes.amazonAsinFallback.value || "").trim() || extractAsinFromText(overviewState.amazon.fileTitle) || "单ASIN"
  };

  if (!mapping.date || !mapping.sales) {
    window.alert("请完成 Amazon 的日期列、销量列映射。ASIN列可留空。");
    return;
  }

  const rows = normalizeAmazonRows(overviewState.amazon.rawRows, mapping);
  if (!rows.length) {
    window.alert("Amazon 数据未解析到有效记录，请检查映射。");
    return;
  }

  overviewState.amazon.mapping = mapping;
  overviewState.amazon.rows = rows;
  tryRenderCorrelation();
}

function applyInfluencerMapping() {
  const mapping = {
    date: overviewNodes.influencerDateColumn.value,
    name: overviewNodes.influencerNameColumn.value,
    dpv: overviewNodes.influencerDpvColumn.value
  };

  if (!mapping.date || !mapping.dpv) {
    window.alert("请完成红人的日期列与指标列映射。");
    return;
  }

  const rows = normalizeInfluencerRows(overviewState.influencer.rawRows, mapping);
  if (!rows.length) {
    window.alert("红人数据未解析到有效记录，请检查映射。");
    return;
  }

  overviewState.influencer.mapping = mapping;
  overviewState.influencer.metricLabel = mapping.dpv || "红人指标";
  overviewState.influencer.rows = rows;
  tryRenderCorrelation();
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
    const key = `${asin}||${date}`;
    const current = grouped.get(key) || { date, asin, sales: 0 };
    current.sales += sales;
    grouped.set(key, current);
  });

  return Array.from(grouped.values()).sort((a, b) => a.date.localeCompare(b.date) || a.asin.localeCompare(b.asin, "zh-CN"));
}

function normalizeInfluencerRows(rows, mapping) {
  const grouped = new Map();

  rows.forEach((row) => {
    const date = normalizeDateLabel(row[mapping.date]);
    if (!date) {
      return;
    }

    const influencer = mapping.name ? String(row[mapping.name] || "").trim() || "未命名红人" : "默认红人";
    const dpv = toNumber(row[mapping.dpv]);

    const key = `${influencer}||${date}`;
    const current = grouped.get(key) || { date, influencer, dpv: 0 };
    current.dpv += dpv;
    grouped.set(key, current);
  });

  return Array.from(grouped.values()).sort((a, b) => a.date.localeCompare(b.date) || a.influencer.localeCompare(b.influencer, "zh-CN"));
}

function tryRenderCorrelation() {
  const ready = Boolean(overviewState.amazon.mapping && overviewState.influencer.mapping);
  if (!ready) {
    showSetupState();
    return;
  }

  initFilterOptions();
  showDashboardState();
  renderCorrelationDashboard();
}

function initFilterOptions() {
  const asinPrev = overviewNodes.asinFilter.value;
  const influencerPrev = overviewNodes.influencerFilter.value;

  const asinList = Array.from(new Set(overviewState.amazon.rows.map((row) => row.asin))).sort((a, b) => a.localeCompare(b, "zh-CN"));
  const influencerList = Array.from(new Set(overviewState.influencer.rows.map((row) => row.influencer))).sort((a, b) =>
    a.localeCompare(b, "zh-CN")
  );

  overviewNodes.asinFilter.innerHTML = "";
  overviewNodes.influencerFilter.innerHTML = "";

  appendFilterOption(overviewNodes.asinFilter, "__ALL__", "全部ASIN（汇总）");
  appendFilterOption(overviewNodes.influencerFilter, "__ALL__", "全部红人（汇总）");

  asinList.forEach((asin) => appendFilterOption(overviewNodes.asinFilter, asin, asin));
  influencerList.forEach((name) => appendFilterOption(overviewNodes.influencerFilter, name, name));

  overviewNodes.asinFilter.value = asinList.includes(asinPrev) ? asinPrev : asinList[0] || "__ALL__";
  overviewNodes.influencerFilter.value = influencerList.includes(influencerPrev) ? influencerPrev : influencerList[0] || "__ALL__";
}

function appendFilterOption(selectNode, value, text) {
  const option = document.createElement("option");
  option.value = value;
  option.textContent = text;
  selectNode.appendChild(option);
}

function renderCorrelationDashboard() {
  if (!overviewState.amazon.mapping || !overviewState.influencer.mapping) {
    return;
  }

  const mergedRows = buildMergedRows();
  overviewState.mergedRows = mergedRows;

  if (!mergedRows.length) {
    overviewNodes.tableBody.innerHTML = "";
    overviewNodes.tableSummary.textContent = "当前筛选范围没有可对齐数据。";
    setKpiFallback();
    return;
  }

  const asinLabel = getAsinLabel();
  const influencerLabel = getInfluencerLabel();
  const metricLabel = overviewState.influencer.metricLabel || "红人指标";

  const salesList = mergedRows.map((row) => row.sales);
  const dpvList = mergedRows.map((row) => row.dpv);
  const logSalesList = mergedRows.map((row) => row.logSales);
  const logDpvList = mergedRows.map((row) => row.logDpv);

  const corrRaw = computePearson(dpvList, salesList);
  const corrLog = computePearson(logDpvList, logSalesList);

  renderKpiLine({ mergedRows, corrRaw, corrLog, asinLabel, influencerLabel });
  renderTrendChart(mergedRows, asinLabel, influencerLabel, metricLabel);
  renderScatterChart(mergedRows, asinLabel, influencerLabel, metricLabel);
  renderTable(mergedRows, asinLabel, influencerLabel, corrRaw, corrLog, metricLabel);
}

function buildMergedRows() {
  const asin = overviewNodes.asinFilter.value;
  const influencer = overviewNodes.influencerFilter.value;

  const amazonSeries = buildAmazonSeries(asin);
  const influencerSeries = buildInfluencerSeries(influencer);

  const allDates = Array.from(new Set([...amazonSeries.keys(), ...influencerSeries.keys()])).sort((a, b) =>
    a.localeCompare(b)
  );

  let rows = allDates.map((date) => {
    const sales = amazonSeries.get(date) || 0;
    const dpv = influencerSeries.get(date) || 0;
    return {
      date,
      sales,
      dpv,
      logSales: Math.log1p(Math.max(0, sales)),
      logDpv: Math.log1p(Math.max(0, dpv))
    };
  });

  const range = overviewNodes.rangeSelect.value;
  if (range !== "all") {
    const count = parseInt(range, 10);
    if (Number.isFinite(count)) {
      rows = rows.slice(-count);
    }
  }

  return rows;
}

function buildAmazonSeries(selectedAsin) {
  const map = new Map();

  overviewState.amazon.rows.forEach((row) => {
    if (selectedAsin !== "__ALL__" && row.asin !== selectedAsin) {
      return;
    }
    map.set(row.date, (map.get(row.date) || 0) + row.sales);
  });

  return map;
}

function buildInfluencerSeries(selectedInfluencer) {
  const map = new Map();

  overviewState.influencer.rows.forEach((row) => {
    if (selectedInfluencer !== "__ALL__" && row.influencer !== selectedInfluencer) {
      return;
    }
    map.set(row.date, (map.get(row.date) || 0) + row.dpv);
  });

  return map;
}

function getAsinLabel() {
  return overviewNodes.asinFilter.value === "__ALL__" ? "全部ASIN" : overviewNodes.asinFilter.value;
}

function getInfluencerLabel() {
  return overviewNodes.influencerFilter.value === "__ALL__" ? "全部红人" : overviewNodes.influencerFilter.value;
}

function renderKpiLine(payload) {
  const { mergedRows, corrRaw, corrLog, asinLabel, influencerLabel } = payload;
  const salesSum = mergedRows.reduce((sum, row) => sum + row.sales, 0);
  const dpvSum = mergedRows.reduce((sum, row) => sum + row.dpv, 0);

  setText("kpiAsin", asinLabel);
  setText("kpiInfluencer", influencerLabel);
  setText("kpiCorrRaw", formatCorrelation(corrRaw));
  setText("kpiCorrLog", formatCorrelation(corrLog));
  setText("kpiDays", `${mergedRows.length} 天`);
  setText("kpiSalesSum", formatNumber(salesSum, 0));
  setText("kpiDpvSum", formatNumber(dpvSum, 0));
}

function setKpiFallback() {
  setText("kpiAsin", getAsinLabel());
  setText("kpiInfluencer", getInfluencerLabel());
  setText("kpiCorrRaw", "-");
  setText("kpiCorrLog", "-");
  setText("kpiDays", "0 天");
  setText("kpiSalesSum", "0");
  setText("kpiDpvSum", "0");
}

function renderTrendChart(rows, asinLabel, influencerLabel, metricLabel) {
  if (overviewState.charts.trend) {
    overviewState.charts.trend.destroy();
  }

  overviewNodes.trendChartTitle.textContent = `${asinLabel} 销量 vs ${influencerLabel} ${metricLabel}`;
  overviewNodes.trendChartNote.textContent = `同日期对齐，共 ${rows.length} 天`;

  overviewState.charts.trend = new Chart(document.getElementById("correlationTrendChart"), {
    type: "line",
    data: {
      labels: rows.map((row) => row.date),
      datasets: [
        {
          label: "ASIN销量",
          data: rows.map((row) => row.sales),
          borderColor: "#f0c66a",
          backgroundColor: "rgba(240, 198, 106, 0.22)",
          yAxisID: "ySales",
          fill: true,
          tension: 0.3,
          pointRadius: 2
        },
        {
          label: metricLabel,
          data: rows.map((row) => row.dpv),
          borderColor: "#bf8d35",
          backgroundColor: "rgba(191, 141, 53, 0.16)",
          yAxisID: "yDpv",
          fill: false,
          tension: 0.3,
          pointRadius: 2
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
        yDpv: {
          type: "linear",
          position: "right",
          beginAtZero: true,
          ticks: { color: "#ba975f" },
          grid: { drawOnChartArea: false }
        }
      }
    }
  });
}

function renderScatterChart(rows, asinLabel, influencerLabel, metricLabel) {
  if (overviewState.charts.scatter) {
    overviewState.charts.scatter.destroy();
  }

  const points = rows.map((row) => ({ x: row.logDpv, y: row.logSales }));
  const regression = computeLinearRegression(points);

  const datasets = [
    {
      label: "样本点",
      data: points,
      backgroundColor: "rgba(240, 198, 106, 0.76)",
      borderColor: "rgba(240, 198, 106, 1)",
      pointRadius: 4,
      pointHoverRadius: 4,
      showLine: false
    }
  ];

  if (regression) {
    const minX = Math.min(...points.map((point) => point.x));
    const maxX = Math.max(...points.map((point) => point.x));
    datasets.push({
      label: "拟合线",
      type: "line",
      data: [
        { x: minX, y: regression.intercept + regression.slope * minX },
        { x: maxX, y: regression.intercept + regression.slope * maxX }
      ],
      borderColor: "#8e6628",
      borderWidth: 2,
      pointRadius: 0,
      fill: false,
      tension: 0
    });
  }

  overviewNodes.scatterChartTitle.textContent = `log(1+${metricLabel}) 与 log(1+销量) 散点`;
  overviewNodes.scatterChartNote.textContent = `${asinLabel} / ${influencerLabel}`;

  overviewState.charts.scatter = new Chart(document.getElementById("correlationScatterChart"), {
    type: "scatter",
    data: { datasets },
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
          grid: { color: "rgba(233, 191, 93, 0.12)" },
          title: {
            display: true,
            text: `log(1 + ${metricLabel})`,
            color: "#ba975f"
          }
        },
        y: {
          ticks: { color: "#ba975f" },
          grid: { color: "rgba(233, 191, 93, 0.12)" },
          title: {
            display: true,
            text: "log(1 + 销量)",
            color: "#ba975f"
          }
        }
      }
    }
  });
}

function computeLinearRegression(points) {
  if (points.length < 2) {
    return null;
  }

  const meanX = points.reduce((sum, point) => sum + point.x, 0) / points.length;
  const meanY = points.reduce((sum, point) => sum + point.y, 0) / points.length;

  let numerator = 0;
  let denominator = 0;

  points.forEach((point) => {
    numerator += (point.x - meanX) * (point.y - meanY);
    denominator += (point.x - meanX) ** 2;
  });

  if (denominator === 0) {
    return null;
  }

  const slope = numerator / denominator;
  const intercept = meanY - slope * meanX;
  return { slope, intercept };
}

function renderTable(rows, asinLabel, influencerLabel, corrRaw, corrLog, metricLabel) {
  overviewNodes.metricHeaderLabel.textContent = metricLabel;
  overviewNodes.metricLogHeaderLabel.textContent = `log(1+${metricLabel})`;
  overviewNodes.tableBody.innerHTML = rows
    .slice()
    .reverse()
    .map(
      (row) => `
      <tr>
        <td>${row.date}</td>
        <td>${formatNumber(row.sales, 0)}</td>
        <td>${formatNumber(row.dpv, 0)}</td>
        <td>${formatNumber(row.logSales, 4)}</td>
        <td>${formatNumber(row.logDpv, 4)}</td>
      </tr>
    `
    )
    .join("");

  overviewNodes.tableSummary.textContent = `${asinLabel} × ${influencerLabel}(${metricLabel})｜对齐 ${rows.length} 天｜相关系数 ${formatCorrelation(
    corrRaw
  )}｜log相关系数 ${formatCorrelation(corrLog)}`;
}

function computePearson(xList, yList) {
  if (!xList.length || xList.length !== yList.length || xList.length < 2) {
    return NaN;
  }

  const n = xList.length;
  const meanX = xList.reduce((sum, value) => sum + value, 0) / n;
  const meanY = yList.reduce((sum, value) => sum + value, 0) / n;

  let numerator = 0;
  let sumX = 0;
  let sumY = 0;

  for (let i = 0; i < n; i += 1) {
    const dx = xList[i] - meanX;
    const dy = yList[i] - meanY;
    numerator += dx * dy;
    sumX += dx * dx;
    sumY += dy * dy;
  }

  const denominator = Math.sqrt(sumX * sumY);
  if (denominator === 0) {
    return NaN;
  }

  return numerator / denominator;
}

function showSetupState() {
  overviewNodes.filterPanel.classList.add("hidden");
  overviewNodes.kpiPanel.classList.add("hidden");
  overviewNodes.formulaPanel.classList.add("hidden");
  overviewNodes.chartPanel.classList.add("hidden");
  overviewNodes.tablePanel.classList.add("hidden");
  overviewNodes.emptyState.classList.remove("hidden");
}

function showDashboardState() {
  overviewNodes.filterPanel.classList.remove("hidden");
  overviewNodes.kpiPanel.classList.remove("hidden");
  overviewNodes.formulaPanel.classList.remove("hidden");
  overviewNodes.chartPanel.classList.remove("hidden");
  overviewNodes.tablePanel.classList.remove("hidden");
  overviewNodes.emptyState.classList.add("hidden");
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

function formatNumber(value, fractionDigits) {
  return new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits
  }).format(Number.isFinite(value) ? value : 0);
}

function formatCorrelation(value) {
  if (!Number.isFinite(value)) {
    return "-";
  }
  return formatNumber(value, 4);
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

function extractAsinFromText(text) {
  const source = String(text || "").toUpperCase();
  const match = source.match(/\bB0[A-Z0-9]{8}\b/);
  return match ? match[0] : "";
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
