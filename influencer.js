const influencerState = {
  fileTitle: "",
  sheetName: "",
  rawRows: [],
  normalizedRows: [],
  columnMap: {},
  sort: {
    key: "clicks",
    direction: "desc"
  },
  charts: {
    metric: null,
    publisher: null
  }
};

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

const EXTRA_FIELD_CONFIG = [
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

const nodes = {
  fileInput: document.getElementById("fileInput"),
  fileName: document.getElementById("fileName"),
  presetPanel: document.getElementById("presetPanel"),
  presetFileSelect: document.getElementById("presetFileSelect"),
  loadPresetBtn: document.getElementById("loadPresetBtn"),
  kpiPanel: document.getElementById("kpiPanel"),
  chartPanel: document.getElementById("chartPanel"),
  tablePanel: document.getElementById("tablePanel"),
  emptyState: document.getElementById("emptyState"),
  tableBody: document.getElementById("influencerTableBody"),
  tableSummary: document.getElementById("tableSummary"),
  metricChartTitle: document.getElementById("metricChartTitle"),
  metricChartSubTitle: document.getElementById("metricChartSubTitle"),
  publisherChartTitle: document.getElementById("publisherChartTitle"),
  publisherChartSubTitle: document.getElementById("publisherChartSubTitle")
};

nodes.fileInput.addEventListener("change", handleLocalFileUpload);
nodes.loadPresetBtn.addEventListener("click", loadPresetFromManifest);
initTableSorting();

initPresetFiles();

async function initPresetFiles() {
  try {
    const response = await fetch("./excel/manifest.json", { cache: "no-store" });
    if (!response.ok) {
      return;
    }

    const manifest = await response.json();
    const files = normalizeManifestFiles(manifest);
    if (!files.length) {
      return;
    }

    nodes.presetFileSelect.innerHTML = "";
    files.forEach((item) => {
      const option = document.createElement("option");
      option.value = item.path;
      option.textContent = item.name;
      nodes.presetFileSelect.appendChild(option);
    });

    nodes.presetPanel.classList.remove("hidden");
  } catch (_) {
    // Ignore manifest loading failure in static environments.
  }
}

function normalizeManifestFiles(manifest) {
  const source = Array.isArray(manifest) ? manifest : Array.isArray(manifest?.files) ? manifest.files : [];
  return source
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
}

async function loadPresetFromManifest() {
  const filePath = nodes.presetFileSelect.value;
  if (!filePath) {
    return;
  }

  try {
    const response = await fetch(filePath);
    if (!response.ok) {
      throw new Error("读取仓库文件失败");
    }

    const fileName = getFileBaseName(filePath);
    const workbook = await readWorkbookFromResponse(response, filePath);
    nodes.fileName.textContent = `已加载仓库文件: ${fileName}`;
    processWorkbook(workbook, fileName);
  } catch (error) {
    window.alert(`加载失败：${error.message}`);
  }
}

function handleLocalFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  nodes.fileName.textContent = `已上传: ${file.name}`;
  const reader = new FileReader();
  reader.onload = function onload(e) {
    try {
      const arrayBuffer = e.target.result;
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      processWorkbook(workbook, file.name);
    } catch (error) {
      window.alert(`解析失败：${error.message}`);
    }
  };
  reader.readAsArrayBuffer(file);
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

function processWorkbook(workbook, fileName) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  if (!rows.length) {
    window.alert("文件内容为空，请检查上传文件。");
    return;
  }

  const headers = Object.keys(rows[0]);
  const columnMap = resolveColumns(headers);
  const missingMetrics = METRIC_CONFIG.filter((item) => !columnMap[item.key]);

  if (missingMetrics.length) {
    const labels = missingMetrics.map((item) => item.label).join("、");
    window.alert(`缺少必须列：${labels}。请检查表头后重试。`);
    return;
  }

  influencerState.fileTitle = stripFileExtension(fileName);
  influencerState.sheetName = sheetName;
  influencerState.rawRows = rows;
  influencerState.columnMap = columnMap;
  influencerState.normalizedRows = normalizeRows(rows, columnMap);

  updateSortHeaderState();
  renderDashboard();
  showDashboardPanels();
}

function resolveColumns(headers) {
  const map = {};
  const headerMap = new Map(headers.map((header) => [normalizeHeader(header), header]));

  [...METRIC_CONFIG, ...EXTRA_FIELD_CONFIG].forEach((item) => {
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
        const normalizedHeader = normalizeHeader(header);
        return item.aliases.some((alias) => normalizedHeader.includes(normalizeHeader(alias)));
      });
      matched = fuzzy || "";
    }

    map[item.key] = matched;
  });

  return map;
}

function normalizeRows(rows, columnMap) {
  return rows.map((row) => ({
    adGroup: columnMap.adGroup ? String(row[columnMap.adGroup] || "") : "",
    channel: columnMap.channel ? String(row[columnMap.channel] || "") : "",
    publisher: columnMap.publisher ? String(row[columnMap.publisher] || "") : "",
    clicks: toNumber(row[columnMap.clicks]),
    dpv: toNumber(row[columnMap.dpv]),
    atc: toNumber(row[columnMap.atc]),
    revenue: toNumber(row[columnMap.revenue]),
    units: toNumber(row[columnMap.units]),
    bonus: toNumber(row[columnMap.bonus])
  }));
}

function renderDashboard() {
  const totals = calculateTotals(influencerState.normalizedRows);

  setText("sumClicks", formatNumber(totals.clicks, 0));
  setText("sumDpv", formatNumber(totals.dpv, 0));
  setText("sumAtc", formatNumber(totals.atc, 0));
  setText("sumRevenue", formatNumber(totals.revenue, 2));
  setText("sumUnits", formatNumber(totals.units, 0));
  setText("sumBonus", formatNumber(totals.bonus, 2));

  renderMetricBarChart(totals);
  renderPublisherBarChart(influencerState.normalizedRows);
  renderTable(influencerState.normalizedRows, totals);
}

function calculateTotals(rows) {
  return rows.reduce(
    (sum, row) => ({
      clicks: sum.clicks + row.clicks,
      dpv: sum.dpv + row.dpv,
      atc: sum.atc + row.atc,
      revenue: sum.revenue + row.revenue,
      units: sum.units + row.units,
      bonus: sum.bonus + row.bonus
    }),
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

function renderMetricBarChart(totals) {
  const labels = METRIC_CONFIG.map((item) => item.label);
  const data = [totals.clicks, totals.dpv, totals.atc, totals.revenue, totals.units, totals.bonus];
  const title = `${influencerState.fileTitle} - 核心指标总量柱状图`;

  nodes.metricChartTitle.textContent = title;
  nodes.metricChartSubTitle.textContent = `工作表：${influencerState.sheetName}｜记录数：${influencerState.normalizedRows.length}`;

  if (influencerState.charts.metric) {
    influencerState.charts.metric.destroy();
  }

  influencerState.charts.metric = new Chart(document.getElementById("metricBarChart"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "汇总值",
          data,
          backgroundColor: [
            "rgba(227, 188, 109, 0.86)",
            "rgba(211, 168, 84, 0.86)",
            "rgba(196, 149, 64, 0.86)",
            "rgba(173, 128, 45, 0.86)",
            "rgba(154, 110, 34, 0.86)",
            "rgba(235, 202, 136, 0.86)"
          ],
          borderRadius: 8
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }
      },
      scales: {
        x: {
          ticks: { color: "#ba975f" },
          grid: { color: "rgba(233, 191, 93, 0.12)" }
        },
        y: {
          beginAtZero: true,
          ticks: { color: "#ba975f" },
          grid: { color: "rgba(233, 191, 93, 0.12)" }
        }
      }
    }
  });
}

function renderPublisherBarChart(rows) {
  const grouped = new Map();
  rows.forEach((row) => {
    const name = row.publisher || "未命名出版商";
    grouped.set(name, (grouped.get(name) || 0) + row.clicks);
  });

  const topPublishers = Array.from(grouped.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);

  const labels = topPublishers.map((item) => item[0]);
  const values = topPublishers.map((item) => item[1]);

  nodes.publisherChartTitle.textContent = `${influencerState.fileTitle} - Top 发布商点击量`;
  nodes.publisherChartSubTitle.textContent = `按“出版商”聚合，展示前 ${topPublishers.length} 名`;

  if (influencerState.charts.publisher) {
    influencerState.charts.publisher.destroy();
  }

  influencerState.charts.publisher = new Chart(document.getElementById("publisherBarChart"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "点击量",
          data: values,
          backgroundColor: "rgba(217, 177, 99, 0.82)",
          borderRadius: 8
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: "y",
      plugins: {
        legend: { display: false }
      },
      scales: {
        x: {
          beginAtZero: true,
          ticks: { color: "#ba975f" },
          grid: { color: "rgba(233, 191, 93, 0.12)" }
        },
        y: {
          ticks: { color: "#ba975f" },
          grid: { color: "rgba(233, 191, 93, 0.12)" }
        }
      }
    }
  });
}

function renderTable(rows, totals) {
  const shownRows = sortRows(rows, influencerState.sort.key, influencerState.sort.direction);
  nodes.tableBody.innerHTML = shownRows
    .map(
      (row) => `
      <tr>
        <td>${escapeHtml(row.adGroup)}</td>
        <td>${escapeHtml(row.channel)}</td>
        <td>${escapeHtml(row.publisher)}</td>
        <td>${formatNumber(row.clicks, 0)}</td>
        <td>${formatNumber(row.dpv, 0)}</td>
        <td>${formatNumber(row.atc, 0)}</td>
        <td>${formatNumber(row.revenue, 2)}</td>
        <td>${formatNumber(row.units, 0)}</td>
        <td>${formatNumber(row.bonus, 2)}</td>
      </tr>
    `
    )
    .join("");

  const sortLabel = getSortLabel(influencerState.sort.key, influencerState.sort.direction);
  nodes.tableSummary.textContent = `${influencerState.fileTitle}｜共 ${rows.length} 行｜当前排序：${sortLabel}｜点击量合计 ${formatNumber(
    totals.clicks,
    0
  )}`;
}

function showDashboardPanels() {
  nodes.emptyState.classList.add("hidden");
  nodes.kpiPanel.classList.remove("hidden");
  nodes.chartPanel.classList.remove("hidden");
  nodes.tablePanel.classList.remove("hidden");
}

function initTableSorting() {
  const sortableHeaders = document.querySelectorAll("th.sortable");
  sortableHeaders.forEach((header) => {
    header.addEventListener("click", () => {
      const key = header.dataset.sortKey;
      if (!key) {
        return;
      }
      if (influencerState.sort.key === key) {
        influencerState.sort.direction = influencerState.sort.direction === "asc" ? "desc" : "asc";
      } else {
        influencerState.sort.key = key;
        influencerState.sort.direction = isNumericSortKey(key) ? "desc" : "asc";
      }
      updateSortHeaderState();
      if (!influencerState.normalizedRows.length) {
        return;
      }
      const totals = calculateTotals(influencerState.normalizedRows);
      renderTable(influencerState.normalizedRows, totals);
    });
  });
  updateSortHeaderState();
}

function updateSortHeaderState() {
  const sortableHeaders = document.querySelectorAll("th.sortable");
  sortableHeaders.forEach((header) => {
    const key = header.dataset.sortKey;
    const baseText = header.textContent.replace(/\s*[↑↓]$/, "");
    if (key === influencerState.sort.key) {
      const arrow = influencerState.sort.direction === "asc" ? "↑" : "↓";
      header.textContent = `${baseText} ${arrow}`;
      header.classList.add("active-sort");
    } else {
      header.textContent = baseText;
      header.classList.remove("active-sort");
    }
  });
}

function sortRows(rows, key, direction) {
  const cloned = rows.slice();
  cloned.sort((a, b) => {
    const av = a[key];
    const bv = b[key];
    let result = 0;
    if (isNumericSortKey(key)) {
      result = (Number(av) || 0) - (Number(bv) || 0);
    } else {
      result = String(av || "").localeCompare(String(bv || ""), "zh-CN");
    }
    return direction === "asc" ? result : -result;
  });
  return cloned;
}

function isNumericSortKey(key) {
  return ["clicks", "dpv", "atc", "revenue", "units", "bonus"].includes(key);
}

function getSortLabel(key, direction) {
  const dict = {
    adGroup: "广告组",
    channel: "Channel",
    publisher: "出版商",
    clicks: "点击量",
    dpv: "总 DPV",
    atc: "ATC 总计",
    revenue: "购买总额",
    units: "商品销量总计",
    bonus: "品牌引流奖励计划"
  };
  const label = dict[key] || key;
  const dirText = direction === "asc" ? "升序" : "降序";
  return `${label}（${dirText}）`;
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

function getFileBaseName(pathValue) {
  const path = String(pathValue || "");
  const parts = path.split("/");
  return decodeURIComponent(parts[parts.length - 1] || path);
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

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

