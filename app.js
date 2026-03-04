const state = {
  headers: [],
  rows: [],
  mappedRows: [],
  charts: {
    sales: null,
    traffic: null
  }
};

const fileInput = document.getElementById("fileInput");
const fileName = document.getElementById("fileName");
const mappingPanel = document.getElementById("mappingPanel");
const filterPanel = document.getElementById("filterPanel");
const kpiPanel = document.getElementById("kpiPanel");
const chartPanel = document.getElementById("chartPanel");
const tablePanel = document.getElementById("tablePanel");
const emptyState = document.getElementById("emptyState");
const rangeSelect = document.getElementById("rangeSelect");
const applyMappingBtn = document.getElementById("applyMappingBtn");

const mappingSelectors = {
  date: document.getElementById("dateColumn"),
  amazonSales: document.getElementById("amazonColumn"),
  siteSales: document.getElementById("siteColumn"),
  influencerTraffic: document.getElementById("influencerColumn"),
  socialTraffic: document.getElementById("socialColumn")
};

const keywordMap = {
  date: ["date", "day", "日期", "时间", "report", "dt"],
  amazonSales: ["amazon", "amz", "亚马逊", "amazon销量", "amazon sales", "platform"],
  siteSales: ["独立站", "site", "shopify", "store", "website", "d2c"],
  influencerTraffic: ["红人", "kol", "influencer", "creator", "达人"],
  socialTraffic: ["社媒", "social", "instagram", "facebook", "tiktok", "youtube", "媒体"]
};

fileInput.addEventListener("change", handleFileUpload);
applyMappingBtn.addEventListener("click", applyMappingAndRender);
rangeSelect.addEventListener("change", renderAll);

function handleFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  fileName.textContent = `已上传: ${file.name}`;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const firstSheet = workbook.Sheets[firstSheetName];
      const json = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

      if (!json.length) {
        throw new Error("文件内容为空");
      }

      state.rows = json;
      state.headers = Object.keys(json[0]);
      initMappingSelectors(state.headers);
      mappingPanel.classList.remove("hidden");
      showOnlySetupView();
    } catch (err) {
      alert(`解析失败：${err.message}`);
    }
  };
  reader.readAsArrayBuffer(file);
}

function initMappingSelectors(headers) {
  Object.values(mappingSelectors).forEach((selector) => {
    selector.innerHTML = "";
    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = "请选择列";
    selector.appendChild(emptyOption);

    headers.forEach((header) => {
      const option = document.createElement("option");
      option.value = header;
      option.textContent = header;
      selector.appendChild(option);
    });
  });

  Object.entries(mappingSelectors).forEach(([field, selector]) => {
    selector.value = guessHeaderByKeyword(headers, keywordMap[field]);
  });
}

function guessHeaderByKeyword(headers, keywords) {
  const matched = headers.find((header) => {
    const normalized = header.toLowerCase();
    return keywords.some((key) => normalized.includes(key.toLowerCase()));
  });
  return matched || "";
}

function applyMappingAndRender() {
  const mapping = {
    date: mappingSelectors.date.value,
    amazonSales: mappingSelectors.amazonSales.value,
    siteSales: mappingSelectors.siteSales.value,
    influencerTraffic: mappingSelectors.influencerTraffic.value,
    socialTraffic: mappingSelectors.socialTraffic.value
  };

  if (!mapping.date || !mapping.amazonSales || !mapping.siteSales) {
    alert("请至少选择 日期、Amazon销量、独立站销量 三个字段。");
    return;
  }

  state.mappedRows = normalizeRows(state.rows, mapping);
  if (!state.mappedRows.length) {
    alert("没有成功解析到有效日期数据，请检查日期列格式。");
    return;
  }

  showDashboardView();
  renderAll();
}

function normalizeRows(rows, mapping) {
  const dayMap = new Map();

  rows.forEach((row) => {
    const dateLabel = normalizeDateLabel(row[mapping.date]);
    if (!dateLabel) {
      return;
    }

    const current = dayMap.get(dateLabel) || {
      date: dateLabel,
      amazonSales: 0,
      siteSales: 0,
      influencerTraffic: 0,
      socialTraffic: 0
    };

    current.amazonSales += toNumber(row[mapping.amazonSales]);
    current.siteSales += toNumber(row[mapping.siteSales]);
    current.influencerTraffic += mapping.influencerTraffic ? toNumber(row[mapping.influencerTraffic]) : 0;
    current.socialTraffic += mapping.socialTraffic ? toNumber(row[mapping.socialTraffic]) : 0;
    dayMap.set(dateLabel, current);
  });

  return Array.from(dayMap.values())
    .sort((a, b) => a.date.localeCompare(b.date))
    .map((item) => ({
      ...item,
      totalSales: item.amazonSales + item.siteSales
    }));
}

function normalizeDateLabel(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  let dateObj = null;

  if (typeof value === "number") {
    const excelDate = XLSX.SSF.parse_date_code(value);
    if (excelDate) {
      dateObj = new Date(excelDate.y, excelDate.m - 1, excelDate.d);
    }
  } else if (value instanceof Date) {
    dateObj = new Date(value.getFullYear(), value.getMonth(), value.getDate());
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
      dateObj = new Date(y, m - 1, d);
    } else {
      const parsed = new Date(cleaned);
      if (!Number.isNaN(parsed.getTime())) {
        dateObj = new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
      }
    }
  }

  if (!dateObj) {
    return "";
  }

  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, "0");
  const d = String(dateObj.getDate()).padStart(2, "0");
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

function renderAll() {
  const rows = getFilteredRows();
  if (!rows.length) {
    return;
  }
  renderKpis(rows);
  renderCharts(rows);
  renderTable(rows);
}

function getFilteredRows() {
  const rows = state.mappedRows;
  const range = rangeSelect.value;
  if (range === "all") {
    return rows;
  }
  const count = parseInt(range, 10);
  if (!Number.isFinite(count)) {
    return rows;
  }
  return rows.slice(-count);
}

function renderKpis(rows) {
  const latest = rows[rows.length - 1];
  const previous = rows[rows.length - 2] || null;

  setKpi("kpiTotalSales", "kpiTotalChange", latest.totalSales, previous?.totalSales);
  setKpi("kpiAmazonSales", "kpiAmazonChange", latest.amazonSales, previous?.amazonSales);
  setKpi("kpiSiteSales", "kpiSiteChange", latest.siteSales, previous?.siteSales);
  setKpi(
    "kpiInfluencerTraffic",
    "kpiInfluencerChange",
    latest.influencerTraffic,
    previous?.influencerTraffic
  );
  setKpi("kpiSocialTraffic", "kpiSocialChange", latest.socialTraffic, previous?.socialTraffic);
}

function setKpi(valueId, changeId, value, previousValue) {
  const valueNode = document.getElementById(valueId);
  const changeNode = document.getElementById(changeId);
  valueNode.textContent = formatNumber(value);

  const text = formatChange(value, previousValue);
  changeNode.textContent = text;
  changeNode.classList.remove("change-up", "change-down", "change-flat");

  if (!Number.isFinite(previousValue) || previousValue === null) {
    changeNode.classList.add("change-flat");
    return;
  }
  if (value > previousValue) {
    changeNode.classList.add("change-up");
  } else if (value < previousValue) {
    changeNode.classList.add("change-down");
  } else {
    changeNode.classList.add("change-flat");
  }
}

function renderCharts(rows) {
  const labels = rows.map((item) => item.date);

  const salesData = {
    labels,
    datasets: [
      {
        label: "Amazon 销量",
        data: rows.map((item) => item.amazonSales),
        borderColor: "#1a8574",
        backgroundColor: "rgba(26,133,116,0.2)",
        tension: 0.25,
        pointRadius: 2
      },
      {
        label: "独立站销量",
        data: rows.map((item) => item.siteSales),
        borderColor: "#2457d6",
        backgroundColor: "rgba(36,87,214,0.18)",
        tension: 0.25,
        pointRadius: 2
      },
      {
        label: "总销量",
        data: rows.map((item) => item.totalSales),
        borderColor: "#cc6e1d",
        backgroundColor: "rgba(204,110,29,0.18)",
        borderDash: [4, 4],
        tension: 0.25,
        pointRadius: 2
      }
    ]
  };

  const trafficData = {
    labels,
    datasets: [
      {
        label: "红人流量",
        data: rows.map((item) => item.influencerTraffic),
        borderColor: "#be3a6e",
        backgroundColor: "rgba(190,58,110,0.18)",
        tension: 0.3,
        pointRadius: 2
      },
      {
        label: "社媒流量",
        data: rows.map((item) => item.socialTraffic),
        borderColor: "#1f8c9f",
        backgroundColor: "rgba(31,140,159,0.16)",
        tension: 0.3,
        pointRadius: 2
      }
    ]
  };

  const commonOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        labels: {
          boxWidth: 10,
          boxHeight: 10
        }
      }
    },
    scales: {
      y: {
        beginAtZero: true
      }
    }
  };

  if (state.charts.sales) {
    state.charts.sales.destroy();
  }
  if (state.charts.traffic) {
    state.charts.traffic.destroy();
  }

  state.charts.sales = new Chart(document.getElementById("salesChart"), {
    type: "line",
    data: salesData,
    options: commonOptions
  });

  state.charts.traffic = new Chart(document.getElementById("trafficChart"), {
    type: "line",
    data: trafficData,
    options: commonOptions
  });
}

function renderTable(rows) {
  const tableBody = document.getElementById("dataTableBody");
  tableBody.innerHTML = rows
    .slice()
    .reverse()
    .map(
      (row) => `
        <tr>
          <td>${row.date}</td>
          <td>${formatNumber(row.amazonSales)}</td>
          <td>${formatNumber(row.siteSales)}</td>
          <td>${formatNumber(row.totalSales)}</td>
          <td>${formatNumber(row.influencerTraffic)}</td>
          <td>${formatNumber(row.socialTraffic)}</td>
        </tr>
      `
    )
    .join("");
}

function formatNumber(value) {
  const safeValue = Number.isFinite(value) ? value : 0;
  return new Intl.NumberFormat("zh-CN", {
    maximumFractionDigits: 0
  }).format(safeValue);
}

function formatChange(current, previous) {
  if (!Number.isFinite(previous) || previous === null) {
    return "较昨日: -";
  }
  if (previous === 0) {
    return current === 0 ? "较昨日: 0%" : "较昨日: 新增";
  }
  const ratio = ((current - previous) / Math.abs(previous)) * 100;
  const symbol = ratio > 0 ? "+" : "";
  return `较昨日: ${symbol}${ratio.toFixed(1)}%`;
}

function showOnlySetupView() {
  emptyState.classList.remove("hidden");
  filterPanel.classList.add("hidden");
  kpiPanel.classList.add("hidden");
  chartPanel.classList.add("hidden");
  tablePanel.classList.add("hidden");
}

function showDashboardView() {
  emptyState.classList.add("hidden");
  filterPanel.classList.remove("hidden");
  kpiPanel.classList.remove("hidden");
  chartPanel.classList.remove("hidden");
  tablePanel.classList.remove("hidden");
}
