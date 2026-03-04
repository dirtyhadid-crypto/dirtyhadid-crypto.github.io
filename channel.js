const channelState = {
  headers: [],
  rows: [],
  mappedRows: [],
  chart: null
};

const bodyConfig = document.body.dataset;
const metricLabel = bodyConfig.metricLabel || "指标值";
const metricKeywords = (bodyConfig.metricKeywords || "")
  .split(",")
  .map((item) => item.trim())
  .filter(Boolean);

const channelNodes = {
  fileInput: document.getElementById("fileInput"),
  fileName: document.getElementById("fileName"),
  mappingPanel: document.getElementById("mappingPanel"),
  filterPanel: document.getElementById("filterPanel"),
  kpiPanel: document.getElementById("kpiPanel"),
  chartPanel: document.getElementById("chartPanel"),
  tablePanel: document.getElementById("tablePanel"),
  emptyState: document.getElementById("emptyState"),
  rangeSelect: document.getElementById("rangeSelect"),
  applyMappingBtn: document.getElementById("applyMappingBtn"),
  dateColumn: document.getElementById("dateColumn"),
  valueColumn: document.getElementById("valueColumn"),
  tableBody: document.getElementById("metricTableBody")
};

document.querySelectorAll("[data-metric-label]").forEach((node) => {
  node.textContent = metricLabel;
});

channelNodes.fileInput.addEventListener("change", handleChannelFileUpload);
channelNodes.applyMappingBtn.addEventListener("click", applyChannelMapping);
channelNodes.rangeSelect.addEventListener("change", renderChannelDashboard);

function handleChannelFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  channelNodes.fileName.textContent = `已上传: ${file.name}`;
  const reader = new FileReader();

  reader.onload = function onload(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      if (!rows.length) {
        throw new Error("文件没有数据");
      }

      channelState.rows = rows;
      channelState.headers = Object.keys(rows[0]);
      initChannelSelectors(channelState.headers);
      channelNodes.mappingPanel.classList.remove("hidden");
      showChannelSetupOnly();
    } catch (error) {
      window.alert(`读取失败：${error.message}`);
    }
  };

  reader.readAsArrayBuffer(file);
}

function initChannelSelectors(headers) {
  [channelNodes.dateColumn, channelNodes.valueColumn].forEach((selectNode) => {
    selectNode.innerHTML = "";

    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = "请选择列";
    selectNode.appendChild(emptyOption);

    headers.forEach((header) => {
      const option = document.createElement("option");
      option.value = header;
      option.textContent = header;
      selectNode.appendChild(option);
    });
  });

  channelNodes.dateColumn.value = guessHeader(headers, ["date", "day", "日期", "时间", "dt"]);
  channelNodes.valueColumn.value = guessHeader(headers, [
    ...metricKeywords,
    "流量",
    "traffic",
    "visits",
    "value"
  ]);
}

function guessHeader(headers, keywords) {
  return (
    headers.find((header) => {
      const lowerName = String(header).toLowerCase();
      return keywords.some((keyword) => lowerName.includes(String(keyword).toLowerCase()));
    }) || ""
  );
}

function applyChannelMapping() {
  const dateKey = channelNodes.dateColumn.value;
  const valueKey = channelNodes.valueColumn.value;

  if (!dateKey || !valueKey) {
    window.alert("请选择日期列和指标列。");
    return;
  }

  channelState.mappedRows = normalizeChannelRows(channelState.rows, dateKey, valueKey);
  if (!channelState.mappedRows.length) {
    window.alert("没有解析到有效日期，请检查日期格式。");
    return;
  }

  showChannelDashboard();
  renderChannelDashboard();
}

function normalizeChannelRows(rows, dateKey, valueKey) {
  const dayMap = new Map();

  rows.forEach((row) => {
    const dateLabel = normalizeDateLabel(row[dateKey]);
    if (!dateLabel) {
      return;
    }

    const currentValue = dayMap.get(dateLabel) || { date: dateLabel, value: 0 };
    currentValue.value += toNumber(row[valueKey]);
    dayMap.set(dateLabel, currentValue);
  });

  return Array.from(dayMap.values()).sort((a, b) => a.date.localeCompare(b.date));
}

function renderChannelDashboard() {
  const rows = getChannelFilteredRows();
  if (!rows.length) {
    return;
  }

  renderChannelKpis(rows);
  renderChannelChart(rows);
  renderChannelTable(rows);
}

function getChannelFilteredRows() {
  const range = channelNodes.rangeSelect.value;
  if (range === "all") {
    return channelState.mappedRows;
  }
  const count = parseInt(range, 10);
  if (!Number.isFinite(count)) {
    return channelState.mappedRows;
  }
  return channelState.mappedRows.slice(-count);
}

function renderChannelKpis(rows) {
  const latest = rows[rows.length - 1];
  const previous = rows[rows.length - 2] || null;
  const total = rows.reduce((sum, row) => sum + row.value, 0);
  const lastSevenRows = rows.slice(-7);
  const average =
    lastSevenRows.reduce((sum, row) => sum + row.value, 0) / Math.max(lastSevenRows.length, 1);
  const peak = rows.reduce((max, row) => (row.value > max.value ? row : max), rows[0]);

  document.getElementById("latestValue").textContent = formatNumber(latest.value);
  const changeNode = document.getElementById("latestChange");
  changeNode.textContent = formatChange(latest.value, previous?.value);
  changeNode.classList.remove("change-up", "change-down", "change-flat");
  if (!Number.isFinite(previous?.value)) {
    changeNode.classList.add("change-flat");
  } else if (latest.value > previous.value) {
    changeNode.classList.add("change-up");
  } else if (latest.value < previous.value) {
    changeNode.classList.add("change-down");
  } else {
    changeNode.classList.add("change-flat");
  }

  document.getElementById("avgValue").textContent = formatNumber(average);
  document.getElementById("totalValue").textContent = formatNumber(total);
  document.getElementById("peakValue").textContent = formatNumber(peak.value);
  document.getElementById("peakDate").textContent = `峰值日期: ${peak.date}`;
  document.getElementById("rowCountValue").textContent = `${rows.length} 天`;
}

function renderChannelChart(rows) {
  const labels = rows.map((row) => row.date);
  const values = rows.map((row) => row.value);

  if (channelState.chart) {
    channelState.chart.destroy();
  }

  channelState.chart = new Chart(document.getElementById("metricChart"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: metricLabel,
          data: values,
          borderColor: "#3f73ff",
          backgroundColor: "rgba(63, 115, 255, 0.24)",
          fill: true,
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
          labels: { boxWidth: 10, boxHeight: 10 }
        }
      },
      scales: {
        y: { beginAtZero: true }
      }
    }
  });
}

function renderChannelTable(rows) {
  channelNodes.tableBody.innerHTML = rows
    .slice()
    .reverse()
    .map(
      (row) => `
      <tr>
        <td>${row.date}</td>
        <td>${formatNumber(row.value)}</td>
      </tr>
    `
    )
    .join("");
}

function showChannelSetupOnly() {
  channelNodes.emptyState.classList.remove("hidden");
  channelNodes.filterPanel.classList.add("hidden");
  channelNodes.kpiPanel.classList.add("hidden");
  channelNodes.chartPanel.classList.add("hidden");
  channelNodes.tablePanel.classList.add("hidden");
}

function showChannelDashboard() {
  channelNodes.emptyState.classList.add("hidden");
  channelNodes.filterPanel.classList.remove("hidden");
  channelNodes.kpiPanel.classList.remove("hidden");
  channelNodes.chartPanel.classList.remove("hidden");
  channelNodes.tablePanel.classList.remove("hidden");
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

function formatNumber(value) {
  const safeValue = Number.isFinite(value) ? value : 0;
  return new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 0 }).format(safeValue);
}

function formatChange(current, previous) {
  if (!Number.isFinite(previous)) {
    return "较前一天: -";
  }
  if (previous === 0) {
    return current === 0 ? "较前一天: 0%" : "较前一天: 新增";
  }
  const ratio = ((current - previous) / Math.abs(previous)) * 100;
  const sign = ratio > 0 ? "+" : "";
  return `较前一天: ${sign}${ratio.toFixed(1)}%`;
}
