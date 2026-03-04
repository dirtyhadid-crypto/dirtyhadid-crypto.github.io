const overviewState = {
  headers: [],
  rows: [],
  mappedRows: [],
  chart: null,
  mapping: null
};

const overviewNodes = {
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
  tableBody: document.getElementById("overviewTableBody")
};

const overviewSelectors = {
  date: document.getElementById("dateColumn"),
  amazon: document.getElementById("amazonColumn"),
  influencer: document.getElementById("influencerColumn"),
  social: document.getElementById("socialColumn"),
  media: document.getElementById("mediaColumn")
};

const keywordMap = {
  date: ["date", "day", "日期", "时间", "dt"],
  amazon: ["amazon", "amz", "亚马逊"],
  influencer: ["红人", "kol", "达人", "influencer", "creator"],
  social: ["社媒", "social", "facebook", "instagram", "tiktok", "youtube"],
  media: ["媒体", "media", "press", "news", "公关"]
};

overviewNodes.fileInput.addEventListener("change", handleOverviewFileUpload);
overviewNodes.applyMappingBtn.addEventListener("click", applyOverviewMapping);
overviewNodes.rangeSelect.addEventListener("change", renderOverviewDashboard);

function handleOverviewFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  overviewNodes.fileName.textContent = `已上传: ${file.name}`;
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

      overviewState.rows = rows;
      overviewState.headers = Object.keys(rows[0]);
      initOverviewSelectors(overviewState.headers);
      overviewNodes.mappingPanel.classList.remove("hidden");
      showOverviewSetupOnly();
    } catch (error) {
      window.alert(`读取失败：${error.message}`);
    }
  };

  reader.readAsArrayBuffer(file);
}

function initOverviewSelectors(headers) {
  Object.values(overviewSelectors).forEach((selectNode) => {
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

  Object.entries(overviewSelectors).forEach(([field, selectNode]) => {
    selectNode.value = guessHeader(overviewState.headers, keywordMap[field]);
  });
}

function guessHeader(headers, keywords) {
  return (
    headers.find((header) => {
      const name = String(header).toLowerCase();
      return keywords.some((keyword) => name.includes(keyword.toLowerCase()));
    }) || ""
  );
}

function applyOverviewMapping() {
  const mapping = {
    date: overviewSelectors.date.value,
    amazon: overviewSelectors.amazon.value,
    influencer: overviewSelectors.influencer.value,
    social: overviewSelectors.social.value,
    media: overviewSelectors.media.value
  };

  if (!mapping.date) {
    window.alert("请至少选择日期列。");
    return;
  }

  if (!mapping.amazon && !mapping.influencer && !mapping.social && !mapping.media) {
    window.alert("请至少选择一个流量字段（亚马逊/红人/社媒/媒体）。");
    return;
  }

  overviewState.mapping = mapping;
  overviewState.mappedRows = normalizeOverviewRows(overviewState.rows, mapping);

  if (!overviewState.mappedRows.length) {
    window.alert("没有解析到有效日期，请检查日期格式。");
    return;
  }

  showOverviewDashboard();
  renderOverviewDashboard();
}

function normalizeOverviewRows(rows, mapping) {
  const dayMap = new Map();

  rows.forEach((row) => {
    const dateLabel = normalizeDateLabel(row[mapping.date]);
    if (!dateLabel) {
      return;
    }

    const current = dayMap.get(dateLabel) || {
      date: dateLabel,
      amazon: 0,
      influencer: 0,
      social: 0,
      media: 0
    };

    current.amazon += mapping.amazon ? toNumber(row[mapping.amazon]) : 0;
    current.influencer += mapping.influencer ? toNumber(row[mapping.influencer]) : 0;
    current.social += mapping.social ? toNumber(row[mapping.social]) : 0;
    current.media += mapping.media ? toNumber(row[mapping.media]) : 0;

    dayMap.set(dateLabel, current);
  });

  return Array.from(dayMap.values())
    .sort((a, b) => a.date.localeCompare(b.date))
    .map((item) => ({
      ...item,
      total: item.amazon + item.influencer + item.social + item.media
    }));
}

function renderOverviewDashboard() {
  const rows = getOverviewFilteredRows();
  if (!rows.length) {
    return;
  }

  renderOverviewKpis(rows);
  renderOverviewChart(rows);
  renderOverviewTable(rows);
}

function getOverviewFilteredRows() {
  const range = overviewNodes.rangeSelect.value;
  if (range === "all") {
    return overviewState.mappedRows;
  }
  const count = parseInt(range, 10);
  if (!Number.isFinite(count)) {
    return overviewState.mappedRows;
  }
  return overviewState.mappedRows.slice(-count);
}

function renderOverviewKpis(rows) {
  const latest = rows[rows.length - 1];
  const previous = rows[rows.length - 2] || null;

  setKpiValue("kpiTotalValue", "kpiTotalChange", latest.total, previous?.total);
  setKpiValue("kpiAmazonValue", "kpiAmazonChange", latest.amazon, previous?.amazon);
  setKpiValue("kpiInfluencerValue", "kpiInfluencerChange", latest.influencer, previous?.influencer);
  setKpiValue("kpiSocialValue", "kpiSocialChange", latest.social, previous?.social);
  setKpiValue("kpiMediaValue", "kpiMediaChange", latest.media, previous?.media);
}

function setKpiValue(valueId, changeId, currentValue, previousValue) {
  const valueNode = document.getElementById(valueId);
  const changeNode = document.getElementById(changeId);

  valueNode.textContent = formatNumber(currentValue);
  const changeText = formatChange(currentValue, previousValue);
  changeNode.textContent = changeText;
  changeNode.classList.remove("change-up", "change-down", "change-flat");

  if (!Number.isFinite(previousValue)) {
    changeNode.classList.add("change-flat");
    return;
  }

  if (currentValue > previousValue) {
    changeNode.classList.add("change-up");
  } else if (currentValue < previousValue) {
    changeNode.classList.add("change-down");
  } else {
    changeNode.classList.add("change-flat");
  }
}

function renderOverviewChart(rows) {
  const labels = rows.map((item) => item.date);
  const datasets = [
    {
      label: "总流量",
      data: rows.map((item) => item.total),
      borderColor: "#2f5bde",
      backgroundColor: "rgba(47, 91, 222, 0.25)",
      fill: true,
      tension: 0.3,
      pointRadius: 2
    }
  ];

  if (overviewState.mapping.amazon) {
    datasets.push({
      label: "亚马逊",
      data: rows.map((item) => item.amazon),
      borderColor: "#5e8cff",
      backgroundColor: "rgba(94, 140, 255, 0.12)",
      tension: 0.3,
      pointRadius: 2
    });
  }
  if (overviewState.mapping.influencer) {
    datasets.push({
      label: "红人",
      data: rows.map((item) => item.influencer),
      borderColor: "#db5f8b",
      backgroundColor: "rgba(219, 95, 139, 0.12)",
      tension: 0.3,
      pointRadius: 2
    });
  }
  if (overviewState.mapping.social) {
    datasets.push({
      label: "社媒",
      data: rows.map((item) => item.social),
      borderColor: "#0d9bc5",
      backgroundColor: "rgba(13, 155, 197, 0.12)",
      tension: 0.3,
      pointRadius: 2
    });
  }
  if (overviewState.mapping.media) {
    datasets.push({
      label: "媒体",
      data: rows.map((item) => item.media),
      borderColor: "#8a62ff",
      backgroundColor: "rgba(138, 98, 255, 0.12)",
      tension: 0.3,
      pointRadius: 2
    });
  }

  if (overviewState.chart) {
    overviewState.chart.destroy();
  }

  overviewState.chart = new Chart(document.getElementById("overviewChart"), {
    type: "line",
    data: { labels, datasets },
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

function renderOverviewTable(rows) {
  overviewNodes.tableBody.innerHTML = rows
    .slice()
    .reverse()
    .map(
      (row) => `
      <tr>
        <td>${row.date}</td>
        <td>${formatNumber(row.total)}</td>
        <td>${formatNumber(row.amazon)}</td>
        <td>${formatNumber(row.influencer)}</td>
        <td>${formatNumber(row.social)}</td>
        <td>${formatNumber(row.media)}</td>
      </tr>
    `
    )
    .join("");
}

function showOverviewSetupOnly() {
  overviewNodes.emptyState.classList.remove("hidden");
  overviewNodes.filterPanel.classList.add("hidden");
  overviewNodes.kpiPanel.classList.add("hidden");
  overviewNodes.chartPanel.classList.add("hidden");
  overviewNodes.tablePanel.classList.add("hidden");
}

function showOverviewDashboard() {
  overviewNodes.emptyState.classList.add("hidden");
  overviewNodes.filterPanel.classList.remove("hidden");
  overviewNodes.kpiPanel.classList.remove("hidden");
  overviewNodes.chartPanel.classList.remove("hidden");
  overviewNodes.tablePanel.classList.remove("hidden");
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
