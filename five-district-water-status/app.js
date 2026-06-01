const message = document.querySelector("#message");
const fileInput = document.querySelector("#fileInput");
const uploadMeta = document.querySelector("#uploadMeta");
const datePanel = document.querySelector("#datePanel");
const dateMeta = document.querySelector("#dateMeta");
const yearSelect = document.querySelector("#yearSelect");
const monthSelect = document.querySelector("#monthSelect");
const daySelect = document.querySelector("#daySelect");

const metricTargets = {
  "原水總取水量": document.querySelector("#metricRaw"),
  "總出水量": document.querySelector("#metricOutflow"),
  "支(受)援水量": document.querySelector("#metricSupport"),
  "計算後供水量": document.querySelector("#metricCalculated"),
  "各場所統計供水量": document.querySelector("#metricSiteTotal"),
};

let availableDates = [];
let publicRecords = {};
let activeDateKey = "";
let dateListenersBound = false;

initialisePage();

function initialisePage() {
  clearDashboard();
  setMessage("請選擇年度供水日報表 Excel 檔案。", "info");
  if (fileInput) {
    fileInput.addEventListener("change", handleFileUpload);
  }
}

async function handleFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) {
    setMessage("請先選擇水量 Excel 檔案。", "info");
    return;
  }

  setMessage("Excel 檔案解析中，請稍候。", "info");
  clearDashboard();
  try {
    if (!window.XLSX) {
      throw new Error("Excel 解析套件尚未載入，請確認網路連線或重新整理頁面。");
    }
    if (!window.WaterStatusCalc) {
      throw new Error("水量計算模組尚未載入，請重新整理頁面後再試。");
    }

    const buffer = await file.arrayBuffer();
    const workbook = window.XLSX.read(buffer, { type: "array", cellFormula: true });
    const dataset = window.WaterStatusCalc.buildDataset(workbook, file.name);
    setupDateControls(dataset);
    renderSelectedDate(dataset.selectedKey);

    const firstDate = dataset.dates[0];
    const lastDate = dataset.dates.at(-1);
    uploadMeta.textContent = `已讀取 ${file.name}；可查詢日期 ${firstDate.label} 至 ${lastDate.label}，共 ${dataset.dates.length} 日。`;
  } catch (error) {
    publicRecords = {};
    availableDates = [];
    datePanel.hidden = true;
    uploadMeta.textContent = "檔案解析未完成，請確認是否為既有年度供水日報表格式。";
    setMessage(error.message || "Excel 檔案解析失敗。", "error");
  }
}

function setupDateControls(dataset) {
  publicRecords = dataset.records;
  availableDates = dataset.dates;
  const selected = publicRecords[dataset.selectedKey] || publicRecords[availableDates.at(-1)?.key];
  if (!selected) {
    throw new Error("未找到可供查詢之完整日期資料。");
  }

  bindDateListeners();
  setSelectOptions(yearSelect, uniqueSorted(availableDates.map((item) => item.year)), selected.date.year, (value) => `${value} 年`);
  refreshMonthOptions(selected.date.month);
  refreshDayOptions(selected.date.day);
  datePanel.hidden = false;
}

function bindDateListeners() {
  if (dateListenersBound) {
    return;
  }
  yearSelect.addEventListener("change", () => {
    refreshMonthOptions(Number(monthSelect.value));
    refreshDayOptions(Number(daySelect.value));
    renderDateFromControls();
  });
  monthSelect.addEventListener("change", () => {
    refreshDayOptions(Number(daySelect.value));
    renderDateFromControls();
  });
  daySelect.addEventListener("change", renderDateFromControls);
  dateListenersBound = true;
}

function refreshMonthOptions(preferredMonth) {
  const year = Number(yearSelect.value);
  const months = uniqueSorted(availableDates.filter((item) => item.year === year).map((item) => item.month));
  const selectedMonth = months.includes(preferredMonth) ? preferredMonth : months.at(-1);
  setSelectOptions(monthSelect, months, selectedMonth, (value) => `${pad2(value)} 月`);
}

function refreshDayOptions(preferredDay) {
  const year = Number(yearSelect.value);
  const month = Number(monthSelect.value);
  const days = uniqueSorted(availableDates.filter((item) => item.year === year && item.month === month).map((item) => item.day));
  const selectedDay = days.includes(preferredDay) ? preferredDay : days.at(-1);
  setSelectOptions(daySelect, days, selectedDay, (value) => `${pad2(value)} 日`);
}

function setSelectOptions(select, values, selectedValue, labelForValue) {
  select.replaceChildren();
  for (const value of values) {
    const option = document.createElement("option");
    option.value = String(value);
    option.textContent = labelForValue(value);
    select.appendChild(option);
  }
  select.value = String(selectedValue);
}

function renderDateFromControls() {
  renderSelectedDate(`${yearSelect.value}-${pad2(monthSelect.value)}-${pad2(daySelect.value)}`);
}

function renderSelectedDate(key) {
  const payload = publicRecords[key];
  if (!payload) {
    setMessage("查無該日期之完整水量資料。", "error");
    return;
  }
  activeDateKey = key;
  renderDashboard(payload);
  dateMeta.textContent = `目前顯示 ${payload.date.label}；本檔案共可查詢 ${availableDates.length} 日。`;
  setMessage(`已產出 ${payload.date.label} 水情資料。`, "info");
}

function clearDashboard() {
  activeDateKey = "";
  document.querySelector("#sourceLine").textContent = "請先選擇水量 Excel 檔案";
  for (const target of Object.values(metricTargets)) {
    target.textContent = "--";
  }
  for (const selector of ["#reservoirCards", "#outflowChart", "#rawWaterChart", "#crossSupportChart", "#yunlinSupportChart", "#minxiongSupportChart", "#controlTable", "#digestChart", "#auditList"]) {
    document.querySelector(selector).replaceChildren();
  }
}

function uniqueSorted(values) {
  return [...new Set(values)].sort((left, right) => left - right);
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function setMessage(text, type) {
  message.hidden = false;
  message.className = `message ${type}`;
  message.textContent = text;
}

function renderDashboard(payload) {
  document.querySelector("#sourceLine").textContent =
    `${payload.date.label}；原水 ${payload.source.rawSheet} 第 ${payload.source.rawRow} 列；供水 ${payload.source.supplySheet} 第 ${payload.source.supplyRow} 列`;

  for (const item of payload.digest) {
    if (metricTargets[item.label]) {
      metricTargets[item.label].textContent = item.display || "--";
    }
  }

  renderReservoir(payload.sections.reservoir);
  renderBars("#outflowChart", payload.sections.outflow);
  renderBars("#rawWaterChart", payload.sections.rawWater);
  renderBars("#crossSupportChart", payload.sections.crossSupport);
  renderBars("#yunlinSupportChart", payload.sections.yunlinSupport);
  renderBars("#minxiongSupportChart", payload.sections.minxiongSupport);
  renderControlTable(payload.sections.control);
  renderBars("#digestChart", payload.sections.digestChart);
  renderAudit(payload.audit);
}

function renderReservoir(rows) {
  const container = document.querySelector("#reservoirCards");
  container.replaceChildren();
  for (const row of rows) {
    const rate = toNumber(row.rate);
    const card = document.createElement("article");
    card.className = "reservoir-card";
    card.innerHTML = `
      <h3>${escapeHtml(row.name)}</h3>
      <dl>
        <dt>水位</dt>
        <dd>${row.level == null ? "--" : `${formatNumber(row.level, 2)} M`}</dd>
        <dt>有效蓄水量</dt>
        <dd>${formatNumber(row.storage, 2)} 萬m3</dd>
        <dt>蓄水率</dt>
        <dd>${formatNumber(rate, 2)}%</dd>
        <div class="rate-track"><i style="width:${clamp(rate, 0, 100)}%"></i></div>
      </dl>
    `;
    container.appendChild(card);
  }
}

function renderBars(selector, rows) {
  const container = document.querySelector(selector);
  container.replaceChildren();
  const maxValue = Math.max(1, ...rows.map((row) => Math.abs(toNumber(row.value))));
  for (const row of rows) {
    const value = toNumber(row.value);
    const width = clamp(Math.abs(value) / maxValue * 100, 2, 100);
    const item = document.createElement("div");
    item.className = "bar-row";
    item.innerHTML = `
      <div class="bar-label">${escapeHtml(row.name)}</div>
      <div class="bar-track"><i class="bar-fill" style="width:${width}%"></i></div>
      <div class="bar-value">${formatNumber(value, 0)}</div>
    `;
    container.appendChild(item);
  }
}

function renderControlTable(rows) {
  const body = document.querySelector("#controlTable");
  body.replaceChildren();
  for (const row of rows) {
    const difference = toNumber(row.difference);
    const signClass = difference > 0 ? "positive" : difference < 0 ? "negative" : "";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(row.name)}</td>
      <td>${formatNumber(row.supply, 0)}</td>
      <td>${formatNumber(row.control, 0)}</td>
      <td class="${signClass}">${formatNumber(row.difference, 0)}</td>
      <td class="${signClass}">${formatNumber(row.differenceRate, 2)}%</td>
    `;
    body.appendChild(tr);
  }
}

function renderAudit(rows) {
  const container = document.querySelector("#auditList");
  container.replaceChildren();
  for (const row of rows) {
    const statusClass = row.status === "通過" ? "pass" : row.status === "提醒" ? "warn" : "fail";
    const item = document.createElement("div");
    item.className = "audit-item";
    item.innerHTML = `
      <strong>${escapeHtml(row.item)}</strong>
      <span class="audit-status ${statusClass}">${escapeHtml(row.status)}</span>
      <span>${escapeHtml(row.detail)}</span>
    `;
    container.appendChild(item);
  }
}

function toNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function formatNumber(value, digits) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return "--";
  }
  return new Intl.NumberFormat("zh-TW", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  }).format(number);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
