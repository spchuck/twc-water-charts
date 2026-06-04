const message = document.querySelector("#message");
const fileInput = document.querySelector("#fileInput");
const uploadMeta = document.querySelector("#uploadMeta");
const sourceLine = document.querySelector("#sourceLine");
const rangePanel = document.querySelector("#rangePanel");
const rangeForm = document.querySelector("#rangeForm");
const rangeMeta = document.querySelector("#rangeMeta");
const startDateSelect = document.querySelector("#startDateSelect");
const endDateSelect = document.querySelector("#endDateSelect");
const calculateButton = document.querySelector("#calculateButton");
const downloadCsvButton = document.querySelector("#downloadCsvButton");
const resultPanel = document.querySelector("#resultPanel");
const resultMeta = document.querySelector("#resultMeta");
const averageTables = document.querySelector("#averageTables");

const summaryTargets = {
  "納入天數": document.querySelector("#metricDayCount"),
  "原水總取水量": document.querySelector("#metricRawAverage"),
  "總出水量": document.querySelector("#metricOutflowAverage"),
  "支(受)援水量": document.querySelector("#metricSupportAverage"),
  "計算後供水量": document.querySelector("#metricCalculatedAverage"),
};

let availableDates = [];
let publicRecords = {};
let activeResult = null;

initialisePage();

function initialisePage() {
  clearResults();
  setMessage("請選擇年度供水日報表 Excel 檔案。", "info");
  fileInput?.addEventListener("change", handleFileUpload);
  rangeForm?.addEventListener("submit", handleRangeSubmit);
  downloadCsvButton?.addEventListener("click", downloadActiveCsv);
}

async function handleFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) {
    setMessage("請先選擇水量 Excel 檔案。", "info");
    return;
  }

  clearResults();
  setMessage("Excel 檔案解析中，請稍候。", "info");
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
    setupRangeControls(dataset, file.name);
  } catch (error) {
    availableDates = [];
    publicRecords = {};
    rangePanel.hidden = true;
    uploadMeta.textContent = "檔案解析未完成，請確認是否為既有年度供水日報表格式。";
    setMessage(error.message || "Excel 檔案解析失敗。", "error");
  }
}

function setupRangeControls(dataset, fileName) {
  publicRecords = dataset.records;
  availableDates = dataset.dates;
  if (!availableDates.length) {
    throw new Error("未找到可供查詢之完整日期資料。");
  }

  const firstDate = availableDates[0];
  const lastDate = availableDates.at(-1);
  setDateOptions(startDateSelect, availableDates, firstDate.key);
  setDateOptions(endDateSelect, availableDates, lastDate.key);
  sourceLine.textContent = `已讀取 ${fileName}`;
  uploadMeta.textContent = `可計算日期 ${firstDate.label} 至 ${lastDate.label}，共 ${availableDates.length} 日。`;
  rangeMeta.textContent = `請選擇起日及迄日；計算範圍包含起訖日期。`;
  rangePanel.hidden = false;
  calculateAndRenderAverage();
}

function setDateOptions(select, dates, selectedKey) {
  select.replaceChildren();
  for (const date of dates) {
    const option = document.createElement("option");
    option.value = date.key;
    option.textContent = date.label;
    select.appendChild(option);
  }
  select.value = selectedKey;
}

function handleRangeSubmit(event) {
  event.preventDefault();
  calculateAndRenderAverage();
}

function calculateAndRenderAverage() {
  const startKey = startDateSelect.value;
  const endKey = endDateSelect.value;
  if (!startKey || !endKey) {
    setMessage("請先選擇起日及迄日。", "error");
    return;
  }
  if (dateIndex(startKey) > dateIndex(endKey)) {
    clearResults();
    setMessage("起日不得晚於迄日，請重新選擇日期範圍。", "error");
    return;
  }

  const selectedDates = availableDates.filter((date) => date.key >= startKey && date.key <= endKey);
  const payloads = selectedDates.map((date) => publicRecords[date.key]).filter(Boolean);
  if (!payloads.length) {
    clearResults();
    setMessage("選取區間內無可計算之完整水量資料。", "error");
    return;
  }

  activeResult = buildAverageResult(payloads, startKey, endKey);
  renderAverageResult(activeResult);
  setMessage(`已完成 ${activeResult.startLabel} 至 ${activeResult.endLabel} 區間平均值計算。`, "info");
}

function buildAverageResult(payloads, startKey, endKey) {
  const groups = [
    averageRows("水庫有效蓄水量", "萬m3", payloads, (payload) => payload.sections.reservoir.map((row) => ({ name: row.name, value: row.storage }))),
    averageRows("出水量", "CMD", payloads, (payload) => payload.sections.outflow),
    averageRows("原水量", "CMD", payloads, (payload) => payload.sections.rawWater),
    averageRows("跨區處支(受)援", "CMD", payloads, (payload) => payload.sections.crossSupport),
    averageRows("雲林支援嘉義", "CMD", payloads, (payload) => payload.sections.yunlinSupport),
    averageRows("民雄支援嘉義", "CMD", payloads, (payload) => payload.sections.minxiongSupport),
    averageRows("彙整表2摘要", "CMD", payloads, (payload) => payload.sections.digestChart),
    averageRows("各廠所供水量", "CMD", payloads, (payload) => payload.sections.control.map((row) => ({ name: row.name, value: row.supply }))),
  ];

  return {
    startKey,
    endKey,
    startLabel: publicRecords[startKey]?.date.label || startKey,
    endLabel: publicRecords[endKey]?.date.label || endKey,
    dayCount: payloads.length,
    groups,
  };
}

function averageRows(groupName, unit, payloads, rowsForPayload) {
  const rowMap = new Map();
  for (const payload of payloads) {
    for (const row of rowsForPayload(payload)) {
      const value = toNumberOrNull(row.value);
      if (value === null) {
        continue;
      }
      const existing = rowMap.get(row.name) || { name: row.name, sum: 0, count: 0 };
      existing.sum += value;
      existing.count += 1;
      rowMap.set(row.name, existing);
    }
  }

  return {
    name: groupName,
    unit,
    rows: Array.from(rowMap.values()).map((row) => ({
      name: row.name,
      average: row.sum / row.count,
      count: row.count,
      unit,
    })),
  };
}

function renderAverageResult(result) {
  resultPanel.hidden = false;
  downloadCsvButton.disabled = false;
  summaryTargets["納入天數"].textContent = formatNumber(result.dayCount, 0);
  for (const label of ["原水總取水量", "總出水量", "支(受)援水量", "計算後供水量"]) {
    const row = findAverageRow(result, "彙整表2摘要", label);
    summaryTargets[label].textContent = row ? formatNumber(row.average, 2) : "--";
  }

  resultMeta.textContent = `${result.startLabel} 至 ${result.endLabel}；納入 ${result.dayCount} 日。`;
  averageTables.replaceChildren();
  for (const group of result.groups) {
    averageTables.appendChild(renderAverageGroup(group));
  }
}

function renderAverageGroup(group) {
  const section = document.createElement("section");
  section.className = "average-group";

  const title = document.createElement("h3");
  title.textContent = `${group.name}（${group.unit}）`;
  section.appendChild(title);

  if (!group.rows.length) {
    const empty = document.createElement("div");
    empty.className = "average-empty";
    empty.textContent = "本項目於選取區間內無可計算數值。";
    section.appendChild(empty);
    return section;
  }

  const wrap = document.createElement("div");
  wrap.className = "table-wrap";
  wrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>項目</th>
          <th>平均值</th>
          <th>單位</th>
          <th>納入筆數</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
  `;
  const body = wrap.querySelector("tbody");
  for (const row of group.rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(row.name)}</td>
      <td>${formatNumber(row.average, 2)}</td>
      <td>${escapeHtml(row.unit)}</td>
      <td>${formatNumber(row.count, 0)}</td>
    `;
    body.appendChild(tr);
  }
  section.appendChild(wrap);
  return section;
}

function findAverageRow(result, groupName, rowName) {
  return result.groups.find((group) => group.name === groupName)?.rows.find((row) => row.name === rowName) || null;
}

function downloadActiveCsv() {
  if (!activeResult) {
    setMessage("請先完成平均值計算後再下載 CSV。", "error");
    return;
  }

  const csv = buildCsv(activeResult);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `五區水情平均值_${activeResult.startKey}_${activeResult.endKey}.csv`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function buildCsv(result) {
  const rows = [["起日", result.startLabel], ["迄日", result.endLabel], ["納入天數", result.dayCount], [], ["分類", "項目", "平均值", "單位", "納入筆數"]];
  for (const group of result.groups) {
    for (const row of group.rows) {
      rows.push([group.name, row.name, roundTo(row.average, 2), row.unit, row.count]);
    }
  }
  return `\uFEFF${rows.map((row) => row.map(csvCell).join(",")).join("\r\n")}`;
}

function csvCell(value) {
  const text = String(value ?? "");
  return /[",\r\n]/u.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function clearResults() {
  activeResult = null;
  resultPanel.hidden = true;
  downloadCsvButton.disabled = true;
  averageTables.replaceChildren();
  for (const target of Object.values(summaryTargets)) {
    target.textContent = "--";
  }
}

function setMessage(text, type) {
  message.hidden = false;
  message.className = `message ${type}`;
  message.textContent = text;
}

function dateIndex(key) {
  return availableDates.findIndex((date) => date.key === key);
}

function toNumberOrNull(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function roundTo(value, digits) {
  const factor = 10 ** digits;
  return Math.round((Number(value) + Number.EPSILON) * factor) / factor;
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
