const message = document.querySelector("#message");
const sourceLine = document.querySelector("#sourceLine");
const uploadMeta = document.querySelector("#uploadMeta");
const waterFileInput = document.querySelector("#waterFileInput");
const jiakeFileInput = document.querySelector("#jiakeFileInput");
const datePanel = document.querySelector("#datePanel");
const dateMeta = document.querySelector("#dateMeta");
const dateForm = document.querySelector("#dateForm");
const dateSelect = document.querySelector("#dateSelect");
const chartDateSelect = document.querySelector("#chartDateSelect");
const resultPanel = document.querySelector("#resultPanel");
const tableMeta = document.querySelector("#tableMeta");
const chartMeta = document.querySelector("#chartMeta");
const supplyTableWrap = document.querySelector("#supplyTableWrap");
const differenceChart = document.querySelector("#differenceChart");
const metricDateCount = document.querySelector("#metricDateCount");
const metricLatestDate = document.querySelector("#metricLatestDate");

const MAX_SELECTED_DATES = 7;
const PUZI_NAME = "朴子所";
const PUZI_SPLIT_NAME = "朴子所不含嘉科";
const JIAKE_NAME = "嘉科";

let waterDataset = null;
let jiakeRecords = {};
let availableDates = [];
let activeRowsByDate = {};
let selectedDateKeys = [];

initialisePage();

function initialisePage() {
  setMessage("請選擇供水日報表及嘉科每日用水量檔案。", "info");
  waterFileInput?.addEventListener("change", handleWaterFileUpload);
  jiakeFileInput?.addEventListener("change", handleJiakeFileUpload);
  dateForm?.addEventListener("submit", handleDateSubmit);
  dateSelect?.addEventListener("change", handleDateSelectionChange);
  chartDateSelect?.addEventListener("change", renderSelectedChart);
}

async function handleWaterFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) {
    waterDataset = null;
    resetAvailableDates();
    return;
  }
  setMessage("供水日報表解析中，請稍候。", "info");
  try {
    ensureDependencies();
    const workbook = await readWorkbook(file);
    waterDataset = window.WaterStatusCalc.buildDataset(workbook, file.name);
    refreshAvailableDates();
  } catch (error) {
    waterDataset = null;
    resetAvailableDates();
    setMessage(error.message || "供水日報表解析失敗。", "error");
  }
}

async function handleJiakeFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) {
    jiakeRecords = {};
    resetAvailableDates();
    return;
  }
  setMessage("嘉科每日用水量解析中，請稍候。", "info");
  try {
    ensureDependencies();
    const workbook = await readWorkbook(file);
    jiakeRecords = parseJiakeWorkbook(workbook);
    refreshAvailableDates();
  } catch (error) {
    jiakeRecords = {};
    resetAvailableDates();
    setMessage(error.message || "嘉科每日用水量解析失敗。", "error");
  }
}

function ensureDependencies() {
  if (!window.XLSX) {
    throw new Error("Excel 解析套件尚未載入，請確認網路連線或重新整理頁面。");
  }
  if (!window.WaterStatusCalc) {
    throw new Error("水量計算模組尚未載入，請重新整理頁面後再試。");
  }
}

async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return window.XLSX.read(buffer, { type: "array", cellFormula: true, cellDates: false });
}

function refreshAvailableDates() {
  clearResults();
  if (!waterDataset || !Object.keys(jiakeRecords).length) {
    datePanel.hidden = true;
    uploadMeta.textContent = "請完成 2 個來源檔案上傳後再選擇日期。";
    sourceLine.textContent = "請先選擇供水日報表及嘉科每日用水量檔案";
    setMessage("尚待完成 2 個來源檔案解析。", "info");
    return;
  }

  availableDates = waterDataset.dates.filter((date) => waterDataset.records[date.key] && jiakeRecords[date.key]);
  if (!availableDates.length) {
    resetAvailableDates();
    setMessage("兩個來源檔案未找到可共同比對且完整之日期。", "error");
    return;
  }

  const defaultDates = availableDates.slice(-Math.min(3, MAX_SELECTED_DATES));
  selectedDateKeys = defaultDates.map((date) => date.key);
  setDateOptions(dateSelect, availableDates, selectedDateKeys);
  setChartDateOptions(defaultDates, defaultDates.at(-1)?.key || "");
  datePanel.hidden = false;
  uploadMeta.textContent = `可比對日期 ${availableDates[0].label} 至 ${availableDates.at(-1).label}，共 ${availableDates.length} 日。`;
  dateMeta.textContent = "請選擇欲呈現之日期，最多 7 日；差異圖日期限已選日期。";
  sourceLine.textContent = "供水日報表及嘉科每日用水量均已完成解析";
  renderDailyReport();
  setMessage("已完成來源資料解析，可選擇日期產生日報表。", "info");
}

function resetAvailableDates() {
  availableDates = [];
  activeRowsByDate = {};
  selectedDateKeys = [];
  datePanel.hidden = true;
  clearResults();
}

function setDateOptions(select, dates, selectedKeys) {
  select.replaceChildren();
  const selectedSet = new Set(selectedKeys);
  for (const date of dates) {
    const option = document.createElement("option");
    option.value = date.key;
    option.textContent = date.label;
    option.selected = selectedSet.has(date.key);
    select.appendChild(option);
  }
}

function setChartDateOptions(dates, selectedKey) {
  chartDateSelect.replaceChildren();
  for (const date of dates) {
    const option = document.createElement("option");
    option.value = date.key;
    option.textContent = date.label;
    chartDateSelect.appendChild(option);
  }
  chartDateSelect.value = selectedKey || dates.at(-1)?.key || "";
}

function handleDateSelectionChange() {
  const selected = selectedOptions(dateSelect).map((option) => option.value);
  if (selected.length > MAX_SELECTED_DATES) {
    const limited = selected.slice(-MAX_SELECTED_DATES);
    selectedDateKeys = limited;
    setDateOptions(dateSelect, availableDates, limited);
    setMessage("最多只能選擇 7 個日期，已保留最近選取之 7 日。", "error");
  } else {
    selectedDateKeys = selected;
  }
  syncChartOptions();
}

function handleDateSubmit(event) {
  event.preventDefault();
  selectedDateKeys = selectedOptions(dateSelect).map((option) => option.value);
  if (!selectedDateKeys.length) {
    clearResults();
    setMessage("請至少選擇 1 個日期。", "error");
    return;
  }
  if (selectedDateKeys.length > MAX_SELECTED_DATES) {
    selectedDateKeys = selectedDateKeys.slice(-MAX_SELECTED_DATES);
    setDateOptions(dateSelect, availableDates, selectedDateKeys);
  }
  syncChartOptions();
  renderDailyReport();
}

function selectedOptions(select) {
  return Array.from(select.selectedOptions || []);
}

function syncChartOptions() {
  const dates = selectedDateKeys.map((key) => dateByKey(key)).filter(Boolean);
  const current = dates.some((date) => date.key === chartDateSelect.value) ? chartDateSelect.value : dates.at(-1)?.key || "";
  setChartDateOptions(dates, current);
}

function renderDailyReport() {
  if (!selectedDateKeys.length) {
    return;
  }
  activeRowsByDate = Object.fromEntries(selectedDateKeys.map((key) => [key, buildRowsForDate(key)]));
  renderSupplyTable();
  renderSelectedChart();
  resultPanel.hidden = false;
  metricDateCount.textContent = formatNumber(selectedDateKeys.length, 0);
  metricLatestDate.textContent = dateByKey(selectedDateKeys.at(-1))?.label || "--";
  tableMeta.textContent = `已選 ${selectedDateKeys.length} 日；單位 CMD`;
  setMessage("日報表及差異圖已產生。", "info");
}

function buildRowsForDate(dateKey) {
  const payload = waterDataset.records[dateKey];
  const jiakeSupply = jiakeRecords[dateKey]?.total ?? null;
  const rows = [];
  for (const row of payload.sections.control) {
    if (row.name === PUZI_NAME) {
      const puziWithoutJiake = toNumber(row.supply) - toNumber(jiakeSupply);
      rows.push(buildControlRow(PUZI_SPLIT_NAME, puziWithoutJiake, row.control));
      rows.push({
        name: JIAKE_NAME,
        supply: jiakeSupply,
        control: null,
        difference: null,
        status: "none",
      });
      continue;
    }
    rows.push(buildControlRow(row.name, row.supply, row.control));
  }
  return rows;
}

function buildControlRow(name, supply, control) {
  const supplyNumber = toNumber(supply);
  const controlNumber = toNumber(control);
  const difference = supplyNumber - controlNumber;
  return {
    name,
    supply: supplyNumber,
    control: controlNumber,
    difference,
    status: difference > 0 ? "over" : difference < 0 ? "under" : "equal",
  };
}

function renderSupplyTable() {
  const dates = selectedDateKeys.map((key) => dateByKey(key)).filter(Boolean);
  const rowNames = activeRowsByDate[selectedDateKeys[0]]?.map((row) => row.name) || [];
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  headerRow.innerHTML = `<th>廠所別</th><th>管控值</th>${dates.map((date) => `<th>${escapeHtml(date.label)}</th>`).join("")}`;
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const name of rowNames) {
    const tr = document.createElement("tr");
    const firstRows = activeRowsByDate[selectedDateKeys[0]] || [];
    const firstRow = firstRows.find((row) => row.name === name);
    tr.appendChild(cell(name, "th"));
    tr.appendChild(cell(firstRow?.control === null ? "無管控值" : formatNumber(firstRow?.control, 0)));
    for (const dateKey of selectedDateKeys) {
      const row = activeRowsByDate[dateKey]?.find((item) => item.name === name);
      const td = document.createElement("td");
      td.innerHTML = renderSupplyValue(row);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  supplyTableWrap.replaceChildren(table);
}

function renderSupplyValue(row) {
  if (!row || !Number.isFinite(row.supply)) {
    return `<span class="supply-no-control">無資料</span>`;
  }
  if (row.control === null) {
    return `<span class="supply-value"><span class="supply-amount">${formatNumber(row.supply, 0)}</span><span class="supply-no-control">未設管控值</span></span>`;
  }
  const amountClass = row.status === "over" ? "over" : row.status === "under" ? "under" : "";
  return `<span class="supply-value"><span class="supply-amount ${amountClass}">${formatNumber(row.supply, 0)}</span><span class="supply-difference">(${formatSigned(row.difference)})</span></span>`;
}

function cell(text, elementName = "td") {
  const element = document.createElement(elementName);
  element.textContent = text;
  return element;
}

function renderSelectedChart() {
  const dateKey = chartDateSelect.value || selectedDateKeys.at(-1);
  const rows = activeRowsByDate[dateKey]?.filter((row) => row.control !== null) || [];
  const date = dateByKey(dateKey);
  chartMeta.textContent = date ? `${date.label}；差異量 CMD` : "CMD";
  if (!rows.length) {
    differenceChart.textContent = "請先選擇圖表日期。";
    return;
  }
  differenceChart.innerHTML = buildDifferenceChartSvg(rows);
}

function buildDifferenceChartSvg(rows) {
  const width = 1100;
  const height = 520;
  const margin = { top: 58, right: 38, bottom: 118, left: 58 };
  const plotWidth = width - margin.left - margin.right;
  const zeroY = 252;
  const maxBarHeight = 170;
  const maxAbs = Math.max(1, ...rows.map((row) => Math.abs(row.difference)));
  const step = plotWidth / rows.length;
  const barWidth = Math.min(46, step * 0.62);
  const bars = rows.map((row, index) => {
    const barHeight = Math.max(2, Math.abs(row.difference) / maxAbs * maxBarHeight);
    const x = margin.left + index * step + (step - barWidth) / 2;
    const y = row.difference >= 0 ? zeroY - barHeight : zeroY;
    const labelY = row.difference >= 0 ? y - 10 : y + barHeight + 22;
    const status = row.difference > 0 ? "over" : row.difference < 0 ? "under" : "equal";
    return `
      <rect class="chart-bar ${status}" x="${round(x)}" y="${round(y)}" width="${round(barWidth)}" height="${round(barHeight)}" rx="6"></rect>
      <text class="chart-value" x="${round(x + barWidth / 2)}" y="${round(labelY)}" text-anchor="middle">${escapeHtml(formatSigned(row.difference))}</text>
      <text class="chart-label" x="${round(x + barWidth / 2)}" y="426" text-anchor="end" transform="rotate(-36 ${round(x + barWidth / 2)} 426)">${escapeHtml(row.name)}</text>
    `;
  }).join("");
  return `
    <svg viewBox="0 0 ${width} ${height}" aria-label="當日管控差異圖">
      <line class="chart-grid" x1="${margin.left}" y1="${zeroY - maxBarHeight}" x2="${width - margin.right}" y2="${zeroY - maxBarHeight}"></line>
      <line class="chart-grid" x1="${margin.left}" y1="${zeroY + maxBarHeight}" x2="${width - margin.right}" y2="${zeroY + maxBarHeight}"></line>
      <line class="chart-axis" x1="${margin.left}" y1="${zeroY}" x2="${width - margin.right}" y2="${zeroY}"></line>
      <text class="chart-note" x="${margin.left}" y="${zeroY - maxBarHeight - 18}">高於管控值</text>
      <text class="chart-note" x="${margin.left}" y="${zeroY + maxBarHeight + 34}">低於管控值</text>
      <text class="chart-note" x="${width - margin.right}" y="${zeroY - 10}" text-anchor="end">0</text>
      ${bars}
    </svg>
  `;
}

function parseJiakeWorkbook(workbook) {
  const sheetName = workbook.SheetNames.find((name) => name.includes("115年")) || workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error("嘉科每日用水量檔案未找到可讀取之工作表。");
  }
  const header = findJiakeHeader(worksheet);
  const records = {};
  const range = window.XLSX.utils.decode_range(worksheet["!ref"] || "A1:D1");
  for (let row = header.row + 1; row <= range.e.r; row += 1) {
    const dateValue = readCell(worksheet, row, header.dateColumn);
    const key = jiakeDateKey(dateValue);
    if (!key) {
      continue;
    }
    const intakeOne = readCell(worksheet, row, header.intakeOneColumn);
    const intakeTwo = readCell(worksheet, row, header.intakeTwoColumn);
    const total = numberFromCell(readCell(worksheet, row, header.totalColumn));
    const first = numberFromCell(intakeOne);
    const second = numberFromCell(intakeTwo);
    const hasIntakeValue = first !== null || second !== null;
    const computedTotal = total !== null ? total : toNumber(first) + toNumber(second);
    if (!hasIntakeValue && computedTotal === 0) {
      continue;
    }
    records[key] = {
      key,
      total: computedTotal,
      intakeOne: first,
      intakeTwo: second,
      sourceRow: row + 1,
      sheetName,
    };
  }
  if (!Object.keys(records).length) {
    throw new Error("嘉科每日用水量檔案未找到完整之日期及合計水量。");
  }
  return records;
}

function findJiakeHeader(worksheet) {
  const range = window.XLSX.utils.decode_range(worksheet["!ref"] || "A1:D1");
  const scanLastRow = Math.min(range.e.r, range.s.r + 10);
  for (let row = range.s.r; row <= scanLastRow; row += 1) {
    const values = [];
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      values[column] = String(readCell(worksheet, row, column) ?? "").replace(/\s+/g, "");
    }
    const dateColumn = values.findIndex((value) => value.includes("日期"));
    const totalColumn = values.findIndex((value) => value.includes("合計"));
    if (dateColumn >= 0 && totalColumn >= 0) {
      const intakeColumns = values.map((value, index) => value.includes("500#") ? index : -1).filter((index) => index >= 0);
      return {
        row,
        dateColumn,
        intakeOneColumn: intakeColumns[0] ?? 1,
        intakeTwoColumn: intakeColumns[1] ?? 2,
        totalColumn,
      };
    }
  }
  return { row: 0, dateColumn: 0, intakeOneColumn: 1, intakeTwoColumn: 2, totalColumn: 3 };
}

function readCell(worksheet, row, column) {
  const cell = worksheet[window.XLSX.utils.encode_cell({ r: row, c: column })];
  if (!cell) {
    return null;
  }
  if (cell.v !== undefined && cell.v !== null) {
    return cell.v;
  }
  return cell.w ?? null;
}

function jiakeDateKey(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return gregorianDateKey(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = window.XLSX.SSF.parse_date_code(value);
    return parsed ? gregorianDateKey(parsed.y, parsed.m, parsed.d) : null;
  }
  const text = String(value ?? "").trim();
  const match = text.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/u);
  if (!match) {
    return null;
  }
  return gregorianDateKey(Number(match[1]), Number(match[2]), Number(match[3]));
}

function gregorianDateKey(year, month, day) {
  const rocYear = year - 1911;
  if (!Number.isInteger(rocYear) || rocYear < 90) {
    return null;
  }
  return `${rocYear}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function dateByKey(key) {
  return availableDates.find((date) => date.key === key) || waterDataset?.records[key]?.date || null;
}

function numberFromCell(value) {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  const text = String(value).trim().replaceAll(",", "");
  if (!text) {
    return null;
  }
  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}

function toNumber(value) {
  const number = numberFromCell(value);
  return number === null ? 0 : number;
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

function formatSigned(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return "--";
  }
  if (number > 0) {
    return `+${formatNumber(number, 0)}`;
  }
  if (number < 0) {
    return `-${formatNumber(Math.abs(number), 0)}`;
  }
  return "0";
}

function round(value) {
  return Math.round(value * 100) / 100;
}

function clearResults() {
  activeRowsByDate = {};
  resultPanel.hidden = true;
  supplyTableWrap.replaceChildren();
  differenceChart.replaceChildren();
  metricDateCount.textContent = "--";
  metricLatestDate.textContent = "--";
}

function setMessage(text, type) {
  message.hidden = false;
  message.className = `message ${type}`;
  message.textContent = text;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
