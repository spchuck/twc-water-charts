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
const exportMeta = document.querySelector("#exportMeta");
const exportForm = document.querySelector("#exportForm");
const exportOrientation = document.querySelector("#exportOrientation");
const exportSize = document.querySelector("#exportSize");
const exportButton = document.querySelector("#exportButton");

const MAX_SELECTED_DATES = 7;
const PUZI_NAME = "朴子所";
const PUZI_SPLIT_NAME = "朴子所不含嘉科";
const JIAKE_NAME = "嘉科";
const EXPORT_SIZES = {
  long: { label: "長圖", portrait: [2160, 3840], landscape: [3840, 2160] },
  a4: { label: "A4", portrait: [2480, 3508], landscape: [3508, 2480] },
  a3: { label: "A3", portrait: [3508, 4960], landscape: [4960, 3508] },
  slide_16_9: { label: "簡報 16:9", portrait: [2160, 3840], landscape: [3840, 2160] },
  slide_4_3: { label: "簡報 4:3", portrait: [2400, 3200], landscape: [3200, 2400] },
  square: { label: "方形", portrait: [2400, 2400], landscape: [2400, 2400] },
  mobile: { label: "手機直式", portrait: [2160, 3840], landscape: [3840, 2160] },
};
const EXPORT_ORIENTATION_LABELS = {
  portrait: "直式",
  landscape: "橫式",
};

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
  exportForm?.addEventListener("submit", (event) => event.preventDefault());
  exportOrientation?.addEventListener("change", updateExportMeta);
  exportSize?.addEventListener("change", updateExportMeta);
  exportButton?.addEventListener("click", exportSupplyInfographic);
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
  updateExportMeta();
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
  if (exportMeta) {
    exportMeta.textContent = "請先產生日報表後再匯出。";
  }
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

function exportConfig() {
  const orientation = exportOrientation?.value || "portrait";
  const sizeKey = exportSize?.value || "long";
  const size = EXPORT_SIZES[sizeKey] || EXPORT_SIZES.long;
  const dimensions = size[orientation] || size.portrait;
  return {
    orientation,
    orientationLabel: EXPORT_ORIENTATION_LABELS[orientation] || "直式",
    sizeKey,
    sizeLabel: size.label,
    width: dimensions[0],
    height: dimensions[1],
  };
}

function updateExportMeta() {
  if (!exportMeta) {
    return;
  }
  const config = exportConfig();
  const dateText = selectedDateKeys.length ? `已選 ${selectedDateKeys.length} 日；` : "";
  exportMeta.textContent = `${dateText}${config.orientationLabel}／${config.sizeLabel}，輸出 ${formatNumber(config.width, 0)} x ${formatNumber(config.height, 0)} px PNG。`;
}

async function exportSupplyInfographic() {
  if (!selectedDateKeys.length || !Object.keys(activeRowsByDate).length) {
    setMessage("請先上傳 2 個來源檔案並產生日報表後再匯出。", "error");
    return;
  }
  const config = exportConfig();
  exportButton.disabled = true;
  try {
    const canvas = buildSupplyExportCanvas(config);
    const link = document.createElement("a");
    link.download = `各廠所供水量日報表_${selectedDateKeys.at(-1)}_${config.sizeKey}_${config.orientation}.png`;
    link.href = canvas.toDataURL("image/png");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setMessage("已匯出各廠所供水量日報表資訊圖表。", "info");
  } catch (error) {
    setMessage(error.message || "資訊圖表匯出失敗。", "error");
  } finally {
    exportButton.disabled = false;
  }
}

function buildSupplyExportCanvas(config) {
  const canvas = document.createElement("canvas");
  canvas.width = config.width;
  canvas.height = config.height;
  const context = canvas.getContext("2d");
  const layout = exportLayout(config);
  paintExportBackground(context, config.width, config.height);
  let y = layout.margin;
  y = drawExportTitle(context, config, layout, y);
  y = drawExportSummary(context, layout, y);
  y = drawExportChart(context, layout, y);
  drawExportTable(context, layout, y);
  return canvas;
}

function exportLayout(config) {
  const compact = config.height <= 2400;
  const margin = Math.round(Math.min(config.width, config.height) * (compact ? 0.035 : 0.045));
  return {
    margin,
    width: config.width,
    height: config.height,
    contentWidth: config.width - margin * 2,
    titleSize: Math.max(54, Math.round(config.width * 0.038)),
    subtitleSize: Math.max(28, Math.round(config.width * 0.015)),
    sectionTitleSize: Math.max(32, Math.round(config.width * 0.018)),
    tableHeaderSize: Math.max(24, Math.round(config.width * 0.012)),
    tableSize: Math.max(23, Math.round(config.width * 0.011)),
    noteSize: Math.max(22, Math.round(config.width * 0.01)),
    compact,
  };
}

function paintExportBackground(context, width, height) {
  context.fillStyle = "#eef4f8";
  context.fillRect(0, 0, width, height);
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, Math.round(height * 0.17));
}

function drawExportTitle(context, config, layout, y) {
  const latestDate = dateByKey(selectedDateKeys.at(-1))?.label || "";
  const chartDate = dateByKey(chartDateSelect.value || selectedDateKeys.at(-1))?.label || latestDate;
  drawCanvasText(context, "各廠所供水量日報表", layout.margin, y + layout.titleSize, layout.titleSize, 900, "#123a46", "left");
  drawCanvasText(context, `選取日期 ${selectedDateKeys.length} 日；差異圖日期：${chartDate}`, layout.margin, y + layout.titleSize + layout.subtitleSize + 24, layout.subtitleSize, 700, "#526b7a", "left");
  drawCanvasText(context, `${config.orientationLabel} / ${config.sizeLabel}`, layout.width - layout.margin, y + layout.subtitleSize, layout.subtitleSize, 800, "#2f6f88", "right");
  return y + layout.titleSize + layout.subtitleSize + 60;
}

function drawExportSummary(context, layout, y) {
  const row = activeRowsByDate[selectedDateKeys.at(-1)] || [];
  const overCount = row.filter((item) => item.control !== null && item.difference > 0).length;
  const underCount = row.filter((item) => item.control !== null && item.difference < 0).length;
  const jiake = row.find((item) => item.name === JIAKE_NAME);
  const items = [
    ["選取日期", `${selectedDateKeys.length} 日`],
    ["高於管控值", `${overCount} 所`],
    ["低於管控值", `${underCount} 所`],
    ["嘉科供水量", jiake ? `${formatNumber(jiake.supply, 0)} CMD` : "--"],
  ];
  const gap = Math.round(layout.contentWidth * 0.018);
  const cardWidth = (layout.contentWidth - gap * (items.length - 1)) / items.length;
  const cardHeight = Math.round(layout.height * (layout.compact ? 0.07 : 0.06));
  items.forEach((item, index) => {
    const x = layout.margin + index * (cardWidth + gap);
    roundRectCanvas(context, x, y, cardWidth, cardHeight, 18, "#ffffff", "#d8e2ea", 2);
    drawCanvasText(context, item[0], x + 24, y + Math.round(cardHeight * 0.38), layout.noteSize, 700, "#526b7a", "left");
    drawCanvasText(context, item[1], x + 24, y + Math.round(cardHeight * 0.76), layout.sectionTitleSize, 900, "#123a46", "left");
  });
  return y + cardHeight + Math.round(layout.margin * 0.55);
}

function drawExportChart(context, layout, y) {
  const rows = activeRowsByDate[chartDateSelect.value || selectedDateKeys.at(-1)]?.filter((row) => row.control !== null) || [];
  const chartHeight = Math.round(layout.height * (layout.compact ? 0.29 : 0.25));
  roundRectCanvas(context, layout.margin, y, layout.contentWidth, chartHeight, 18, "#ffffff", "#d8e2ea", 2);
  drawCanvasText(context, "當日管控差異圖（CMD）", layout.margin + 28, y + layout.sectionTitleSize + 20, layout.sectionTitleSize, 900, "#123a46", "left");
  if (!rows.length) {
    drawCanvasText(context, "無可繪製資料", layout.margin + 28, y + chartHeight / 2, layout.sectionTitleSize, 700, "#526b7a", "left");
    return y + chartHeight + Math.round(layout.margin * 0.55);
  }
  const chartTop = y + layout.sectionTitleSize + 56;
  const chartBottom = y + chartHeight - Math.max(92, layout.sectionTitleSize * 2.1);
  const zeroY = Math.round((chartTop + chartBottom) / 2);
  const maxBarHeight = Math.max(80, Math.round((chartBottom - chartTop) * 0.45));
  const maxAbs = Math.max(1, ...rows.map((row) => Math.abs(row.difference)));
  const plotLeft = layout.margin + 50;
  const plotRight = layout.margin + layout.contentWidth - 36;
  const plotWidth = plotRight - plotLeft;
  const step = plotWidth / rows.length;
  const barWidth = Math.min(Math.round(step * 0.58), Math.max(36, Math.round(layout.width * 0.018)));
  context.strokeStyle = "#cdd8e2";
  context.lineWidth = 2;
  drawLine(context, plotLeft, zeroY - maxBarHeight, plotRight, zeroY - maxBarHeight);
  drawLine(context, plotLeft, zeroY + maxBarHeight, plotRight, zeroY + maxBarHeight);
  context.strokeStyle = "#526b7a";
  context.lineWidth = 4;
  drawLine(context, plotLeft, zeroY, plotRight, zeroY);
  rows.forEach((row, index) => {
    const barHeight = Math.max(4, Math.abs(row.difference) / maxAbs * maxBarHeight);
    const x = plotLeft + index * step + (step - barWidth) / 2;
    const barY = row.difference >= 0 ? zeroY - barHeight : zeroY;
    context.fillStyle = row.difference > 0 ? "#d92d20" : row.difference < 0 ? "#079455" : "#64748b";
    roundedRectPath(context, x, barY, barWidth, barHeight, 8);
    context.fill();
    const valueY = row.difference >= 0 ? barY - 14 : barY + barHeight + layout.noteSize + 8;
    drawCanvasText(context, formatSigned(row.difference), x + barWidth / 2, valueY, layout.noteSize, 900, "#123a46", "center");
    drawRotatedLabel(context, row.name, x + barWidth / 2, y + chartHeight - 30, layout.noteSize, "#123a46");
  });
  drawCanvasText(context, "0", plotRight, zeroY - 12, layout.noteSize, 800, "#526b7a", "right");
  return y + chartHeight + Math.round(layout.margin * 0.55);
}

function drawExportTable(context, layout, y) {
  const rows = activeRowsByDate[selectedDateKeys[0]] || [];
  const dates = selectedDateKeys.map((key) => dateByKey(key)).filter(Boolean);
  const availableHeight = layout.height - y - layout.margin;
  const rowHeight = Math.max(54, Math.floor(availableHeight / (rows.length + 2)));
  const headerHeight = Math.round(rowHeight * 1.08);
  const tableHeight = headerHeight + rows.length * rowHeight;
  roundRectCanvas(context, layout.margin, y, layout.contentWidth, tableHeight, 18, "#ffffff", "#d8e2ea", 2);
  const dateColumnWidth = Math.max(150, Math.floor((layout.contentWidth - 310) / Math.max(1, dates.length)));
  const nameWidth = Math.max(150, layout.contentWidth - 160 - dateColumnWidth * dates.length);
  const controlWidth = 160;
  const columns = [
    { label: "廠所別", x: layout.margin, width: nameWidth, align: "left" },
    { label: "管控值", x: layout.margin + nameWidth, width: controlWidth, align: "right" },
    ...dates.map((date, index) => ({
      label: date.label.replace("年", ".").replace("月", ".").replace("日", ""),
      x: layout.margin + nameWidth + controlWidth + index * dateColumnWidth,
      width: dateColumnWidth,
      align: "right",
    })),
  ];
  context.fillStyle = "#dfeaf2";
  context.fillRect(layout.margin, y, layout.contentWidth, headerHeight);
  columns.forEach((column) => drawTableText(context, column.label, column, y + headerHeight * 0.64, layout.tableHeaderSize, 900, "#123a46"));
  rows.forEach((row, rowIndex) => {
    const rowY = y + headerHeight + rowIndex * rowHeight;
    context.fillStyle = rowIndex % 2 === 0 ? "#ffffff" : "#f7fafc";
    context.fillRect(layout.margin, rowY, layout.contentWidth, rowHeight);
    drawTableText(context, row.name, columns[0], rowY + rowHeight * 0.62, layout.tableSize, 900, "#123a46");
    drawTableText(context, row.control === null ? "無管控值" : formatNumber(row.control, 0), columns[1], rowY + rowHeight * 0.62, layout.tableSize, 800, "#526b7a");
    selectedDateKeys.forEach((dateKey, index) => {
      const item = activeRowsByDate[dateKey]?.find((candidate) => candidate.name === row.name);
      const column = columns[index + 2];
      const valueColor = !item || item.control === null ? "#123a46" : item.difference > 0 ? "#b42318" : item.difference < 0 ? "#067647" : "#123a46";
      const weight = item?.difference > 0 ? 900 : 800;
      drawTableText(context, item ? formatNumber(item.supply, 0) : "--", column, rowY + rowHeight * 0.43, layout.tableSize, weight, valueColor);
      const diffText = item?.control === null ? "未設管控值" : item ? `(${formatSigned(item.difference)})` : "";
      drawTableText(context, diffText, column, rowY + rowHeight * 0.78, Math.max(18, layout.tableSize - 4), 700, "#526b7a");
    });
  });
  context.strokeStyle = "#d8e2ea";
  context.lineWidth = 2;
  columns.slice(1).forEach((column) => drawLine(context, column.x, y, column.x, y + tableHeight));
  for (let rowIndex = 0; rowIndex <= rows.length; rowIndex += 1) {
    const lineY = y + headerHeight + rowIndex * rowHeight;
    drawLine(context, layout.margin, lineY, layout.margin + layout.contentWidth, lineY);
  }
}

function drawTableText(context, text, column, y, size, weight, color) {
  const padding = 20;
  const x = column.align === "right" ? column.x + column.width - padding : column.x + padding;
  drawCanvasText(context, text, x, y, size, weight, color, column.align, column.width - padding * 2);
}

function drawCanvasText(context, text, x, y, size, weight, color, align, maxWidth) {
  context.save();
  context.fillStyle = color;
  context.font = `${weight} ${size}px "Microsoft JhengHei", "Noto Sans TC", Arial, sans-serif`;
  context.textAlign = align;
  context.textBaseline = "alphabetic";
  if (maxWidth) {
    context.fillText(String(text ?? ""), x, y, maxWidth);
  } else {
    context.fillText(String(text ?? ""), x, y);
  }
  context.restore();
}

function drawRotatedLabel(context, text, x, y, size, color) {
  context.save();
  context.translate(x, y);
  context.rotate(-Math.PI / 5.8);
  drawCanvasText(context, text, 0, 0, size, 800, color, "right", 180);
  context.restore();
}

function drawLine(context, x1, y1, x2, y2) {
  context.beginPath();
  context.moveTo(x1, y1);
  context.lineTo(x2, y2);
  context.stroke();
}

function roundRectCanvas(context, x, y, width, height, radius, fill, stroke, lineWidth) {
  context.save();
  roundedRectPath(context, x, y, width, height, radius);
  context.fillStyle = fill;
  context.fill();
  if (stroke) {
    context.strokeStyle = stroke;
    context.lineWidth = lineWidth || 1;
    context.stroke();
  }
  context.restore();
}

function roundedRectPath(context, x, y, width, height, radius) {
  const safeRadius = Math.min(radius, width / 2, height / 2);
  context.beginPath();
  context.moveTo(x + safeRadius, y);
  context.lineTo(x + width - safeRadius, y);
  context.quadraticCurveTo(x + width, y, x + width, y + safeRadius);
  context.lineTo(x + width, y + height - safeRadius);
  context.quadraticCurveTo(x + width, y + height, x + width - safeRadius, y + height);
  context.lineTo(x + safeRadius, y + height);
  context.quadraticCurveTo(x, y + height, x, y + height - safeRadius);
  context.lineTo(x, y + safeRadius);
  context.quadraticCurveTo(x, y, x + safeRadius, y);
  context.closePath();
}
