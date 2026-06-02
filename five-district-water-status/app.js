const message = document.querySelector("#message");
const fileInput = document.querySelector("#fileInput");
const uploadMeta = document.querySelector("#uploadMeta");
const datePanel = document.querySelector("#datePanel");
const dateMeta = document.querySelector("#dateMeta");
const yearSelect = document.querySelector("#yearSelect");
const monthSelect = document.querySelector("#monthSelect");
const daySelect = document.querySelector("#daySelect");
const exportPanel = document.querySelector("#exportPanel");
const exportMeta = document.querySelector("#exportMeta");
const exportOrientation = document.querySelector("#exportOrientation");
const exportSize = document.querySelector("#exportSize");
const exportButton = document.querySelector("#exportButton");

const metricTargets = {
  "原水總取水量": document.querySelector("#metricRaw"),
  "總出水量": document.querySelector("#metricOutflow"),
  "支(受)援水量": document.querySelector("#metricSupport"),
  "計算後供水量": document.querySelector("#metricCalculated"),
  "各場所統計供水量": document.querySelector("#metricSiteTotal"),
};

const exportSizes = {
  long: { label: "長圖", portrait: [1440, 2560], landscape: [2560, 1440] },
  a4: { label: "A4", portrait: [1240, 1754], landscape: [1754, 1240] },
  a3: { label: "A3", portrait: [1754, 2480], landscape: [2480, 1754] },
  slide_16_9: { label: "簡報 16:9", portrait: [1080, 1920], landscape: [1920, 1080] },
  slide_4_3: { label: "簡報 4:3", portrait: [1200, 1600], landscape: [1600, 1200] },
  square: { label: "方形", portrait: [1600, 1600], landscape: [1600, 1600] },
  mobile: { label: "手機直式", portrait: [1080, 1920], landscape: [1920, 1080] },
};

const exportOrientationLabels = {
  portrait: "直式",
  landscape: "橫式",
};

let availableDates = [];
let publicRecords = {};
let activeDateKey = "";
let activePayload = null;
let dateListenersBound = false;

initialisePage();

function initialisePage() {
  clearDashboard();
  setMessage("請選擇年度供水日報表 Excel 檔案。", "info");
  if (fileInput) {
    fileInput.addEventListener("change", handleFileUpload);
  }
  if (exportButton) {
    exportButton.addEventListener("click", exportCurrentChart);
    exportOrientation.addEventListener("change", updateExportMeta);
    exportSize.addEventListener("change", updateExportMeta);
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
  activePayload = payload;
  renderDashboard(payload);
  exportPanel.hidden = false;
  updateExportMeta();
  dateMeta.textContent = `目前顯示 ${payload.date.label|；本檔案共可查詢 ${availableDates.length} 日。`;
  setMessage(`已產出 ${payload.date.label} 水情資料。`, "info");
}

function clearDashboard() {
  activeDateKey = "";
  activePayload = null;
  exportPanel.hidden = true;
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

function updateExportMeta() {
  if (!exportMeta) {
    return;
  }
  const config = selectedExportConfig();
  const dateText = activePayload ? `${activePayload.date.label}；` : "";
  exportMeta.textContent = `${dateText}${config.orientationLabel}／${config.sizeLabel}，輸出 ${config.width} x ${config.height} px PNG。`;
}

function selectedExportConfig() {
  const orientation = exportOrientation?.value || "portrait";
  const sizeKey = exportSize?.value || "long";
  const size = exportSizes[sizeKey] || exportSizes.long;
  const dimensions = size[orientation] || size.portrait;
  return {
    orientation,
    sizeKey,
    sizeLabel: size.label,
    orientationLabel: exportOrientationLabels[orientation] || "直式",
    width: dimensions[0],
    height: dimensions[1],
  };
}

async function exportCurrentChart() {
  if (!activePayload) {
    setMessage("請先上傳 Excel 並選擇日期後再匯出。", "error");
    return;
  }

  const config = selectedExportConfig();
  exportButton.disabled = true;
  try {
    const canvas = buildExportCanvas(activePayload, config);
    const blob = await canvasToPngBlob(canvas);
    const filename = `五區水情_${activePayload.date.key}_${config.sizeKey}_${config.orientation}.png`;
    downloadBlob(blob, filename);
    setMessage(`已匯出 ${activePayload.date.label} 圖表 PNG。`, "info");
  } catch (error) {
    setMessage(error.message || "圖表匯出失敗。", "error");
  } finally {
    exportButton.disabled = false;
  }
}

function canvasToPngBlob(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("瀏覽器未能產生 PNG 圖檔。"));
        return;
      }
      resolve(blob);
    }, "image/png");
  });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function buildExportCanvas(payload, config) {
  const baseWidth = config.orientation === "landscape" ? 1920 : 1240;
  const baseHeight = config.orientation === "landscape" ? 2600 : 3600;
  const reportCanvas = createCanvas(baseWidth, baseHeight);
  const reportContext = reportCanvas.getContext("2d");
  const reportHeight = Math.ceil(drawExportReport(reportContext, payload, baseWidth, config.orientation) + 40);
  const croppedHeight = Math.min(reportHeight, baseHeight);
  const croppedCanvas = createCanvas(baseWidth, croppedHeight);
  const croppedContext = croppedCanvas.getContext("2d");
  croppedContext.drawImage(reportCanvas, 0, 0);

  const targetCanvas = createCanvas(config.width, config.height);
  const targetContext = targetCanvas.getContext("2d");
  targetContext.fillStyle = "#f4f7f9";
  targetContext.fillRect(0, 0, config.width, config.height);
  targetContext.imageSmoothingEnabled = true;
  targetContext.imageSmoothingQuality = "high";

  const scale = Math.min(config.width / croppedCanvas.width, config.height / croppedCanvas.height);
  const drawWidth = croppedCanvas.width * scale;
  const drawHeight = croppedCanvas.height * scale;
  const offsetX = (config.width - drawWidth) / 2;
  const offsetY = (config.height - drawHeight) / 2;
  targetContext.drawImage(croppedCanvas, offsetX, offsetY, drawWidth, drawHeight);
  return targetCanvas;
}

function createCanvas(width, height) {
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  return canvas;
}

function drawExportReport(context, payload, width, orientation) {
  const margin = orientation === "landscape" ? 48 : 44;
  const gap = 22;
  const contentWidth = width - margin * 2;
  let y = margin;

  context.fillStyle = "#f4f7f9";
  context.fillRect(0, 0, width, context.canvas.height);
  y = drawReportHeader(context, payload, margin, y, contentWidth) + gap;
  y = drawMetricSummary(context, payload, margin, y, contentWidth, orientation === "landscape" ? 5 : 3) + gap;

  if (orientation === "landscape") {
    const columnWidth = (contentWidth - gap) / 2;
    let leftY = y;
    let rightY = y;
    leftY = drawReservoirExport(context, payload.sections.reservoir, margin, leftY, columnWidth, 240) + gap;
    leftY = drawBarSection(context, "彙整表2摘要", payload.sections.digestChart, margin, leftY, columnWidth, 285, "#2f6f88") + gap;
    leftY = drawSupportExport(context, payload, margin, leftY, columnWidth, 330) + gap;
    rightY = drawBarSection(context, "出水量", payload.sections.outflow, margin + columnWidth + gap, rightY, columnWidth, 350, "#3f73c5") + gap;
    rightY = drawBarSection(context, "原水量", payload.sections.rawWater, margin + columnWidth + gap, rightY, columnWidth, 350, "#4c7c59") + gap;
    rightY = drawControlTableExport(context, payload.sections.control, margin + columnWidth + gap, rightY, columnWidth, 440) + gap;
    y = Math.max(leftY, rightY);
  } else {
    const halfWidth = (contentWidth - gap) / 2;
    y = drawReservoirExport(context, payload.sections.reservoir, margin, y, contentWidth, 250) + gap;
    drawBarSection(context, "出水量", payload.sections.outflow, margin, y, halfWidth, 390, "#3f73c5");
    drawBarSection(context, "原水量", payload.sections.rawWater, margin + halfWidth + gap, y, halfWidth, 390, "#4c7c59");
    y += 390 + gap;
    drawSupportExport(context, payload, margin, y, halfWidth, 330);
    drawBarSection(context, "彙整表2摘要", payload.sections.digestChart, margin + halfWidth + gap, y, halfWidth, 330, "#2f6f88");
    y += 330 + gap;
    y = drawControlTableExport(context, payload.sections.control, margin, y, contentWidth, 500) + gap;
  }

  y = drawAuditSummary(context, payload.audit, margin, y, contentWidth, 110) + gap;
  setCanvasFont(context, 20, "700");
  context.fillStyle = "#607080";
  context.fillText("資料來源：使用者於瀏覽器端選擇之供水日報表 Excel；本圖由公開頁面本機產製，未上傳檔案。", margin, y + 30);
  return y + 60;
}

function drawReportHeader(context, payload, x, y, width) {
  drawCardBackground(context, x, y, width, 132, "#ffffff", "#cdd8e3");
  setCanvasFont(context, 42, "900");
  context.fillStyle = "#1f4e63";
  context.fillText("五區水情公開資訊圖表", x + 28, y + 54);
  setCanvasFont(context, 30, "800");
  context.fillStyle = "#2f6f88";
  context.fillText(payload.date.label, x + 28, y + 98);
  setCanvasFont(context, 20, "700");
  context.fillStyle = "#607080";
  drawSingleLine(context, `原水 ${payload.source.rawSheet} 第 ${payload.source.rawRow} 列；供水 ${payload.source.supplySheet} 第 ${payload.source.supplyRow} 列`, x + width * 0.45, y + 98, width * 0.52);
  return y + 132;
}

function drawMetricSummary(context, payload, x, y, width, columns) {
  const metrics = payload.digest.filter((item) => metricTargets[item.label]);
  const gap = 14;
  const rowHeight = 112;
  const cardWidth = (width - gap * (columns - 1)) / columns;
  for (const [index, metric] of metrics.entries()) {
    const column = index % columns;
    const row = Math.floor(index / columns);
    const cardX = x + column * (cardWidth + gap);
    const cardY = y + row * (rowHeight + gap);
    drawCardBackground(context, cardX, cardY, cardWidth, rowHeight, "#ffffff", "#cdd8e3");
    setCanvasFont(context, 18, "800");
    context.fillStyle = "#607080";
    drawSingleLine(context, metric.label, cardX + 18, cardY + 32, cardWidth - 36);
    setCanvasFont(context, 34, "900");
    context.fillStyle = "#1f4e63";
    context.fillText(metric.display || "--", cardX + 18, cardY + 78);
    setCanvasFont(context, 16, "800");
    context.fillStyle = "#607080";
    context.fillText("CMD", cardX + cardWidth - 58, cardY + 78);
  }
  const rows = Math.ceil(metrics.length / columns);
  return y + rows * rowHeight + (rows - 1) * gap;
}

function drawReservoirExport(context, rows, x, y, width, height) {
  drawSectionHeader(context, "水庫水情", "水位 / 有效蓄水量 / 蓄水率", x, y, width, height);
  const innerX = x + 18;
  const innerY = y + 74;
  const gap = 12;
  const cardWidth = (width - 36 - gap * 3) / 4;
  const cardHeight = height - 92;
  for (const [index, row] of rows.entries()) {
    const cardX = innerX + index * (cardWidth + gap);
    drawCardBackground(context, cardX, innerY, cardWidth, cardHeight, "#f8fbfd", "#d7e0e8");
    setCanvasFont(context, 18, "900");
    context.fillStyle = "#1f4e63";
    drawSingleLine(context, row.name, cardX + 14, innerY + 30, cardWidth - 28);
    setCanvasFont(context, 16, "800");
    context.fillStyle = "#607080";
    context.fillText("有效蓄水量", cardX + 14, innerY + 64);
    setCanvasFont(context, 24, "900");
    context.fillStyle = "#17212b";
    context.fillText(`${formatNumber(row.storage, 2)} 萬m3`, cardX + 14, innerY + 98);
    const rate = clamp(toNumber(row.rate), 0, 100);
    context.fillStyle = "#e4edf5";
    roundRect(context, cardX + 14, innerY + cardHeight - 34, cardWidth - 28, 14, 7);
    context.fill();
    context.fillStyle = "#4c7c59";
    roundRect(context, cardX + 14, innerY + cardHeight - 34, (cardWidth - 28) * rate / 100, 14, 7);
    context.fill();
    setCanvasFont(context, 16, "900");
    context.fillStyle = "#1f4e63";
    context.fillText(`${formatNumber(row.rate, 2)}%`, cardX + 14, innerY + cardHeight - 48);
  }
  return y + height;
}

function drawBarSection(context, title, rows, x, y, width, height, color) {
  drawSectionHeader(context, title, "CMD", x, y, width, height);
  const innerX = x + 20;
  const innerY = y + 74;
  const innerWidth = width - 40;
  const maxValue = Math.max(1, ...rows.map((row) => Math.abs(toNumber(row.value))));
  const rowHeight = Math.min(34, (height - 92) / rows.length);
  const labelWidth = Math.min(180, innerWidth * 0.36);
  const valueWidth = Math.min(120, innerWidth * 0.24);
  for (const [index, row] of rows.entries()) {
    const rowY = innerY + index * rowHeight;
    const value = toNumber(row.value);
    const barX = innerX + labelWidth + 12;
    const barWidth = innerWidth - labelWidth - valueWidth - 24;
    setCanvasFont(context, 17, "800");
    context.fillStyle = "#3f4e5d";
    drawSingleLine(context, row.name, innerX, rowY + 20, labelWidth);
    context.fillStyle = "#dce8fb";
    roundRect(context, barX, rowY + 8, barWidth, 14, 7);
    context.fill();
    context.fillStyle = color;
    roundRect(context, barX, rowY + 8, Math.max(3, barWidth * Math.abs(value) / maxValue), 14, 7);
    context.fill();
    setCanvasFont(context, 17, "900");
    context.fillStyle = "#17212b";
    context.textAlign = "right";
    context.fillText(formatNumber(value, 0), innerX + innerWidth, rowY + 20);
    context.textAlign = "left";
  }
  return y + height;
}

function drawSupportExport(context, payload, x, y, width, height) {
  drawSectionHeader(context, "支援水量摘要", "CMD", x, y, width, height);
  const groups = [
    { title: "跨區處支(受)援", rows: payload.sections.crossSupport },
    { title: "雲林支援嘉義", rows: payload.sections.yunlinSupport },
    { title: "民雄支援嘉義", rows: payload.sections.minxiongSupport },
  ];
  const innerX = x + 20;
  let rowY = y + 82;
  const maxValue = Math.max(1, ...groups.flatMap((group) => group.rows.map((row) => Math.abs(toNumber(row.value)))));
  for (const group of groups) {
    setCanvasFont(context, 18, "900");
    context.fillStyle = "#1f4e63";
    context.fillText(group.title, innerX, rowY);
    rowY += 12;
    for (const row of group.rows) {
      rowY += 30;
      const value = toNumber(row.value);
      setCanvasFont(context, 16, "800");
      context.fillStyle = "#3f4e5d";
      drawSingleLine(context, row.name, innerX, rowY, width * 0.34);
      const barX = innerX + width * 0.35;
      const barWidth = width * 0.42;
      context.fillStyle = "#dce8fb";
      roundRect(context, barX, rowY - 13, barWidth, 12, 6);
      context.fill();
      context.fillStyle = "#3f73c5";
      roundRect(context, barX, rowY - 13, Math.max(3, barWidth * Math.abs(value) / maxValue), 12, 6);
      context.fill();
      setCanvasFont(context, 16, "900");
      context.fillStyle = "#17212b";
      context.textAlign = "right";
      context.fillText(formatNumber(value, 0), x + width - 20, rowY);
      context.textAlign = "left";
    }
    rowY += 18;
  }
  return y + height;
}

function drawControlTableExport(context, rows, x, y, width, height) {
  drawSectionHeader(context, "各廠所供水量管控差異", "CMD / %", x, y, width, height);
  const innerX = x + 20;
  let rowY = y + 84;
  const columns = [
    { label: "廠所", width: 0.22, align: "left" },
    { label: "供水量", width: 0.2, align: "right" },
    { label: "管控值", width: 0.2, align: "right" },
    { label: "差異", width: 0.18, align: "right" },
    { label: "差異率", width: 0.2, align: "right" },
  ];
  setCanvasFont(context, 16, "900");
  context.fillStyle = "#1f4e63";
  drawTableRow(context, columns, [columns[0].label, columns[1].label, columns[2].label, columns[3].label, columns[4].label], innerX, rowY, width - 40, 20, "#1f4e63");
  rowY += 28;
  const rowHeight = (height - 120) / rows.length;
  for (const row of rows) {
    const color = toNumber(row.difference) > 0 ? "#c00000" : toNumber(row.difference) < 0 ? "#0b8f2a" : "#17212b";
    drawTableRow(
      context,
      columns,
      [row.name, formatNumber(row.supply, 0), formatNumber(row.control, 0), formatNumber(row.difference, 0), `${formatNumber(row.differenceRate, 2)}%`],
      innerX,
      rowY,
      width - 40,
      rowHeight,
      color,
    );
    rowY += rowHeight;
  }
  return y + height;
}

function drawAuditSummary(context, rows, x, y, width, height) {
  drawSectionHeader(context, "查核結果", "筆數", x, y, width, height);
  const counts = rows.reduce((result, row) => {
    result[row.status] = (result[row.status] || 0) + 1;
    return result;
  }, {});
  const items = [
    ["通過", counts["通過"] || 0, "#0b8f2a"],
    ["提醒", counts["提醒"] || 0, "#b7791f"],
    ["未通過", counts["未通過"] || 0, "#c00000"],
  ];
  const cardWidth = (width - 76) / 3;
  for (const [index, item] of items.entries()) {
    const cardX = x + 20 + index * (cardWidth + 18);
    drawCardBackground(context, cardX, y + 64, cardWidth, height - 82, "#ffffff", "#d7e0e8");
    setCanvasFont(context, 17, "900");
    context.fillStyle = item[2];
    context.fillText(item[0], cardX + 16, y + 94);
    setCanvasFont(context, 26, "900");
    context.fillText(`${item[1]}`, cardX + cardWidth - 46, y + 94);
  }
  return y + height;
}

function drawSectionHeader(context, title, unit, x, y, width, height) {
  drawCardBackground(context, x, y, width, height, "#ffffff", "#cdd8e3");
  context.fillStyle = "#2f6f88";
  roundRect(context, x, y, width, 52, 8);
  context.fill();
  setCanvasFont(context, 24, "900");
  context.fillStyle = "#ffffff";
  context.fillText(title, x + 20, y + 34);
  setCanvasFont(context, 17, "900");
  context.textAlign = "right";
  context.fillText(unit, x + width - 20, y + 34);
  context.textAlign = "left";
}

function drawTableRow(context, columns, values, x, y, width, height, color) {
  let columnX = x;
  for (const [index, column] of columns.entries()) {
    const columnWidth = width * column.width;
    context.fillStyle = index >= 3 ? color : "#17212b";
    context.textAlign = column.align;
    const textX = column.align === "right" ? columnX + columnWidth - 8 : columnX + 4;
    drawSingleLine(context, values[index], textX, y + height * 0.7, columnWidth - 12);
    columnX += columnWidth;
  }
  context.textAlign = "left";
}

function drawCardBackground(context, x, y, width, height, fill, stroke) {
  context.fillStyle = fill;
  roundRect(context, x, y, width, height, 8);
  context.fill();
  context.strokeStyle = stroke;
  context.lineWidth = 1.5;
  context.stroke();
}

function roundRect(context, x, y, width, height, radius) {
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

function setCanvasFont(context, size, weight) {
  context.font = `${weight} ${size}px "Microsoft JhengHei", "Noto Sans TC", Arial, sans-serif`;
}

function drawSingleLine(context, value, x, y, maxWidth) {
  const text = String(value ?? "");
  if (context.measureText(text).width <= maxWidth) {
    context.fillText(text, x, y);
    return;
  }
  let output = text;
  while (output.length > 1 && context.measureText(`${output}...`).width > maxWidth) {
    output = output.slice(0, -1);
  }
  context.fillText(`${output}...`, x, y);
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
