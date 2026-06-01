const message = document.querySelector("#message");
const exportPanel = document.querySelector("#exportPanel");
const exportMeta = document.querySelector("#exportMeta");
const downloadList = document.querySelector("#downloadList");
const previewPanel = document.querySelector("#previewPanel");
const previewMeta = document.querySelector("#previewMeta");
const exportPreview = document.querySelector("#exportPreview");

const metricTargets = {
  "原水總取水量": document.querySelector("#metricRaw"),
  "總出水量": document.querySelector("#metricOutflow"),
  "支(受)援水量": document.querySelector("#metricSupport"),
  "計算後供水量": document.querySelector("#metricCalculated"),
  "各場所統計供水量": document.querySelector("#metricSiteTotal"),
};

loadPublicData();

async function loadPublicData() {
  setMessage("資料載入中", "info");
  try {
    const response = await fetch("./data.json", { cache: "no-store" });
    const payload = await response.json();
    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || "公開資料載入失敗。");
    }
    renderDashboard(payload);
    setMessage(`已載入 ${payload.date.label}`, "info");
  } catch (error) {
    setMessage(error.message, "error");
  }
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
  renderDownloads(payload.downloads || []);
}

function renderDownloads(downloads) {
  downloadList.replaceChildren();
  if (!downloads.length) {
    exportPanel.hidden = true;
    previewPanel.hidden = true;
    return;
  }

  let preview = null;
  for (const file of downloads) {
    const link = document.createElement("a");
    link.className = `download-card ${file.kind === "workbook" ? "secondary" : ""}`;
    link.href = file.href;
    link.download = file.filename || "";
    link.textContent = file.label;
    downloadList.appendChild(link);
    if (!preview && file.kind === "infographic") {
      preview = file;
    }
  }

  exportMeta.textContent = "本頁下載檔均為外站同目錄靜態檔案，未連線至公司內部主機。";
  exportPanel.hidden = false;
  if (preview) {
    exportPreview.src = preview.href;
    previewMeta.textContent = preview.label;
    previewPanel.hidden = false;
  }
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
        <div><dt>水位</dt><dd>${formatNumber(row.level, 2)} M</dd></div>
        <div><dt>有效蓄水量</dt><dd>${formatNumber(row.storage, 2)} 萬m³</dd></div>
        <div><dt>蓄水率</dt><dd>${formatNumber(rate, 2)}%</dd></div>
      </dl>
      <div class="rate-track"><i style="width:${clamp(rate, 0, 100)}%"></i></div>
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
      <span>${escapeHtml(row.name)}</span>
      <div class="bar-track"><i class="bar-fill" style="width:${width}%"></i></div>
      <strong>${formatNumber(value, 0)}</strong>
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
