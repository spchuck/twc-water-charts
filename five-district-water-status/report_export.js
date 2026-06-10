(() => {
  if (window.__a4ReportExportBootstrap) {
    return;
  }
  window.__a4ReportExportBootstrap = true;

  if (typeof window.exportCurrentPdf === "function" && typeof window.exportCurrentExcel === "function") {
    return;
  }

  let latestPayload = null;
  let installed = false;

  function installA4ReportExport() {
    const form = document.querySelector("#exportForm");
    if (!form || typeof window.renderDashboard !== "function") {
      window.setTimeout(installA4ReportExport, 100);
      return;
    }
    if (installed) {
      return;
    }
    installed = true;
    installStyles();
    installButtons(form);
    wrapRenderDashboard();
    wrapExportMeta();
  }

  function installStyles() {
    if (document.querySelector("#a4ReportExportStyle")) {
      return;
    }
    const style = document.createElement("style");
    style.id = "a4ReportExportStyle";
    style.textContent = `
      .export-form {
        grid-template-columns: minmax(100px, 0.7fr) minmax(140px, 1fr) repeat(3, minmax(112px, 0.7fr));
        min-width: min(100%, 820px);
      }
      @media (max-width: 520px) {
        .export-form {
          grid-template-columns: 1fr;
        }
      }
    `;
    document.head.appendChild(style);
  }

  function installButtons(form) {
    const pngButton = document.querySelector("#exportButton");
    let pdfButton = document.querySelector("#exportPdfButton");
    if (!pdfButton) {
      pdfButton = document.createElement("button");
      pdfButton.id = "exportPdfButton";
      pdfButton.type = "button";
      pdfButton.textContent = "匯出 PDF";
      pngButton?.after(pdfButton);
    }
    if (pdfButton.dataset.a4ReportBound !== "true") {
      pdfButton.addEventListener("click", exportCurrentPdf);
      pdfButton.dataset.a4ReportBound = "true";
    }

    let excelButton = document.querySelector("#exportExcelButton");
    if (!excelButton) {
      excelButton = document.createElement("button");
      excelButton.id = "exportExcelButton";
      excelButton.type = "button";
      excelButton.textContent = "匯出 Excel";
      document.querySelector("#exportPdfButton")?.after(excelButton);
    }
    if (excelButton.dataset.a4ReportBound !== "true") {
      excelButton.addEventListener("click", exportCurrentExcel);
      excelButton.dataset.a4ReportBound = "true";
    }
    form.addEventListener("submit", (event) => event.preventDefault());
  }

  function wrapRenderDashboard() {
    if (window.renderDashboard.__a4ReportWrapped) {
      return;
    }
    const original = window.renderDashboard;
    window.renderDashboard = function renderDashboardWithA4Report(payload) {
      latestPayload = payload;
      window.__waterStatusActivePayload = payload;
      return original.apply(this, arguments);
    };
    window.renderDashboard.__a4ReportWrapped = true;
  }

  function wrapExportMeta() {
    if (typeof window.updateExportMeta !== "function" || window.updateExportMeta.__a4ReportWrapped) {
      return;
    }
    const original = window.updateExportMeta;
    window.updateExportMeta = function updateExportMetaWithA4Report() {
      const result = original.apply(this, arguments);
      const target = document.querySelector("#exportMeta");
      if (target && latestPayload) {
        target.textContent = `${target.textContent} PDF、Excel 為 A4 直式、新細明體。`;
      }
      return result;
    };
    window.updateExportMeta.__a4ReportWrapped = true;
  }

  function currentPayload() {
    return latestPayload || window.__waterStatusActivePayload || null;
  }

  function exportCurrentPdf() {
    const payload = currentPayload();
    if (!payload) {
      setPageMessage("請先上傳 Excel 並選擇日期後再匯出。", "error");
      return;
    }
    const button = document.querySelector("#exportPdfButton");
    button.disabled = true;
    try {
      const printWindow = window.open("", "_blank", "noopener,noreferrer");
      if (!printWindow) {
        throw new Error("瀏覽器封鎖彈出視窗，請允許本頁開啟列印視窗後再試。");
      }
      printWindow.document.open();
      printWindow.document.write(buildA4ReportHtml(payload));
      printWindow.document.close();
      printWindow.focus();
      window.setTimeout(() => printWindow.print(), 250);
      setPageMessage("已開啟 A4 直式 PDF 列印版，請於列印目的地選擇「另存為 PDF」。", "info");
    } catch (error) {
      setPageMessage(error.message || "PDF 匯出失敗。", "error");
    } finally {
      button.disabled = false;
    }
  }

  function exportCurrentExcel() {
    const payload = currentPayload();
    if (!payload) {
      setPageMessage("請先上傳 Excel 並選擇日期後再匯出。", "error");
      return;
    }
    const button = document.querySelector("#exportExcelButton");
    button.disabled = true;
    try {
      const html = buildA4ReportHtml(payload);
      const blob = new Blob([`\ufeff${html}`], { type: "application/vnd.ms-excel;charset=utf-8" });
      downloadBlob(blob, `五區水情_${payload.date.key}_A4直式.xls`);
      setPageMessage(`已匯出 ${payload.date.label} A4 直式 Excel 報表。`, "info");
    } catch (error) {
      setPageMessage(error.message || "Excel 匯出失敗。", "error");
    } finally {
      button.disabled = false;
    }
  }

  function buildA4ReportHtml(payload) {
    const title = `五區水情公開資訊圖表 ${payload.date.label}`;
    const summaryRows = payload.digest
      .filter((item) => !["月", "日"].includes(item.label))
      .map((item) => [item.label, item.display || formatNumber(item.value, 0), "CMD"]);
    const supportRows = [
      ...payload.sections.crossSupport.map((row) => ["跨區處支(受)援", row.name, formatNumber(row.value, 0)]),
      ...payload.sections.yunlinSupport.map((row) => ["雲林支援嘉義", row.name, formatNumber(row.value, 0)]),
      ...payload.sections.minxiongSupport.map((row) => ["民雄支援嘉義", row.name, formatNumber(row.value, 0)]),
    ];

    return `<!doctype html>
<html lang="zh-Hant-TW">
<head>
  <meta charset="utf-8" />
  <title>${escapeHtml(title)}</title>
  <style>
    @page { size: A4 portrait; margin: 12mm 10mm; }
    @page Section1 { size: 21cm 29.7cm; margin: 12mm 10mm; mso-page-orientation: portrait; }
    html, body {
      margin: 0;
      padding: 0;
      color: #000000;
      font-family: "PMingLiU", "新細明體", serif;
      font-size: 12pt;
      line-height: 1.35;
      background: #ffffff;
    }
    .page { page: Section1; width: 100%; }
    h1 { margin: 0 0 6pt; text-align: center; font-size: 20pt; font-weight: 700; letter-spacing: 0; }
    .meta { margin: 0 0 10pt; text-align: center; font-size: 11pt; }
    h2 {
      margin: 10pt 0 4pt;
      padding: 3pt 5pt;
      border: 1pt solid #000000;
      background: #d9eaf7;
      font-size: 14pt;
      font-weight: 700;
      page-break-after: avoid;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
      margin: 0 0 7pt;
      page-break-inside: avoid;
    }
    th, td {
      border: 1pt solid #000000;
      padding: 4pt 5pt;
      vertical-align: middle;
      word-break: break-word;
    }
    th { background: #eef5fb; font-weight: 700; text-align: center; }
    .num { text-align: right; mso-number-format: "#,##0"; }
    .positive { color: #c00000; font-weight: 700; }
    .negative { color: #008000; font-weight: 700; }
    .note { margin-top: 9pt; font-size: 10.5pt; }
  </style>
</head>
<body>
  <div class="page">
    <h1>五區水情公開資訊圖表</h1>
    <p class="meta">${escapeHtml(payload.date.label)}；原水 ${escapeHtml(payload.source.rawSheet)} 第 ${payload.source.rawRow} 列；供水 ${escapeHtml(payload.source.supplySheet)} 第 ${payload.source.supplyRow} 列</p>
    ${renderReportTable("彙整表2摘要", ["項目", "數值", "單位"], summaryRows, [52, 30, 18])}
    ${renderReportTable("水庫水情", ["水庫", "水位(M)", "有效蓄水量(萬m3)", "蓄水率(%)"], payload.sections.reservoir.map((row) => [row.name, formatNumber(row.level, 2), formatNumber(row.storage, 2), formatNumber(row.rate, 2)]), [34, 22, 24, 20])}
    ${renderReportTable("出水量", ["項目", "數值(CMD)"], payload.sections.outflow.map((row) => [row.name, formatNumber(row.value, 0)]), [58, 42])}
    ${renderReportTable("原水量", ["項目", "數值(CMD)"], payload.sections.rawWater.map((row) => [row.name, formatNumber(row.value, 0)]), [58, 42])}
    ${renderReportTable("支援水量摘要", ["類別", "項目", "數值(CMD)"], supportRows, [34, 40, 26])}
    ${renderControlReportTable(payload.sections.control)}
    ${renderReportTable("查核結果", ["項目", "狀態", "說明"], payload.audit.map((row) => [row.item, row.status, row.detail]), [32, 14, 54])}
    <p class="note">資料來源：使用者於瀏覽器端選擇之供水日報表 Excel；本報表由公開頁面本機產製，未上傳檔案。</p>
  </div>
</body>
</html>`;
  }

  function renderReportTable(title, headers, rows, widths) {
    return `<h2>${escapeHtml(title)}</h2>
<table>
  <colgroup>${widths.map((width) => `<col style="width:${width}%" />`).join("")}</colgroup>
  <thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>
  <tbody>${rows.map((row) => `<tr>${row.map((value, index) => `<td class="${index > 0 && isReportNumber(value) ? "num" : ""}">${escapeHtml(value)}</td>`).join("")}</tr>`).join("")}</tbody>
</table>`;
  }

  function renderControlReportTable(rows) {
    const body = rows.map((row) => {
      const difference = toNumber(row.difference);
      const signClass = difference > 0 ? "positive" : difference < 0 ? "negative" : "";
      return `<tr>
        <td>${escapeHtml(row.name)}</td>
        <td class="num">${formatNumber(row.supply, 0)}</td>
        <td class="num">${formatNumber(row.control, 0)}</td>
        <td class="num ${signClass}">${formatNumber(row.difference, 0)}</td>
        <td class="num ${signClass}">${formatNumber(row.differenceRate, 2)}%</td>
      </tr>`;
    }).join("");
    return `<h2>各廠所供水量管控差異</h2>
<table>
  <colgroup><col style="width:24%" /><col style="width:19%" /><col style="width:19%" /><col style="width:19%" /><col style="width:19%" /></colgroup>
  <thead><tr><th>廠所</th><th>供水量(CMD)</th><th>管控值(CMD)</th><th>差異(CMD)</th><th>差異率</th></tr></thead>
  <tbody>${body}</tbody>
</table>`;
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

  function setPageMessage(text, type) {
    if (typeof window.setMessage === "function") {
      window.setMessage(text, type);
      return;
    }
    const message = document.querySelector("#message");
    if (message) {
      message.hidden = false;
      message.className = `message ${type}`;
      message.textContent = text;
    }
  }

  function isReportNumber(value) {
    return /^-?\d{1,3}(,\d{3})*(\.\d+)?$|^-?\d+(\.\d+)?$/.test(String(value ?? "").trim());
  }

  function toNumber(value) {
    const number = Number(value);
    return Number.isFinite(number) ? number : 0;
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

  installA4ReportExport();
})();
