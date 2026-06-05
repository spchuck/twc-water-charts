(function () {
  const input = document.querySelector("#chartFileInput");
  const status = document.querySelector("#chartUploadStatus");
  const config = window.__waterHistoryConfig || {};

  if (!input || !status || !config.mode) {
    return;
  }

  input.addEventListener("change", handleUpload);

  async function handleUpload(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) {
      setStatus("未選擇檔案，仍顯示目前已發布資料。", "info");
      return;
    }

    try {
      if (!window.XLSX) {
        throw new Error("Excel 解析套件尚未載入，請確認網路連線後重新整理頁面。");
      }
      if (typeof window.renderWaterHistoryChart !== "function") {
        throw new Error("圖表模組尚未完成初始化，請重新整理頁面後再試。");
      }

      setStatus(`正在解析 ${file.name}，請稍候。`, "info");
      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(buffer, { type: "array", cellDates: true, cellFormula: true });
      const rows = workbookRows(workbook, config.sheetName);
      const payload = buildPayload(rows, config);
      window.renderWaterHistoryChart(payload, `使用者上傳：${file.name}`);
      setStatus(`已完成 ${file.name} 解析；序列 ${payload.series.length} 組，X 軸 ${payload.xLabels.length} 筆。`, "ok");
    } catch (error) {
      setStatus(error && error.message ? error.message : "Excel 檔案解析失敗。", "error");
    }
  }

  function workbookRows(workbook, sheetName) {
    const targetSheetName = resolveSheetName(workbook, sheetName);
    const worksheet = workbook.Sheets[targetSheetName];
    if (!worksheet) {
      throw new Error(`找不到工作表：${targetSheetName}`);
    }
    return window.XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: null, blankrows: false });
  }

  function resolveSheetName(workbook, sheetName) {
    if (!sheetName) {
      return workbook.SheetNames[0];
    }
    if (workbook.Sheets[sheetName]) {
      return sheetName;
    }
    const normalizedTarget = normalizeHeader(sheetName);
    return workbook.SheetNames.find((name) => normalizeHeader(name) === normalizedTarget) || sheetName;
  }

  function buildPayload(rows, chartConfig) {
    if (!rows || rows.length < 2) {
      throw new Error("工作表未包含可解析之歷線圖資料。");
    }

    const payload = {
      colors: chartConfig.colors,
      textColor: chartConfig.textColor,
      subColor: chartConfig.subColor,
      gridColor: chartConfig.gridColor,
      yAxisName: chartConfig.yAxisName,
    };

    if (chartConfig.mode === "monthly") {
      const result = monthlyAverageSeries(rows, chartConfig.divisor || 1);
      payload.xLabels = result.xLabels;
      payload.series = result.series;
      return payload;
    }

    const result = dailySeriesFromRows(rows, chartConfig.divisor || 1, chartConfig.averageLabel || "");
    payload.xLabels = result.xLabels;
    payload.series = result.series;
    return payload;
  }

  function dailySeriesFromRows(rows, divisor, averageLabel) {
    const headers = rows[0].map(normalizeHeader);
    const dataRows = rows.slice(1).filter((row) => row && row[0] != null && String(row[0]).trim() !== "");
    const xLabels = dataRows.map((row) => dayLabel(row[0]));
    const yearColumns = [];
    const extraColumns = [];

    headers.forEach((header, index) => {
      const year = yearFromHeader(header);
      if (year != null) {
        yearColumns.push({ year, index });
      } else if (header.includes("平均") || header.includes("運轉曲線")) {
        extraColumns.push({ header, index });
      }
    });

    if (!yearColumns.length && !extraColumns.length) {
      throw new Error("找不到年度、平均或運轉曲線欄位。");
    }

    yearColumns.sort((a, b) => b.year - a.year);
    const newestYear = yearColumns.reduce((maxYear, item) => Math.max(maxYear, item.year), 0);
    const series = [];

    yearColumns.forEach((item) => {
      series.push(buildLineSeries(`${item.year}年`, valuesForColumn(dataRows, item.index, divisor), 2, false, item.year === newestYear));
    });

    extraColumns.forEach((item) => {
      const label = averageLabel && item.header.includes("平均") ? averageLabel : item.header;
      series.push(buildLineSeries(label, valuesForColumn(dataRows, item.index, divisor), 3, true, false));
    });

    return { xLabels, series };
  }

  function monthlyAverageSeries(rows, divisor) {
    const headers = rows[0].map(normalizeHeader);
    const dataRows = rows.slice(1).filter((row) => row && row[0] != null && String(row[0]).trim() !== "");
    const dayLabels = dataRows.map((row) => dayLabel(row[0]));
    const monthNumbers = dayLabels.map((label) => Number(String(label).split("/")[0]));
    const series = [];

    headers.forEach((header, index) => {
      const year = yearFromHeader(header);
      if (year == null && !header.includes("平均") && !header.includes("運轉曲線")) {
        return;
      }

      const values = [];
      for (let month = 1; month <= 12; month += 1) {
        const sourceValues = [];
        dataRows.forEach((row, rowIndex) => {
          if (monthNumbers[rowIndex] !== month) {
            return;
          }
          const number = asNumber(row[index]);
          if (number != null) {
            sourceValues.push(number / divisor);
          }
        });
        values.push(sourceValues.length ? round(sum(sourceValues) / sourceValues.length, 3) : null);
      }

      if (year != null) {
        series.push(buildLineSeries(`${year}年(月平均)`, values, 2, false, year === 115));
      } else if (header.includes("平均")) {
        series.push(buildLineSeries("近5年平均(月平均)", values, 3, true, false));
      } else {
        series.push(buildLineSeries("運轉曲線(月平均)", values, 3, true, false));
      }
    });

    if (!series.length) {
      throw new Error("找不到可產製月平均曲線之欄位。");
    }

    return { xLabels: Array.from({ length: 12 }, (_, index) => `${index + 1}月`), series };
  }

  function valuesForColumn(rows, columnIndex, divisor) {
    return rows.map((row) => roundedOrNull(row[columnIndex], divisor, 3));
  }

  function buildLineSeries(name, values, width, dashed, symbol) {
    const lineStyle = { width };
    if (dashed) {
      lineStyle.type = "dashed";
    }
    return {
      name,
      type: "line",
      smooth: true,
      showSymbol: symbol,
      symbolSize: symbol ? 6 : 4,
      connectNulls: false,
      lineStyle,
      emphasis: { focus: "series" },
      data: values,
    };
  }

  function dayLabel(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return `${value.getMonth() + 1}/${value.getDate()}`;
    }
    const text = String(value || "").trim();
    const monthDay = text.match(/(\d{1,2})\s*[\/月-]\s*(\d{1,2})/);
    if (monthDay) {
      return `${Number(monthDay[1])}/${Number(monthDay[2])}`;
    }
    return text;
  }

  function normalizeHeader(value) {
    return String(value == null ? "" : value).trim().replace(/\s+/g, "");
  }

  function yearFromHeader(value) {
    const match = normalizeHeader(value).match(/^(\d{3})年$/);
    return match ? Number(match[1]) : null;
  }

  function roundedOrNull(value, divisor, digits) {
    const number = asNumber(value);
    return number == null ? null : round(number / divisor, digits);
  }

  function asNumber(value) {
    if (value == null || value === "") {
      return null;
    }
    if (typeof value === "number") {
      return Number.isFinite(value) ? value : null;
    }
    if (typeof value === "boolean") {
      return null;
    }
    const text = String(value).trim().replace(/,/g, "");
    if (!text) {
      return null;
    }
    const number = Number(text);
    return Number.isFinite(number) ? number : null;
  }

  function round(value, digits) {
    const factor = 10 ** digits;
    return Math.round((value + Number.EPSILON) * factor) / factor;
  }

  function sum(values) {
    return values.reduce((total, value) => total + value, 0);
  }

  function setStatus(text, type) {
    status.textContent = text;
    status.dataset.status = type;
  }
})();
