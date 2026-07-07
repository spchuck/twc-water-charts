(function attachWaterStatusCalc(global) {
  "use strict";

  const RAW_PREFIX = "原水";
  const RAW_ALT_PREFIX = "原水量";
  const SUPPLY_PREFIX = "供水量";
  const SUPPLY_ALT_PREFIX = "供水";
  const RAW_FIRST_DAY_ROW = 4;
  const SUPPLY_FIRST_DAY_ROW = 3;
  const CAPACITY_RENYITAN = 2459.19;
  const CAPACITY_LANTAN = 919.7;
  const CAPACITY_HUSHAN = 5041;
  const CONTROL_VALUES = [131406, 25697, 24574, 65216, 18753, 26573, 70822, 21154, 20993, 53500, 36737, 26950, 31227];
  const REVIEW_SMALL_PLANT_TOTAL_COLS = [30, 31, 32, 33, 34, 35, 36, 37, 38, 39];
  const REVIEW_YUNLIN_GROUNDWATER_COLS = [33, 34, 35, 36, 37, 38, 39];

  function columnLetter(columnNumber) {
    let column = "";
    let number = columnNumber;
    while (number > 0) {
      const remainder = (number - 1) % 26;
      column = String.fromCharCode(65 + remainder) + column;
      number = Math.floor((number - 1) / 26);
    }
    return column;
  }

  function asNumber(value) {
    if (value === null || value === undefined) {
      return null;
    }
    if (typeof value === "boolean") {
      return value ? 1 : 0;
    }
    if (typeof value === "number") {
      return Number.isFinite(value) ? value : null;
    }
    if (typeof value === "string") {
      let text = value.trim().replaceAll(",", "");
      if (!text || text.startsWith("=")) {
        return null;
      }
      const isPercent = text.endsWith("%");
      if (isPercent) {
        text = text.slice(0, -1);
      }
      const number = Number(text);
      if (!Number.isFinite(number)) {
        return null;
      }
      return isPercent ? number / 100 : number;
    }
    return null;
  }

  function cleanCellValue(cell) {
    if (!cell || cell.f) {
      return null;
    }
    if (cell.v !== undefined && cell.v !== null) {
      return cell.v;
    }
    return cell.w ?? null;
  }

  function cellValue(worksheet, row, column) {
    return cleanCellValue(worksheet[`${columnLetter(column)}${row}`]);
  }

  function cellNumber(worksheet, row, column) {
    return asNumber(cellValue(worksheet, row, column));
  }

  function maxWorksheetRow(worksheet) {
    const reference = worksheet["!ref"] || "A1:A1";
    const rowNumbers = reference.match(/\d+/g)?.map((value) => Number(value)) || [1];
    return Math.max(...rowNumbers);
  }

  function findSheet(workbook, candidates) {
    return candidates.find((name) => workbook.SheetNames.includes(name)) || null;
  }

  function rawSheetName(workbook, month) {
    const candidates = [`${RAW_PREFIX}${String(month).padStart(2, "0")}`, `${RAW_ALT_PREFIX}${String(month).padStart(2, "0")}`, `${RAW_PREFIX}${month}`, `${RAW_ALT_PREFIX}${month}`];
    const name = findSheet(workbook, candidates);
    if (!name) {
      throw new Error(`找不到 ${month} 月原水工作表。`);
    }
    return name;
  }

  function supplySheetName(workbook, month) {
    const candidates = [`${SUPPLY_PREFIX}${String(month).padStart(2, "0")}`, `${SUPPLY_ALT_PREFIX}${String(month).padStart(2, "0")}`, `${SUPPLY_PREFIX}${month}`, `${SUPPLY_ALT_PREFIX}${month}`];
    const name = findSheet(workbook, candidates);
    if (!name) {
      throw new Error(`找不到 ${month} 月供水工作表。`);
    }
    return name;
  }

  function preferredDayRow(firstDayRow, day) {
    return firstDayRow + day - 1;
  }

  function rowForDay(worksheet, firstDayRow, day) {
    const lastRow = Math.min(maxWorksheetRow(worksheet), firstDayRow + 35);
    for (let row = firstDayRow; row <= lastRow; row += 1) {
      if (cellNumber(worksheet, row, 3) === day) {
        return row;
      }
    }
    const row = preferredDayRow(firstDayRow, day);
    if (row > maxWorksheetRow(worksheet)) {
      throw new Error(`${worksheet["!name"] || "工作表"} 找不到 ${day} 日資料列。`);
    }
    return row;
  }

  function zeroIfBlank(value) {
    const number = asNumber(value);
    return number === null ? 0 : number;
  }

  function multiply(value, factor) {
    const number = asNumber(value);
    return number === null ? null : number * factor;
  }

  function safeDivide(numerator, denominator) {
    const number = asNumber(numerator);
    return number === null ? null : number / denominator;
  }

  function sumColumns(values, columns) {
    return columns.reduce((total, column) => total + zeroIfBlank(values[column]), 0);
  }

  function jijiRawWater(values) {
    return zeroIfBlank(values[22]) + zeroIfBlank(values[26]);
  }

  function hushanReservoirRawWater(values) {
    return zeroIfBlank(values[24]) + zeroIfBlank(values[27]);
  }

  function isMissing(value) {
    if (value === null || value === undefined) {
      return true;
    }
    if (typeof value === "string") {
      const text = value.trim();
      return text === "" || text.startsWith("=");
    }
    return false;
  }

  function missingColumns(values, columns) {
    return columns.filter((column) => isMissing(values[column])).map(columnLetter);
  }

  function missingGroup(values, columns) {
    return columns.every((column) => isMissing(values[column]));
  }

  function range(start, endExclusive) {
    return Array.from({ length: endExclusive - start }, (_, index) => start + index);
  }

  function rawValuesForRow(worksheet, reviewDate, row) {
    const values = {};
    for (let column = 1; column <= 44; column += 1) {
      values[column] = cellValue(worksheet, row, column);
    }
    values[1] = reviewDate.year;
    values[2] = reviewDate.month;
    values[3] = reviewDate.day;
    values[10] = safeDivide(values[7], CAPACITY_RENYITAN);
    values[11] = safeDivide(values[8], CAPACITY_LANTAN);
    values[12] = safeDivide(values[9], CAPACITY_HUSHAN);
    values[21] = sumColumns(values, [19, 20]);
    values[25] = sumColumns(values, [22, 23, 24]);
    values[28] = sumColumns(values, [26, 27]);
    values[29] = sumColumns(values, [17, 18, 21, 25, 28]);
    values[40] = sumColumns(values, REVIEW_SMALL_PLANT_TOTAL_COLS);
    return values;
  }

  function supplyValuesForRow(worksheet, reviewDate, row, rawValues) {
    const source = {};
    for (let column = 1; column <= 36; column += 1) {
      source[column] = cellValue(worksheet, row, column);
    }
    source[1] = reviewDate.year;
    source[2] = reviewDate.month;
    source[3] = reviewDate.day;
    source[9] = rawValues[40];
    source[11] = sumColumns(source, [4, 5, 6, 7, 8, 9, 10]);
    source[25] = sumColumns(source, range(12, 25));
    source[35] = sumColumns(source, [29, 32]);
    source[36] = safeDivide(source[35], 10000);

    const values = {};
    for (let column = 1; column < 30; column += 1) {
      values[column] = source[column];
    }
    values[30] = source[33];
    values[31] = source[34];
    values[32] = source[31];
    values[33] = source[32];
    return values;
  }

  function requiredChecks(rawValues, supplyValues) {
    const rows = [];
    const add = (name, ok, detail) => rows.push([name, ok ? "通過" : "未通過", detail]);
    const addWarning = (name, missing, okDetail) => {
      rows.push([name, missing.length ? "提醒" : "通過", missing.length ? `缺漏欄位：${missing.join(", ")}；未參與本日摘要及資訊圖表主要水量計算，請後續補填確認。` : okDetail]);
    };

    const rawLevelMissing = missingColumns(rawValues, [4, 5, 6]);
    add("原水：水庫水位", rawLevelMissing.length === 0, rawLevelMissing.length ? `缺漏欄位：${rawLevelMissing.join(", ")}` : "D:F 均有資料");

    const rawStorageMissing = missingColumns(rawValues, [7, 8, 9]);
    add("原水：有效蓄水量", rawStorageMissing.length === 0, rawStorageMissing.length ? `缺漏欄位：${rawStorageMissing.join(", ")}` : "G:I 均有資料");

    const rawMainMissing = missingColumns(rawValues, [15, 16, 19, 22, 28]);
    add("原水：主要原水取水量", rawMainMissing.length === 0, rawMainMissing.length ? `缺漏欄位：${rawMainMissing.join(", ")}` : "O:P、S、V、AB 均有資料");

    add("原水：地下水量", !missingGroup(rawValues, range(30, 40)) && !isMissing(rawValues[40]), "AD:AM 至少一欄有資料且 AN 合計可計算");
    addWarning("原水：濁度", missingColumns(rawValues, [41, 42, 43, 44]), "AO:AR 均有資料");

    const plantMissing = missingColumns(supplyValues, range(4, 11));
    add("供水：各淨水場出水量", plantMissing.length === 0, plantMissing.length ? `缺漏欄位：${plantMissing.join(", ")}` : "D:J 均有資料");

    const serviceMissing = missingColumns(supplyValues, range(12, 25));
    add("供水：各服務所供水量", serviceMissing.length === 0, serviceMissing.length ? `缺漏欄位：${serviceMissing.join(", ")}` : "L:X 均有資料");

    const supportMissing = missingColumns(supplyValues, [26, 27, 28, 29, 30, 31, 32, 33]);
    add("供水：支受援水量", supportMissing.length === 0, supportMissing.length ? `缺漏欄位：${supportMissing.join(", ")}` : "Z:AG 均有資料");

    return rows;
  }

  function allRequiredChecksPass(rows) {
    return rows.every((row) => row[1] !== "未通過");
  }

  function buildSummaryRows(rawValues, supplyValues) {
    const supportRow = [supplyValues[27], supplyValues[28], supplyValues[26], supplyValues[30], supplyValues[29], supplyValues[33], supplyValues[30], supplyValues[31]];
    const supportNet = zeroIfBlank(supportRow[0]) + zeroIfBlank(supportRow[1]) - zeroIfBlank(supportRow[2]);
    const yunlinMutual = zeroIfBlank(supportRow[4]) + zeroIfBlank(supportRow[5]);
    const rawIntake = [rawValues[15], rawValues[16], rawValues[19], rawValues[32], zeroIfBlank(rawValues[31]) + zeroIfBlank(rawValues[30]), 250000, jijiRawWater(rawValues), hushanReservoirRawWater(rawValues), supplyValues[32], sumColumns(rawValues, REVIEW_YUNLIN_GROUNDWATER_COLS), 195000];
    const plantSupply = [supplyValues[6], supplyValues[4], supplyValues[5]];
    const chiayiTotal = plantSupply.reduce((total, value) => total + zeroIfBlank(value), 0);
    const yunlinPlants = [supplyValues[8], supplyValues[7]];
    const yunlinTotal = yunlinPlants.reduce((total, value) => total + zeroIfBlank(value), 0);
    const otherTotal = zeroIfBlank(supplyValues[9]) + zeroIfBlank(supplyValues[10]);
    const grandOutflow = chiayiTotal + yunlinTotal + otherTotal;
    const serviceSupply = range(12, 25).map((column) => supplyValues[column]);
    const rawTotal = rawIntake.reduce((total, value, index) => ([0, 1, 2, 3, 4, 6, 7, 8, 9].includes(index) ? total + zeroIfBlank(value) : total), 0);
    const digest = [null, null, rawTotal, grandOutflow, supportNet, grandOutflow - supportNet, supplyValues[25]];
    return {
      reservoir: [rawValues[4], multiply(rawValues[10], 100), rawValues[5], multiply(rawValues[11], 100), rawValues[6], multiply(rawValues[12], 100)],
      support_row: supportRow,
      support_totals: [null, null, supportNet, null, yunlinMutual],
      raw_intake: rawIntake,
      plant_supply: [...plantSupply, chiayiTotal, ...yunlinPlants, yunlinTotal, otherTotal, null, grandOutflow],
      service_supply: serviceSupply,
      control_values: CONTROL_VALUES,
      digest,
    };
  }

  function formatNumber(value) {
    const number = asNumber(value);
    if (number === null) {
      return "";
    }
    return Number.isInteger(number) ? number.toLocaleString("zh-TW") : number.toLocaleString("zh-TW", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function reconciliationChecks(summaryRows, supplyValues) {
    const digest = summaryRows.digest;
    const totalOutflowFromPlants = zeroIfBlank(summaryRows.plant_supply[9]);
    const netSupport = zeroIfBlank(digest[4]);
    const calcSupply = zeroIfBlank(digest[5]);
    return [
      ["總量勾稽：出水量合計", totalOutflowFromPlants === zeroIfBlank(digest[3]) ? "通過" : "未通過", `彙整表 J22=${formatNumber(digest[3])}；各廠加總=${formatNumber(totalOutflowFromPlants)}`],
      ["總量勾稽：各場所統計供水量", zeroIfBlank(supplyValues[25]) === zeroIfBlank(digest[6]) ? "通過" : "未通過", `出水量 Y3=${formatNumber(supplyValues[25])}；彙整表2 G2=${formatNumber(digest[6])}`],
      ["總量勾稽：計算後供水量", zeroIfBlank(digest[3]) - netSupport === calcSupply ? "通過" : "未通過", `${formatNumber(digest[3])} - ${formatNumber(netSupport)} = ${formatNumber(calcSupply)}`],
    ];
  }

  function resolveSourceRows(workbook, reviewDate) {
    const rawName = rawSheetName(workbook, reviewDate.month);
    const supplyName = supplySheetName(workbook, reviewDate.month);
    const rawSheet = workbook.Sheets[rawName];
    const supplySheet = workbook.Sheets[supplyName];
    rawSheet["!name"] = rawName;
    supplySheet["!name"] = supplyName;
    return {
      rawSheetName: rawName,
      supplySheetName: supplyName,
      rawRow: rowForDay(rawSheet, RAW_FIRST_DAY_ROW, reviewDate.day),
      supplyRow: rowForDay(supplySheet, SUPPLY_FIRST_DAY_ROW, reviewDate.day),
    };
  }

  function dateLabel(reviewDate) {
    return `${reviewDate.year}年${String(reviewDate.month).padStart(2, "0")}月${String(reviewDate.day).padStart(2, "0")}日`;
  }

  function reviewDateKey(reviewDate) {
    return `${reviewDate.year}-${String(reviewDate.month).padStart(2, "0")}-${String(reviewDate.day).padStart(2, "0")}`;
  }

  function buildReviewData(workbook, fileName, reviewDate) {
    const rows = resolveSourceRows(workbook, reviewDate);
    const rawValues = rawValuesForRow(workbook.Sheets[rows.rawSheetName], reviewDate, rows.rawRow);
    const supplyValues = supplyValuesForRow(workbook.Sheets[rows.supplySheetName], reviewDate, rows.supplyRow, rawValues);
    const required = requiredChecks(rawValues, supplyValues);
    if (!allRequiredChecksPass(required)) {
      const failed = required.filter((row) => row[1] === "未通過").map((row) => `${row[0]}（${row[2]}）`);
      throw new Error(`${dateLabel(reviewDate)} 來源資料未完整：${failed.join("；")}`);
    }
    const summaryRows = buildSummaryRows(rawValues, supplyValues);
    summaryRows.digest[0] = reviewDate.month;
    summaryRows.digest[1] = reviewDate.day;
    const auditRows = [
      ["來源檔案", "通過", fileName],
      ["來源工作表：原水", "通過", `${rows.rawSheetName} 第 ${rows.rawRow} 列`],
      ["來源工作表：供水", "通過", `${rows.supplySheetName} 第 ${rows.supplyRow} 列`],
      ["選取日期", "通過", dateLabel(reviewDate)],
      ...required,
      ...reconciliationChecks(summaryRows, supplyValues),
    ];
    return { reviewDate, rows, fileName, rawValues, supplyValues, summaryRows, auditRows };
  }

  function metricRows(rows, headers) {
    return rows.map((row) => Object.fromEntries(headers.map((header, index) => [header, index < row.length ? row[index] : null])));
  }

  function infographicSections(data) {
    const raw = data.rawValues;
    const supply = data.supplyValues;
    const combinedStorage = zeroIfBlank(raw[7]) + zeroIfBlank(raw[8]);
    const combinedRate = combinedStorage / (CAPACITY_RENYITAN + CAPACITY_LANTAN) * 100;
    const digest = data.summaryRows.digest;
    const siteNames = ["嘉義所", "義竹所", "新港所", "民雄所", "竹崎所", "朴子所", "斗六所", "斗南所", "古坑所", "虎尾所", "西螺所", "台西所", "北港所"];
    const controlRows = siteNames.map((siteName, index) => {
      const supplyAmount = supply[12 + index];
      const controlValue = CONTROL_VALUES[index];
      const difference = zeroIfBlank(supplyAmount) - controlValue;
      return [siteName, supplyAmount, controlValue, difference, multiply(safeDivide(difference, controlValue), 100)];
    });
    return {
      reservoir: [["仁義潭水庫", raw[4], raw[7], multiply(raw[10], 100)], ["蘭潭水庫", raw[5], raw[8], multiply(raw[11], 100)], ["蘭潭-仁義潭水庫合計", null, combinedStorage, combinedRate], ["湖山水庫", raw[6], raw[9], multiply(raw[12], 100)]],
      outflow: [["公園淨水場", supply[6]], ["水上淨水場", supply[4]], ["蘭潭淨水場", supply[5]], ["湖山淨水場", supply[8]], ["林內淨水場", supply[7]], ["小型淨水場", supply[9]], ["台化", supply[10]], ["出水量合計", supply[11]]],
      raw_water: [["仁義潭", raw[15]], ["蘭潭", raw[16]], ["嘉南大圳", raw[19]], ["竹崎所地面水", raw[32]], ["新港所地下水", raw[30]], ["民雄所地下水", raw[31]], ["集集堰", jijiRawWater(raw)], ["湖山水庫", hushanReservoirRawWater(raw)], ["伏流水", supply[32]], ["雲林地下水", sumColumns(raw, REVIEW_YUNLIN_GROUNDWATER_COLS)], ["合計", digest[2]]],
      cross_support: [["五區支援六區", supply[27]], ["五區支援11區", supply[28]], ["六區支援五區", supply[26]]],
      yunlin_support: [["台一線", supply[29]], ["複線", supply[33]]],
      minxiong_support: [["華興橋", supply[30]], ["共同管溝", supply[31]]],
      digest_chart: [["原水總取水量", digest[2]], ["總出水量", digest[3]], ["支(受)援水量", digest[4]], ["計算後供水量", digest[5]], ["各場所統計供水量", digest[6]]],
      control: controlRows,
    };
  }

  function buildPayload(data) {
    const sections = infographicSections(data);
    const digestLabels = ["月", "日", "原水總取水量", "總出水量", "支(受)援水量", "計算後供水量", "各場所統計供水量"];
    return {
      ok: true,
      generatedAt: new Date().toISOString(),
      date: { key: reviewDateKey(data.reviewDate), year: data.reviewDate.year, month: data.reviewDate.month, day: data.reviewDate.day, label: dateLabel(data.reviewDate) },
      source: { rawSheet: data.rows.rawSheetName, rawRow: data.rows.rawRow, supplySheet: data.rows.supplySheetName, supplyRow: data.rows.supplyRow, file: data.fileName },
      digest: digestLabels.map((label, index) => ({ label, value: data.summaryRows.digest[index], display: formatNumber(data.summaryRows.digest[index]) })),
      sections: {
        reservoir: metricRows(sections.reservoir, ["name", "level", "storage", "rate"]),
        outflow: metricRows(sections.outflow, ["name", "value"]),
        rawWater: metricRows(sections.raw_water, ["name", "value"]),
        crossSupport: metricRows(sections.cross_support, ["name", "value"]),
        yunlinSupport: metricRows(sections.yunlin_support, ["name", "value"]),
        minxiongSupport: metricRows(sections.minxiong_support, ["name", "value"]),
        control: metricRows(sections.control, ["name", "supply", "control", "difference", "differenceRate"]),
        digestChart: metricRows(sections.digest_chart, ["name", "value"]),
      },
      audit: metricRows(data.auditRows, ["item", "status", "detail"]),
    };
  }

  function collectYears(workbook) {
    const years = new Set();
    for (const sheetName of workbook.SheetNames) {
      if (![RAW_PREFIX, RAW_ALT_PREFIX, SUPPLY_PREFIX, SUPPLY_ALT_PREFIX].some((prefix) => sheetName.startsWith(prefix))) {
        continue;
      }
      const worksheet = workbook.Sheets[sheetName];
      for (let row = 3; row <= Math.min(maxWorksheetRow(worksheet), 40); row += 1) {
        const year = cellNumber(worksheet, row, 1);
        if (Number.isInteger(year) && year >= 90 && year <= 150) {
          years.add(year);
        }
      }
    }
    return [...years].sort((left, right) => left - right);
  }

  function daysInMonth(rocYear, month) {
    return new Date(rocYear + 1911, month, 0).getDate();
  }

  function buildDataset(workbook, fileName) {
    const records = {};
    const availableDates = [];
    for (const year of collectYears(workbook)) {
      for (let month = 1; month <= 12; month += 1) {
        try {
          rawSheetName(workbook, month);
          supplySheetName(workbook, month);
        } catch {
          continue;
        }
        for (let day = 1; day <= daysInMonth(year, month); day += 1) {
          const reviewDate = { year, month, day };
          try {
            const payload = buildPayload(buildReviewData(workbook, fileName, reviewDate));
            records[payload.date.key] = payload;
            availableDates.push(payload.date);
          } catch {
            continue;
          }
        }
      }
    }
    if (!availableDates.length) {
      throw new Error("此 Excel 檔未找到原水及供水均完整之日期，請確認檔案格式與資料內容。");
    }
    return { selectedKey: availableDates.at(-1).key, records, dates: availableDates };
  }

  global.WaterStatusCalc = {
    buildDataset,
    buildReviewData,
    buildPayload,
  };
})(typeof window !== "undefined" ? window : globalThis);
