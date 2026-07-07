(function installSupplyControlSettings() {
  "use strict";

  if (window.__supplyControlSettingsNative) {
    return;
  }

  const panel = document.querySelector("#controlPanel");
  const meta = document.querySelector("#controlMeta");
  const form = document.querySelector("#controlForm");
  const grid = document.querySelector("#controlGrid");
  const resetButton = document.querySelector("#resetControlButton");
  let baseValues = {};
  let customValues = {};

  function plainNumber(value) {
    const number = Number(value);
    return Number.isFinite(number) ? String(Math.round(number)) : "";
  }

  function initControls() {
    const sourceKey = selectedDateKeys.at(-1) || availableDates.at(-1)?.key;
    const rows = waterDataset?.records[sourceKey]?.sections?.control || [];
    baseValues = {};
    customValues = {};
    for (const row of rows) {
      const name = row.name === PUZI_NAME ? PUZI_SPLIT_NAME : row.name;
      baseValues[name] = toNumber(row.control);
    }
    renderControls();
    if (panel) {
      panel.hidden = !Object.keys(baseValues).length;
    }
  }

  function renderControls() {
    if (!grid) {
      return;
    }
    const fields = Object.entries(baseValues).map(([name, value]) => {
      const label = document.createElement("label");
      label.className = "control-field";
      const title = document.createElement("span");
      title.textContent = name;
      const note = document.createElement("small");
      note.textContent = `原管控值 ${formatNumber(value, 0)} CMD`;
      const input = document.createElement("input");
      input.type = "number";
      input.min = "0";
      input.step = "1";
      input.inputMode = "numeric";
      input.dataset.site = name;
      input.value = plainNumber(customValues[name] ?? value);
      input.setAttribute("aria-label", `${name}管控值`);
      label.append(title, note, input);
      return label;
    });
    grid.replaceChildren(...fields);
    updateMeta();
  }

  function updateMeta() {
    if (!meta) {
      return;
    }
    const changed = Object.entries(baseValues).filter(([name, value]) => (customValues[name] ?? value) !== value).length;
    meta.textContent = changed ? `已手動調整 ${changed} 個廠所管控值；欄位留白可恢復原管控值。` : "目前採用原管控值；如需調整，請輸入新值後套用。";
  }

  function effectiveControl(name, originalControl) {
    const baseName = name === PUZI_NAME ? PUZI_SPLIT_NAME : name;
    return toNumber(customValues[baseName] ?? baseValues[baseName] ?? originalControl);
  }

  const originalBuildRowsForDate = buildRowsForDate;
  buildRowsForDate = function buildRowsForDateWithManualControls(dateKey) {
    const payload = waterDataset.records[dateKey];
    const jiakeSupply = jiakeRecords[dateKey]?.total ?? null;
    const rows = [];
    for (const row of payload.sections.control) {
      if (row.name === PUZI_NAME) {
        const puziWithoutJiake = toNumber(row.supply) - toNumber(jiakeSupply);
        rows.push(buildControlRow(PUZI_SPLIT_NAME, puziWithoutJiake, effectiveControl(PUZI_SPLIT_NAME, row.control)));
        rows.push({
          name: JIAKE_NAME,
          supply: jiakeSupply,
          control: null,
          difference: null,
          status: "none",
        });
        continue;
      }
      rows.push(buildControlRow(row.name, row.supply, effectiveControl(row.name, row.control)));
    }
    return rows.length ? rows : originalBuildRowsForDate(dateKey);
  };

  const originalRefreshAvailableDates = refreshAvailableDates;
  refreshAvailableDates = function refreshAvailableDatesWithControls() {
    originalRefreshAvailableDates();
    if (waterDataset && Object.keys(jiakeRecords).length && availableDates.length) {
      initControls();
    }
  };

  const originalResetAvailableDates = resetAvailableDates;
  resetAvailableDates = function resetAvailableDatesWithControls() {
    baseValues = {};
    customValues = {};
    if (grid) {
      grid.replaceChildren();
    }
    if (panel) {
      panel.hidden = true;
    }
    originalResetAvailableDates();
  };

  form?.addEventListener("submit", (event) => {
    event.preventDefault();
    const nextValues = {};
    for (const input of grid.querySelectorAll("input[data-site]")) {
      const site = input.dataset.site;
      const parsed = numberFromCell(input.value);
      if (parsed === null) {
        nextValues[site] = baseValues[site];
        input.value = plainNumber(baseValues[site]);
        continue;
      }
      if (parsed < 0) {
        setMessage("管控值不得小於 0，請修正後再套用。", "error");
        return;
      }
      nextValues[site] = parsed;
    }
    customValues = nextValues;
    renderDailyReport();
    updateMeta();
    setMessage("已套用管控值設定並重新計算日報表。", "info");
  });

  resetButton?.addEventListener("click", () => {
    customValues = {};
    renderControls();
    if (selectedDateKeys.length) {
      renderDailyReport();
    }
    setMessage("已恢復原管控值並重新計算日報表。", "info");
  });
})();
