(() => {
  const improvedExportSizes = {
    long: { label: "長圖", portrait: [2160, 3840], landscape: [3840, 2160] },
    a4: { label: "A4", portrait: [2480, 3508], landscape: [3508, 2480] },
    a3: { label: "A3", portrait: [3508, 4960], landscape: [4960, 3508] },
    slide_16_9: { label: "簡報 16:9", portrait: [2160, 3840], landscape: [3840, 2160] },
    slide_4_3: { label: "簡報 4:3", portrait: [2400, 3200], landscape: [3200, 2400] },
    square: { label: "方形", portrait: [2400, 2400], landscape: [2400, 2400] },
    mobile: { label: "手機直式", portrait: [2160, 3840], landscape: [3840, 2160] },
  };
  const orientationLabels = { portrait: "直式", landscape: "橫式" };
  let attempts = 0;

  function installExportQualityFix() {
    attempts += 1;
    if (typeof window.buildExportCanvas !== "function" || typeof window.drawSectionHeader !== "function") {
      if (attempts < 120) {
        window.setTimeout(installExportQualityFix, 100);
      }
      return;
    }

    window.selectedExportConfig = function selectedExportConfig() {
      const orientation = document.querySelector("#exportOrientation")?.value || "portrait";
      const sizeKey = document.querySelector("#exportSize")?.value || "long";
      const size = improvedExportSizes[sizeKey] || improvedExportSizes.long;
      const dimensions = size[orientation] || size.portrait;
      return {
        orientation,
        sizeKey,
        sizeLabel: size.label,
        orientationLabel: orientationLabels[orientation] || "直式",
        width: dimensions[0],
        height: dimensions[1],
      };
    };

    window.buildExportCanvas = function buildExportCanvas(payload, config) {
      const baseWidth = config.orientation === "landscape" ? 1920 : 1240;
      const baseHeight = config.orientation === "landscape" ? 2600 : 3600;
      const qualityScale = 2;
      const reportCanvas = createCanvas(baseWidth * qualityScale, baseHeight * qualityScale);
      const reportContext = reportCanvas.getContext("2d");
      reportContext.scale(qualityScale, qualityScale);
      const reportHeight = Math.ceil(drawExportReport(reportContext, payload, baseWidth, config.orientation) + 40);
      const croppedHeight = Math.min(reportHeight, baseHeight);
      const croppedCanvas = createCanvas(baseWidth * qualityScale, croppedHeight * qualityScale);
      const croppedContext = croppedCanvas.getContext("2d");
      croppedContext.drawImage(reportCanvas, 0, 0, baseWidth * qualityScale, croppedHeight * qualityScale, 0, 0, baseWidth * qualityScale, croppedHeight * qualityScale);

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
    };

    window.drawExportReport = function drawExportReport(context, payload, width, orientation) {
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
        leftY = drawReservoirExport(context, payload.sections.reservoir, margin, leftY, columnWidth, 300) + gap;
        leftY = drawBarSection(context, "彙整表2摘要", payload.sections.digestChart, margin, leftY, columnWidth, 285, "#2f6f88") + gap;
        leftY = drawSupportExport(context, payload, margin, leftY, columnWidth, 420) + gap;
        rightY = drawBarSection(context, "出水量", payload.sections.outflow, margin + columnWidth + gap, rightY, columnWidth, 350, "#3f73c5") + gap;
        rightY = drawBarSection(context, "原水量", payload.sections.rawWater, margin + columnWidth + gap, rightY, columnWidth, 350, "#4c7c59") + gap;
        rightY = drawControlTableExport(context, payload.sections.control, margin + columnWidth + gap, rightY, columnWidth, 440) + gap;
        y = Math.max(leftY, rightY);
      } else {
        const halfWidth = (contentWidth - gap) / 2;
        y = drawReservoirExport(context, payload.sections.reservoir, margin, y, contentWidth, 310) + gap;
        drawBarSection(context, "出水量", payload.sections.outflow, margin, y, halfWidth, 390, "#3f73c5");
        drawBarSection(context, "原水量", payload.sections.rawWater, margin + halfWidth + gap, y, halfWidth, 390, "#4c7c59");
        y += 390 + gap;
        drawSupportExport(context, payload, margin, y, halfWidth, 440);
        drawBarSection(context, "彙整表2摘要", payload.sections.digestChart, margin + halfWidth + gap, y, halfWidth, 330, "#2f6f88");
        y += 440 + gap;
        y = drawControlTableExport(context, payload.sections.control, margin, y, contentWidth, 500) + gap;
      }

      y = drawAuditSummary(context, payload.audit, margin, y, contentWidth, 110) + gap;
      setCanvasFont(context, 20, "700");
      context.fillStyle = "#607080";
      context.fillText("資料來源：使用者於瀏覽器端選擇之供水日報表 Excel；本圖由公開頁面本機產製，未上傳檔案。", margin, y + 30);
      return y + 60;
    };

    window.drawReservoirExport = function drawReservoirExport(context, rows, x, y, width, height) {
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
        setCanvasFont(context, 22, "900");
        context.fillStyle = "#17212b";
        drawSingleLine(context, `${formatNumber(row.storage, 2)} 萬m3`, cardX + 14, innerY + 96, cardWidth - 28);
        const rate = clamp(toNumber(row.rate), 0, 100);
        setCanvasFont(context, 16, "800");
        context.fillStyle = "#607080";
        context.fillText("蓄水率", cardX + 14, innerY + 126);
        setCanvasFont(context, 22, "900");
        context.fillStyle = "#1f4e63";
        context.fillText(`${formatNumber(row.rate, 2)}%`, cardX + 14, innerY + 156);
        context.fillStyle = "#e4edf5";
        roundRect(context, cardX + 14, innerY + cardHeight - 30, cardWidth - 28, 16, 8);
        context.fill();
        context.fillStyle = "#4c7c59";
        roundRect(context, cardX + 14, innerY + cardHeight - 30, (cardWidth - 28) * rate / 100, 16, 8);
        context.fill();
      }
      return y + height;
    };

    window.drawSupportExport = function drawSupportExport(context, payload, x, y, width, height) {
      drawSectionHeader(context, "支援水量摘要", "CMD", x, y, width, height);
      const groups = [
        { title: "跨區處支(受)援", rows: payload.sections.crossSupport },
        { title: "雲林支援嘉義", rows: payload.sections.yunlinSupport },
        { title: "民雄支援嘉義", rows: payload.sections.minxiongSupport },
      ];
      const innerX = x + 20;
      let rowY = y + 88;
      const maxValue = Math.max(1, ...groups.flatMap((group) => group.rows.map((row) => Math.abs(toNumber(row.value)))));
      for (const group of groups) {
        setCanvasFont(context, 18, "900");
        context.fillStyle = "#1f4e63";
        context.fillText(group.title, innerX, rowY);
        rowY += 18;
        for (const row of group.rows) {
          rowY += 36;
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
        rowY += 28;
      }
      return y + height;
    };

    window.__exportQualityFixApplied = true;
  }

  installExportQualityFix();
})();
