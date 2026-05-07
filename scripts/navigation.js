(function () {
  const clock = document.getElementById("bastionClock");
  const fileState = document.getElementById("contextFileState");

  function tick() {
    if (!clock) return;
    const now = new Date();
    clock.textContent = now.toLocaleTimeString("uk-UA", {
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit"
    });
  }

  tick();
  setInterval(tick, 1000);

  function activateView(view) {
    document.querySelectorAll("[data-view]").forEach((button) => {
      button.classList.toggle("is-active", button.dataset.view === view);
    });

    document.querySelectorAll("[data-view-panel]").forEach((panel) => {
      panel.classList.toggle("is-active", panel.dataset.viewPanel === view);
    });
  }

  document.querySelectorAll("[data-view]").forEach((button) => {
    button.addEventListener("click", () => activateView(button.dataset.view));
  });

  const excelInput = document.getElementById("excelFile");
  if (excelInput && fileState) {
    excelInput.addEventListener("change", (event) => {
      const file = event.target.files?.[0];
      const name = file ? file.name : "очікується";

      fileState.textContent = name;
      setText("dashFileState", name);

      if (file) activateView("analytics");
    });
  }

  function setText(id, value) {
    const node = document.getElementById(id);
    if (node) node.textContent = value;
  }

  function asNumber(value) {
    const n = Number(value || 0);
    return Number.isFinite(n) ? n : 0;
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("uk-UA").format(asNumber(value));
  }

  function getLowThreshold() {
    const thresholdInput = document.getElementById("lowStockThreshold");
    const parsed = Number(thresholdInput?.value || 10);
    return Number.isFinite(parsed) && parsed >= 0 ? parsed : 10;
  }


  function getReportItems() {
    if (typeof window.getCurrentFilteredItems === "function") {
      try { return window.getCurrentFilteredItems(); } catch (error) { return []; }
    }
    return Array.isArray(window.ArtAmmoState?.unitItems) ? window.ArtAmmoState.unitItems : [];
  }

  function getReportFilterLabel() {
    const parts = [];
    const unitFilter = document.getElementById("unitFilter");
    const longRangeOnly = document.getElementById("longRangeOnly");
    const searchFilter = document.getElementById("searchFilter");
    const stockFilter = document.getElementById("stockFilter");
    const sortMode = document.getElementById("sortMode");

    if (unitFilter && unitFilter.value && unitFilter.value !== "all") parts.push(`підрозділ: ${unitFilter.value}`);
    if (longRangeOnly?.checked) parts.push("далекобійні");
    if (searchFilter?.value?.trim()) parts.push(`пошук: ${searchFilter.value.trim()}`);
    if (stockFilter && stockFilter.value && stockFilter.value !== "all") parts.push(`залишки: ${stockFilter.value}`);
    if (sortMode && sortMode.value && sortMode.value !== "file") parts.push(`сортування: ${sortMode.value}`);

    return parts.length ? parts.join("; ") : "без фільтрів";
  }

  function downloadTextFile(filename, content, mime = "application/json") {
    const blob = new Blob([content], { type: `${mime};charset=utf-8` });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function buildDecisionPackageFallback() {
    const state = window.ArtAmmoState || {};
    const items = getReportItems();
    const grouped = typeof window.groupByUnit === "function" ? window.groupByUnit(items) : {};
    const now = new Date();

    return {
      app: "BASTION",
      version: "v0.39",
      generatedAt: now.toISOString(),
      sourceFileName: state.sourceFileName || "unknown",
      filters: getReportFilterLabel(),
      rows: items.length,
      units: Object.keys(grouped).length,
      reportPassport: state.reportPassport || null,
      commanderSummary: state.commanderSummary || [],
      actionLog: state.actionLog || [],
      actionStatuses: state.actionStatuses || {},
      qualityIssues: state.dataQualityIssues || state.qualityIssues || [],
      exportItems: items.map(item => ({
        unit: item.unit,
        projectile: item.projectile,
        charge: item.charge,
        rangeKm: item.rangeKm,
        longRange: item.longRange,
        balance: item.balance,
        received: item.received,
        spent: item.spent
      }))
    };
  }

  function runReportAction(type) {
    const items = getReportItems();
    if (!items.length) {
      alert("Спочатку завантаж Excel-файл і дочекайся аналізу.");
      return;
    }

    if (type === "pdf") {
      document.getElementById("exportPdfBtn")?.click();
      return;
    }

    if (type === "excel") {
      document.getElementById("exportExcelBtn")?.click();
      return;
    }

    const existingPackageButton = document.getElementById("exportDecisionPackageBtn") || document.getElementById("decisionPackageExportBtn");
    if (existingPackageButton) {
      existingPackageButton.click();
      return;
    }

    const payload = buildDecisionPackageFallback();
    const date = new Date().toISOString().slice(0, 10);
    downloadTextFile(`bastion_decision_package_${date}.json`, JSON.stringify(payload, null, 2));
  }

  function bindReportsCenter() {
    document.getElementById("reportsRunPdfBtn")?.addEventListener("click", () => runReportAction("pdf"));
    document.getElementById("reportsRunExcelBtn")?.addEventListener("click", () => runReportAction("excel"));
    document.getElementById("reportsRunPackageBtn")?.addEventListener("click", () => runReportAction("package"));
  }

  function updateImportCenter() {
    const state = window.ArtAmmoState || {};
    const workbook = state.workbook;
    const sheetNames = workbook?.SheetNames || [];
    const items = Array.isArray(state.unitItems) ? state.unitItems : [];
    const summaryName = window.ART_AMMO_SCHEMA?.SUMMARY_SHEET_NAME || "АГ3+АР";
    const hasSummary = sheetNames.some(name => String(name).trim().toLowerCase() === summaryName.toLowerCase());
    const unitSheets = sheetNames.filter(name => String(name).trim().toLowerCase() !== summaryName.toLowerCase());

    const fileName = state.sourceFileName || document.getElementById("dashFileState")?.textContent || "очікується";

    setText("importFileName", fileName || "очікується");
    setText("importSheetCount", sheetNames.length ? formatNumber(sheetNames.length) : "—");
    setText("importSchemaState", sheetNames.length ? (hasSummary ? "структура виявлена" : "немає зведеного") : "—");
    setText("importAnalysisState", items.length ? `${formatNumber(items.length)} ряд.` : "—");
    setText("schemaSummarySheetText", hasSummary ? `${summaryName} — знайдено` : `${summaryName} — не знайдено`);
    setText("schemaUnitSheetsText", unitSheets.length ? `${formatNumber(unitSheets.length)} арк. підрозділів` : "очікується файл");
    setText("schemaColumnMapText", "C — БК, D — дальність, E — отримання, F — витрата, G — залишок");

    toggleImportCard("schemaSummarySheetCard", sheetNames.length ? (hasSummary ? "ok" : "warn") : "idle");
    toggleImportCard("schemaUnitSheetsCard", unitSheets.length ? "ok" : "idle");
    toggleImportCard("schemaColumnMapCard", sheetNames.length ? "ok" : "idle");

    const preview = document.getElementById("importSheetPreview");
    if (preview) {
      if (!sheetNames.length) {
        preview.innerHTML = `<div class="alert-item muted">Після завантаження Excel тут з’явиться перелік аркушів.</div>`;
      } else {
        preview.innerHTML = sheetNames.map(name => {
          const isSummary = String(name).trim().toLowerCase() === summaryName.toLowerCase();
          const rows = (state.analysisItems || []).filter(item => item.sheetName === name).length;
          return `<div class="import-sheet-item ${isSummary ? "summary" : "unit"}">
            <span>${name}</span>
            <strong>${isSummary ? "зведений" : `${formatNumber(rows)} ряд.`}</strong>
          </div>`;
        }).join("");
      }
    }
  }

  function toggleImportCard(id, state) {
    const node = document.getElementById(id);
    if (!node) return;
    node.classList.remove("is-ok", "is-warn", "is-idle");
    node.classList.add(state === "ok" ? "is-ok" : state === "warn" ? "is-warn" : "is-idle");
  }

  function updateDashboardMetrics() {
    const state = window.ArtAmmoState || {};
    const items = Array.isArray(state.unitItems) ? state.unitItems : [];
    const grouped = state.groupedByUnit || {};
    const threshold = getLowThreshold();

    const totalBalance = items.reduce((sum, item) => sum + asNumber(item.balance), 0);
    const longRange = items
      .filter(item => item.longRange)
      .reduce((sum, item) => sum + asNumber(item.balance), 0);
    const critical = items.filter(item => {
      const balance = asNumber(item.balance);
      return balance <= threshold;
    }).length;

    const exchangeCount = Array.isArray(state.exchangeRecommendations)
      ? state.exchangeRecommendations.length
      : Array.isArray(state.actionLog)
        ? state.actionLog.length
        : 0;

    const qualityIssues = Array.isArray(state.dataQualityIssues)
      ? state.dataQualityIssues.length
      : Array.isArray(state.qualityIssues)
        ? state.qualityIssues.length
        : 0;

    const unitValues = Object.values(grouped);
    const topUnit = unitValues
      .slice()
      .sort((a, b) => asNumber(b.totalBalance) - asNumber(a.totalBalance))[0];

    const readinessPercent = items.length
      ? Math.max(0, Math.min(100, Math.round(100 - (critical / items.length) * 100)))
      : null;

    setText("dashTotalBalance", items.length ? formatNumber(totalBalance) : "—");
    setText("dashLongRange", items.length ? formatNumber(longRange) : "—");
    setText("dashCriticalCount", items.length ? formatNumber(critical) : "—");
    setText("dashUnitCount", Object.keys(grouped).length ? formatNumber(Object.keys(grouped).length) : "—");
    setText("dashRowsCount", items.length ? formatNumber(items.length) : "—");
    setText("dashExchangeCount", items.length ? formatNumber(exchangeCount) : "—");
    setText("dashQualityIssues", items.length ? formatNumber(qualityIssues) : "—");
    setText("dashTopUnit", topUnit ? `${topUnit.unit} / ${formatNumber(topUnit.totalBalance)}` : "—");
    setText("dashCriticalFocus", items.length ? `${formatNumber(critical)} позицій ≤ ${threshold}` : "—");
    setText("dashReadinessPercent", readinessPercent === null ? "—" : `${readinessPercent}%`);
    setText("dashMapMode", items.length ? "LIVE DATA" : "STANDBY");

    const ring = document.getElementById("dashReadinessRing");
    if (ring) ring.style.setProperty("--value", readinessPercent === null ? 0 : readinessPercent);

    if (state.reportPassport?.analysisAt) {
      setText("dashLastAnalysis", new Date(state.reportPassport.analysisAt).toLocaleString("uk-UA"));
    } else if (items.length) {
      setText("dashLastAnalysis", "виконано");
    } else {
      setText("dashLastAnalysis", "—");
    }

    let readiness = "очікує файл";
    if (items.length) {
      if (critical > 0) readiness = "потребує уваги";
      else readiness = "стабільно";
    }
    setText("dashReadinessState", readiness);
    updateContextAlerts({ items, grouped, threshold, critical, longRange, totalBalance, exchangeCount, qualityIssues, readiness });
    updateModePreviews({ items, grouped, threshold, critical, longRange, totalBalance, exchangeCount, qualityIssues });
    updateImportCenter();

  }


  function pluralRows(value) {
    return `${formatNumber(value)} ряд.`;
  }

  function setHTML(id, value) {
    const node = document.getElementById(id);
    if (node) node.innerHTML = value;
  }

  function topUnitsByRisk(grouped, threshold) {
    return Object.values(grouped)
      .map(unit => {
        const risky = (unit.items || []).filter(item => asNumber(item.balance) <= threshold).length;
        return { unit: unit.unit, risky, total: asNumber(unit.totalBalance) };
      })
      .filter(row => row.risky > 0)
      .sort((a, b) => b.risky - a.risky || b.total - a.total)
      .slice(0, 4);
  }

  function updateContextAlerts(data) {
    const { items, grouped, threshold, critical, longRange, totalBalance, exchangeCount, qualityIssues, readiness } = data;
    setText("contextAnalysisState", readiness || "очікує файл");

    if (!items.length) {
      setHTML("contextAlertList", `<div class="alert-item muted">Завантаж Excel-файл для формування попереджень.</div>`);
      return;
    }

    const alerts = [];
    if (qualityIssues > 0) alerts.push({ tone: "red", text: `Проблем якості даних: ${formatNumber(qualityIssues)}` });
    if (critical > 0) alerts.push({ tone: "amber", text: `Критичні/малі залишки: ${formatNumber(critical)} позицій ≤ ${threshold}` });
    if (exchangeCount > 0) alerts.push({ tone: "blue", text: `Рекомендацій обміну: ${formatNumber(exchangeCount)}` });
    if (longRange === 0 && totalBalance > 0) alerts.push({ tone: "red", text: "Далекобійні залишки не виявлені" });

    topUnitsByRisk(grouped, threshold).forEach(row => {
      alerts.push({ tone: "amber", text: `${row.unit}: ${formatNumber(row.risky)} ризикових позицій` });
    });

    if (!alerts.length) alerts.push({ tone: "green", text: "Критичних попереджень не виявлено" });

    setHTML("contextAlertList", alerts.slice(0, 6).map(alert => `
      <div class="alert-item ${alert.tone}">${alert.text}</div>
    `).join(""));
  }

  function uniqueCount(items, field) {
    return new Set(items.map(item => item[field]).filter(Boolean)).size;
  }



  function normalizeQualityIssue(raw, index) {
    if (!raw || typeof raw !== "object") {
      return { type: "unknown", severity: "warn", title: String(raw || "Проблема"), unit: "—", row: "—" };
    }

    const type = raw.type || raw.code || raw.kind || raw.category || "unknown";
    const severity = raw.severity || raw.level || (String(type).includes("negative") ? "error" : "warn");
    const title = raw.title || raw.message || raw.text || qualityTypeLabel(type);
    const unit = raw.unit || raw.sheetName || raw.sheet || raw.підрозділ || "—";
    const row = raw.row || raw.rowIndex || raw.line || "—";

    return { ...raw, type, severity, title, unit, row, index };
  }

  function qualityTypeLabel(type) {
    const value = String(type || "").toLowerCase();
    if (value.includes("negative")) return "Від’ємне значення";
    if (value.includes("duplicate")) return "Дубль позиції";
    if (value.includes("range")) return "Відсутня або нульова дальність";
    if (value.includes("formula") || value.includes("balance")) return "Порушення формули залишку";
    if (value.includes("summary") || value.includes("reconcile")) return "Розбіжність зі зведеним аркушем";
    return "Проблема якості";
  }

  function buildFallbackQualityIssues(items, threshold) {
    const issues = [];
    const duplicateMap = new Map();

    items.forEach(item => {
      const balance = asNumber(item.balance);
      const received = asNumber(item.received);
      const spent = asNumber(item.spent);
      const range = asNumber(item.rangeMeters);
      const key = `${item.unit}||${item.combination}`;

      duplicateMap.set(key, (duplicateMap.get(key) || 0) + 1);

      if (range <= 0) {
        issues.push({ type: "missing_range", severity: "warn", title: "Відсутня або нульова дальність", unit: item.unit, row: item.row, projectile: item.projectile, charge: item.charge });
      }
      if (balance < 0 || received < 0 || spent < 0) {
        issues.push({ type: "negative_value", severity: "error", title: "Від’ємне числове значення", unit: item.unit, row: item.row, projectile: item.projectile, charge: item.charge });
      }
      if ((received || spent || balance) && Math.abs((received - spent) - balance) > 0.0001) {
        issues.push({ type: "balance_formula", severity: "warn", title: "Отримання - витрата ≠ залишок", unit: item.unit, row: item.row, projectile: item.projectile, charge: item.charge });
      }
    });

    duplicateMap.forEach((count, key) => {
      if (count <= 1) return;
      const [unit, combination] = key.split("||");
      issues.push({ type: "duplicate", severity: "warn", title: `Дубль позиції: ${combination}`, unit, row: "—" });
    });

    return issues;
  }

  function getQualityIssues(items, threshold) {
    const state = window.ArtAmmoState || {};
    const existing = Array.isArray(state.dataQualityIssues)
      ? state.dataQualityIssues
      : Array.isArray(state.qualityIssues)
        ? state.qualityIssues
        : [];

    const base = existing.length ? existing : buildFallbackQualityIssues(items, threshold);
    return base.map(normalizeQualityIssue);
  }

  function renderQualityCenter(items, grouped, threshold) {
    const issues = getQualityIssues(items, threshold);
    const summaryItems = Array.isArray(window.ArtAmmoState?.summaryItems) ? window.ArtAmmoState.summaryItems : [];
    const unitTotal = items.reduce((sum, item) => sum + asNumber(item.balance), 0);
    const summaryTotal = summaryItems.reduce((sum, item) => sum + asNumber(item.balance), 0);
    const diff = unitTotal - summaryTotal;

    const byType = issues.reduce((acc, issue) => {
      const label = qualityTypeLabel(issue.type);
      acc[label] = (acc[label] || 0) + 1;
      return acc;
    }, {});

    const errorCount = issues.filter(issue => String(issue.severity).toLowerCase().includes("error")).length;
    const warnCount = Math.max(0, issues.length - errorCount);

    setText("qualityTotalBadge", issues.length ? `${formatNumber(issues.length)} проблем` : "стабільно");
    setText("qualityIssueCount", issues.length ? `${formatNumber(issues.length)} записів` : "немає проблем");
    setText("qualityReconcileBadge", summaryItems.length ? (diff === 0 ? "сходиться" : `Δ ${formatNumber(diff)}`) : "немає зведеного");

    setHTML("qualityBreakdownList", issues.length ? Object.entries(byType).map(([label, count]) => `
      <div class="quality-break-row">
        <span>${label}</span>
        <strong>${formatNumber(count)}</strong>
      </div>
    `).join("") + `
      <div class="quality-break-row quality-total">
        <span>Помилки / попередження</span>
        <strong>${formatNumber(errorCount)} / ${formatNumber(warnCount)}</strong>
      </div>
    ` : `<div class="alert-item green">Критичних проблем якості не виявлено.</div>`);

    setHTML("qualityReconcileList", items.length ? `
      <div class="quality-break-row"><span>Залишок по підрозділах</span><strong>${formatNumber(unitTotal)}</strong></div>
      <div class="quality-break-row"><span>Зведений аркуш</span><strong>${summaryItems.length ? formatNumber(summaryTotal) : "—"}</strong></div>
      <div class="quality-break-row ${diff === 0 ? "ok" : "warn"}"><span>Різниця</span><strong>${summaryItems.length ? formatNumber(diff) : "немає даних"}</strong></div>
    ` : `<div class="alert-item muted">Завантаж Excel-файл для контрольної звірки.</div>`);

    setHTML("qualityIssueTable", issues.length ? `
      <div class="quality-table-head">
        <span>Тип</span><span>Рівень</span><span>Підрозділ</span><span>Рядок</span><span>Опис</span>
      </div>
      ${issues.slice(0, 80).map(issue => `
        <div class="quality-table-row ${String(issue.severity).toLowerCase().includes("error") ? "is-error" : "is-warn"}">
          <span>${qualityTypeLabel(issue.type)}</span>
          <span>${issue.severity}</span>
          <span>${issue.unit}</span>
          <span>${issue.row}</span>
          <span>${issue.title}</span>
        </div>
      `).join("")}
      ${issues.length > 80 ? `<div class="alert-item muted">Показано перші 80 проблем із ${formatNumber(issues.length)}.</div>` : ""}
    ` : `<div class="alert-item green">Журнал проблем якості порожній.</div>`);
  }

  function updateModePreviews(data) {
    const { items, grouped, threshold, critical, exchangeCount, qualityIssues } = data;
    const longRangeActions = Array.isArray(window.ArtAmmoState?.actionLog)
      ? window.ArtAmmoState.actionLog.filter(action => action.longRange).length
      : 0;

    setText("exchangeRecsPreview", items.length ? formatNumber(exchangeCount) : "—");
    setText("exchangeImpactPreview", items.length ? `${formatNumber(critical)} ризиків` : "—");
    setText("exchangeActionsPreview", items.length ? formatNumber((window.ArtAmmoState?.actionLog || []).length) : "—");
    setHTML("exchangeLiveList", items.length ? `
      <div class="mode-live-row"><span>Далекобійні дії</span><strong>${formatNumber(longRangeActions)}</strong></div>
      <div class="mode-live-row"><span>Поріг малого залишку</span><strong>≤ ${threshold}</strong></div>
    ` : `<div class="alert-item muted">Після аналізу тут з’являться рекомендації обміну.</div>`);

    setText("dictProjectilesPreview", items.length ? formatNumber(uniqueCount(items, "projectile")) : "—");
    setText("dictChargesPreview", items.length ? formatNumber(uniqueCount(items, "charge")) : "—");
    setText("dictCombinationsPreview", items.length ? formatNumber(uniqueCount(items, "combination")) : "—");
    setHTML("dictLiveList", items.length ? `
      <div class="mode-live-row"><span>Підрозділів у файлі</span><strong>${formatNumber(Object.keys(grouped).length)}</strong></div>
      <div class="mode-live-row"><span>Рядків для нормалізації</span><strong>${pluralRows(items.length)}</strong></div>
    ` : `<div class="alert-item muted">Довідники сформуються автоматично після імпорту Excel.</div>`);

    setText("qualityErrorsPreview", items.length ? formatNumber(qualityIssues) : "—");
    setText("qualityWarningsPreview", items.length ? formatNumber(critical) : "—");
    setText("qualityCheckPreview", items.length ? (qualityIssues ? "потребує уваги" : "стабільно") : "—");
    setHTML("qualityLiveList", items.length ? `
      <div class="mode-live-row"><span>Проблем якості</span><strong>${formatNumber(qualityIssues)}</strong></div>
      <div class="mode-live-row"><span>Малі/нульові залишки</span><strong>${formatNumber(critical)}</strong></div>
      <div class="mode-live-row"><span>Поріг контролю</span><strong>≤ ${threshold}</strong></div>
    ` : `<div class="alert-item muted">Після аналізу система покаже помилки й попередження файлу.</div>`);

    renderQualityCenter(items, grouped, threshold);

    setText("reportsPackagePreview", items.length ? "можна формувати" : "—");
    setText("reportsRowsCount", items.length ? pluralRows(items.length) : "—");
    setText("reportsUnitsCount", items.length ? formatNumber(Object.keys(grouped).length) : "—");
    setText("reportsFiltersState", items.length ? getReportFilterLabel() : "—");
    setHTML("reportsLiveList", items.length ? `
      <div class="mode-live-row"><span>Поточний обсяг експорту</span><strong>${pluralRows(items.length)}</strong></div>
      <div class="mode-live-row"><span>Активних підрозділів</span><strong>${formatNumber(Object.keys(grouped).length)}</strong></div>
      <div class="mode-live-row"><span>Фільтри</span><strong>${getReportFilterLabel()}</strong></div>
    ` : `<div class="alert-item muted">Завантаж файл, щоб активувати PDF, Excel і пакет рішення.</div>`);
  }

  const analysisPanel = document.getElementById("analysisPanel");
  if (analysisPanel) {
    const observer = new MutationObserver(updateDashboardMetrics);
    observer.observe(analysisPanel, { childList: true, subtree: true });
  }

  bindReportsCenter();
  setInterval(updateDashboardMetrics, 1500);
  updateDashboardMetrics();
})();
