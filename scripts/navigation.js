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
    ` : `<div class="alert-item muted">Після аналізу система покаже помилки й попередження файлу.</div>`);

    setText("reportsPackagePreview", items.length ? "можна формувати" : "—");
    setHTML("reportsLiveList", items.length ? `
      <div class="mode-live-row"><span>Поточний обсяг експорту</span><strong>${pluralRows(items.length)}</strong></div>
      <div class="mode-live-row"><span>Активних підрозділів</span><strong>${formatNumber(Object.keys(grouped).length)}</strong></div>
    ` : `<div class="alert-item muted">Завантаж файл, щоб активувати PDF, Excel і пакет рішення.</div>`);
  }

  const analysisPanel = document.getElementById("analysisPanel");
  if (analysisPanel) {
    const observer = new MutationObserver(updateDashboardMetrics);
    observer.observe(analysisPanel, { childList: true, subtree: true });
  }

  setInterval(updateDashboardMetrics, 1500);
  updateDashboardMetrics();
})();
