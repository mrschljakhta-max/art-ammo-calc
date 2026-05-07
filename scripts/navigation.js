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

    setText("dashTotalBalance", items.length ? formatNumber(totalBalance) : "—");
    setText("dashLongRange", items.length ? formatNumber(longRange) : "—");
    setText("dashCriticalCount", items.length ? formatNumber(critical) : "—");
    setText("dashUnitCount", Object.keys(grouped).length ? formatNumber(Object.keys(grouped).length) : "—");
    setText("dashRowsCount", items.length ? formatNumber(items.length) : "—");
    setText("dashExchangeCount", items.length ? formatNumber(exchangeCount) : "—");
    setText("dashQualityIssues", items.length ? formatNumber(qualityIssues) : "—");

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
  }

  const analysisPanel = document.getElementById("analysisPanel");
  if (analysisPanel) {
    const observer = new MutationObserver(updateDashboardMetrics);
    observer.observe(analysisPanel, { childList: true, subtree: true });
  }

  setInterval(updateDashboardMetrics, 1500);
  updateDashboardMetrics();
})();
