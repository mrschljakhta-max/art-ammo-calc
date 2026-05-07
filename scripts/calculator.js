const analyzeBtn = document.getElementById("analyzeBtn");
const analysisPanel = document.getElementById("analysisPanel");
const filtersPanel = document.getElementById("filtersPanel");
const unitFilter = document.getElementById("unitFilter");
const longRangeOnly = document.getElementById("longRangeOnly");
const searchFilter = document.getElementById("searchFilter");
const resetFiltersBtn = document.getElementById("resetFiltersBtn");
const exportStatusesBtn = document.getElementById("exportStatusesBtn");
const importStatusesFile = document.getElementById("importStatusesFile");
const filterStatus = document.getElementById("filterStatus");
const lowBalanceThreshold = document.getElementById("lowBalanceThreshold");
const stockFilter = document.getElementById("stockFilter");
const sortFilter = document.getElementById("sortFilter");
const actionPriorityFilter = document.getElementById("actionPriorityFilter");
const actionLongOnly = document.getElementById("actionLongOnly");
const actionStatusFilter = document.getElementById("actionStatusFilter");
const exportDecisionPackageBtn = document.getElementById("exportDecisionPackageBtn");
const importDecisionPackageFile = document.getElementById("importDecisionPackageFile");
const appVersion = document.getElementById("appVersion");

function getAppMeta() {
  return window.ART_AMMO_APP_META || {
    appName: "Art Ammo",
    version: "0.25",
    buildLabel: "v25-auto-dictionaries",
    buildDate: "2026-05-07",
    logicProfile: "Excel локально + аналітика залишків + журнал дій"
  };
}

function initAppVersionBadge() {
  if (!appVersion) return;
  const meta = getAppMeta();
  appVersion.textContent = `${meta.appName} ${meta.version}`;
  appVersion.title = `${meta.buildLabel} · ${meta.buildDate}`;
}

initAppVersionBadge();

analyzeBtn.addEventListener("click", analyzeWorkbook);
unitFilter.addEventListener("change", applyFilters);
longRangeOnly.addEventListener("change", applyFilters);
searchFilter.addEventListener("input", applyFilters);
resetFiltersBtn.addEventListener("click", resetFilters);
if (exportStatusesBtn) exportStatusesBtn.addEventListener("click", exportActionStatusesToJson);
if (importStatusesFile) importStatusesFile.addEventListener("change", importActionStatusesFromJson);
lowBalanceThreshold.addEventListener("input", applyFilters);
stockFilter.addEventListener("change", applyFilters);
sortFilter.addEventListener("change", applyFilters);
if (actionPriorityFilter) actionPriorityFilter.addEventListener("change", applyFilters);
if (actionLongOnly) actionLongOnly.addEventListener("change", applyFilters);
if (actionStatusFilter) actionStatusFilter.addEventListener("change", applyFilters);
if (exportDecisionPackageBtn) exportDecisionPackageBtn.addEventListener("click", exportDecisionPackage);
if (importDecisionPackageFile) importDecisionPackageFile.addEventListener("change", importDecisionPackageFromFile);

function analyzeWorkbook() {
  const workbook = window.ArtAmmoState?.workbook;

  if (!workbook) {
    analysisPanel.innerHTML = `<p>Спочатку завантаж Excel.</p>`;
    return;
  }

  const allItems = [];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: ""
    });

    allItems.push(...parseStructuredSheet(sheetName, rows));
  });

  const unitItems = allItems.filter(item => !item.isSummarySheet);
  const summaryItems = allItems.filter(item => item.isSummarySheet);
  const grouped = groupByUnit(unitItems);

  window.ArtAmmoState.analysisItems = allItems;
  window.ArtAmmoState.unitItems = unitItems;
  window.ArtAmmoState.summaryItems = summaryItems;
  window.ArtAmmoState.groupedByUnit = grouped;
  window.ArtAmmoState.analyzedAt = new Date().toISOString();

  setupFilters(grouped);

  document.getElementById("exportExcelBtn").disabled = false;
  document.getElementById("exportPdfBtn").disabled = false;
  if (exportStatusesBtn) exportStatusesBtn.disabled = false;
  if (exportDecisionPackageBtn) exportDecisionPackageBtn.disabled = false;

  if (typeof setStatus === "function") {
    setStatus(`Проаналізовано: ${unitItems.length} рядків`, "ok");
  }

  const initialItems = getCurrentFilteredItems();
  updateFilterStatus(initialItems);
  renderAnalysis(initialItems, initialItems, summaryItems, groupByUnit(initialItems));
}

function parseStructuredSheet(sheetName, rows) {
  const items = [];
  const schema = ART_AMMO_SCHEMA;
  const cols = schema.COLUMNS;

  const isSummarySheet =
    sheetName.trim().toLowerCase() === schema.SUMMARY_SHEET_NAME.toLowerCase();

  rows.forEach((row, rowIndex) => {
    if (rowIndex < schema.START_ROW_INDEX) return;

    const category = cleanCell(row[cols.category]);
    const ammoText = cleanCell(row[cols.ammo]);
    const rangeMeters = toNumber(row[cols.rangeMeters]);
    const received = toNumber(row[cols.received]);
    const spent = toNumber(row[cols.spent]);
    const balance = toNumber(row[cols.balance]);

    if (!ammoText) return;
    if (isProbablyHeader(ammoText)) return;

    const parsed = parseAmmoName(ammoText);
    if (!parsed) return;

    const rangeKm = rangeMeters ? rangeMeters / 1000 : null;

    items.push({
      sheetName,
      unit: sheetName,
      row: rowIndex + 1,
      category,
      ammoRaw: ammoText,
      projectile: parsed.projectile,
      charge: parsed.charge,
      note: parsed.note,
      combination: `${parsed.projectile} + ${parsed.charge}`,
      rangeMeters,
      rangeKm,
      rangeLabel: rangeKm ? `${rangeKm.toFixed(1)} км` : "",
      longRange: Boolean(rangeKm && rangeKm >= schema.LONG_RANGE_KM),
      received,
      spent,
      quantity: balance,
      balance,
      isSummarySheet
    });
  });

  return items;
}

function parseAmmoName(text) {
  const value = cleanCell(text);
  const match = value.match(/^(.+?)\s*\((.+?)\)(.*)$/);

  if (!match) return null;

  return {
    projectile: normalizeName(match[1]),
    charge: normalizeName(match[2]),
    note: cleanCell(match[3])
  };
}

function normalizeName(value) {
  return cleanCell(value)
    .replace(/^М/g, "M")
    .replace(/\s*\/\s*/g, "/")
    .replace(/\s*,\s*/g, ", ")
    .replace(/\s*-\s*/g, "-")
    .trim();
}

function cleanCell(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return 0;

  const number = Number(
    String(value)
      .replace(",", ".")
      .replace(/\s+/g, "")
  );

  return Number.isNaN(number) ? 0 : number;
}

function isProbablyHeader(text) {
  const value = cleanCell(text).toLowerCase();

  return [
    "тип",
    "найменування",
    "дальність",
    "отримання",
    "витрата",
    "залишок"
  ].some(word => value.includes(word));
}

function groupByUnit(items) {
  const grouped = {};

  items.forEach((item) => {
    if (!grouped[item.unit]) {
      grouped[item.unit] = {
        unit: item.unit,
        items: [],
        totalReceived: 0,
        totalSpent: 0,
        totalBalance: 0,
        longRangeBalance: 0,
        combinations: new Set()
      };
    }

    grouped[item.unit].items.push(item);
    grouped[item.unit].totalReceived += Number(item.received || 0);
    grouped[item.unit].totalSpent += Number(item.spent || 0);
    grouped[item.unit].totalBalance += Number(item.balance || 0);

    if (item.longRange) {
      grouped[item.unit].longRangeBalance += Number(item.balance || 0);
    }

    grouped[item.unit].combinations.add(item.combination);
  });

  return grouped;
}

function enrichUnitStats(grouped, threshold = 10) {
  return Object.values(grouped).map(unit => {
    const zeroCount = unit.items.filter(item => Number(item.balance || 0) === 0).length;
    const lowCount = unit.items.filter(item => {
      const balance = Number(item.balance || 0);
      return balance > 0 && balance <= threshold;
    }).length;
    const longRangeShare = unit.totalBalance > 0
      ? Math.round((unit.longRangeBalance / unit.totalBalance) * 100)
      : 0;

    return {
      ...unit,
      zeroCount,
      lowCount,
      riskCount: zeroCount + lowCount,
      longRangeShare
    };
  });
}

function getUnitRanking(grouped, mode, limit = 6) {
  const threshold = getLowBalanceThreshold();
  const stats = enrichUnitStats(grouped, threshold);

  const sorters = {
    balance: (a, b) => Number(b.totalBalance || 0) - Number(a.totalBalance || 0),
    longRange: (a, b) => Number(b.longRangeBalance || 0) - Number(a.longRangeBalance || 0),
    risk: (a, b) => Number(b.riskCount || 0) - Number(a.riskCount || 0) || Number(a.totalBalance || 0) - Number(b.totalBalance || 0)
  };

  return stats
    .sort((a, b) => (sorters[mode] || sorters.balance)(a, b) || String(a.unit).localeCompare(String(b.unit), "uk"))
    .slice(0, limit);
}

function renderUnitRankList(units, valueKey, suffix = "") {
  if (!units.length) {
    return `<div class="insight-empty">Даних по підрозділах не знайдено</div>`;
  }

  return units.map((unit, index) => `
    <div class="unit-rank-row">
      <div class="rank-index">${index + 1}</div>
      <div class="rank-main">
        <strong>${escapeHtml(unit.unit)}</strong>
        <span>${unit.combinations.size} комбінацій · ${unit.longRangeShare}% далекобійних</span>
      </div>
      <div class="rank-value">${unit[valueKey]}${suffix}</div>
    </div>
  `).join("");
}


function getRecommendations(items, grouped) {
  const threshold = getLowBalanceThreshold();
  const unitStats = enrichUnitStats(grouped, threshold);
  const zeroItems = getCriticalItems(items);
  const lowItems = getLowBalanceItems(items, threshold);
  const longRangeItems = items.filter(item => item.longRange);
  const totalBalance = items.reduce((sum, item) => sum + Number(item.balance || 0), 0);
  const longRangeBalance = longRangeItems.reduce((sum, item) => sum + Number(item.balance || 0), 0);
  const recommendations = [];

  const push = (level, title, text, meta = "") => {
    recommendations.push({ level, title, text, meta });
  };

  if (!items.length) {
    push("neutral", "Дані відсутні", "За поточними фільтрами немає рядків для аналізу.");
    return recommendations;
  }

  const zeroLong = zeroItems.filter(item => item.longRange).length;
  if (zeroLong > 0) {
    push(
      "danger",
      "Критично: нульові далекобійні позиції",
      `Знайдено ${zeroLong} далекобійних позицій із нульовим залишком. Їх варто перевірити першочергово.`,
      "Пріоритет: високий"
    );
  }

  if (lowItems.length > 0) {
    push(
      "warning",
      "Є позиції з малим залишком",
      `Позицій із залишком від 1 до ${threshold}: ${lowItems.length}. Доцільно переглянути їх окремо через фільтр «малий залишок».`,
      `Поточний поріг: ≤${threshold}`
    );
  }

  const mostRiskyUnit = [...unitStats]
    .sort((a, b) => Number(b.riskCount || 0) - Number(a.riskCount || 0) || Number(a.totalBalance || 0) - Number(b.totalBalance || 0))[0];

  if (mostRiskyUnit && mostRiskyUnit.riskCount > 0) {
    push(
      "warning",
      "Підрозділ із найбільшим ризиком",
      `${mostRiskyUnit.unit}: ${mostRiskyUnit.riskCount} проблемних позицій, з них нульових — ${mostRiskyUnit.zeroCount}, малих — ${mostRiskyUnit.lowCount}.`,
      "Потрібна звірка залишків"
    );
  }

  const bestLongRangeUnit = [...unitStats]
    .filter(unit => Number(unit.longRangeBalance || 0) > 0)
    .sort((a, b) => Number(b.longRangeBalance || 0) - Number(a.longRangeBalance || 0))[0];

  if (bestLongRangeUnit) {
    push(
      "good",
      "Найбільший запас далекобійних",
      `${bestLongRangeUnit.unit}: ${bestLongRangeUnit.longRangeBalance} далекобійних залишків (${bestLongRangeUnit.longRangeShare}% від залишку підрозділу).`,
      "Сильна позиція"
    );
  }

  const noLongRangeUnits = unitStats.filter(unit => Number(unit.totalBalance || 0) > 0 && Number(unit.longRangeBalance || 0) === 0);
  if (noLongRangeUnits.length) {
    push(
      "neutral",
      "Підрозділи без далекобійних залишків",
      `За поточними даними таких підрозділів: ${noLongRangeUnits.length}. Перевір: ${noLongRangeUnits.slice(0, 3).map(unit => unit.unit).join(", ")}${noLongRangeUnits.length > 3 ? "…" : ""}.`,
      "Інформаційний контроль"
    );
  }

  if (totalBalance > 0) {
    const longShare = Math.round((longRangeBalance / totalBalance) * 100);
    if (longShare < 25) {
      push(
        "warning",
        "Низька частка далекобійних",
        `Далекобійні становлять приблизно ${longShare}% від поточного залишку. Варто окремо контролювати цю групу.`,
        "Баланс дальності"
      );
    } else {
      push(
        "good",
        "Баланс далекобійних прийнятний",
        `Далекобійні становлять приблизно ${longShare}% від поточного залишку.`,
        "Баланс дальності"
      );
    }
  }

  if (!zeroItems.length && !lowItems.length) {
    push(
      "good",
      "Критичних залишків не виявлено",
      "За поточним порогом нульові та малі залишки не знайдені.",
      "Поточний фільтр"
    );
  }

  return recommendations.slice(0, 6);
}

function renderRecommendations(recommendations) {
  if (!recommendations.length) return "";

  return `
    <div class="recommendation-panel">
      <div class="recommendation-header">
        <h2>Рекомендації системи</h2>
        <span>Автоматичні висновки за поточними фільтрами</span>
      </div>

      <div class="recommendation-grid">
        ${recommendations.map(item => `
          <div class="recommendation-card recommendation-${item.level}">
            <div class="recommendation-title">${escapeHtml(item.title)}</div>
            <div class="recommendation-text">${escapeHtml(item.text)}</div>
            ${item.meta ? `<div class="recommendation-meta">${escapeHtml(item.meta)}</div>` : ""}
          </div>
        `).join("")}
      </div>
    </div>
  `;
}


function getCommanderSummary(items, grouped) {
  const threshold = getLowBalanceThreshold();
  const totalBalance = items.reduce((sum, item) => sum + Number(item.balance || 0), 0);
  const totalReceived = items.reduce((sum, item) => sum + Number(item.received || 0), 0);
  const totalSpent = items.reduce((sum, item) => sum + Number(item.spent || 0), 0);
  const longRangeBalance = items
    .filter(item => item.longRange)
    .reduce((sum, item) => sum + Number(item.balance || 0), 0);
  const shortRangeBalance = totalBalance - longRangeBalance;
  const zeroItems = getCriticalItems(items);
  const lowItems = getLowBalanceItems(items, threshold);
  const unitStats = enrichUnitStats(grouped, threshold);
  const longShare = totalBalance > 0 ? Math.round((longRangeBalance / totalBalance) * 100) : 0;
  const riskShare = items.length > 0 ? Math.round(((zeroItems.length + lowItems.length) / items.length) * 100) : 0;

  const mostStockedUnit = [...unitStats]
    .sort((a, b) => Number(b.totalBalance || 0) - Number(a.totalBalance || 0))[0];
  const bestLongRangeUnit = [...unitStats]
    .sort((a, b) => Number(b.longRangeBalance || 0) - Number(a.longRangeBalance || 0))[0];
  const mostRiskyUnit = [...unitStats]
    .sort((a, b) => Number(b.riskCount || 0) - Number(a.riskCount || 0))[0];

  const topZeroLong = zeroItems
    .filter(item => item.longRange)
    .slice(0, 5)
    .map(item => `${item.unit}: ${item.projectile} (${item.charge})`);

  const lines = [];

  lines.push({
    label: "Загальна оцінка",
    text: `За поточним відбором обліковано ${items.length} позицій по ${Object.keys(grouped).length} підрозділах. Загальний залишок — ${totalBalance}, з них далекобійних — ${longRangeBalance} (${longShare}%).`
  });

  lines.push({
    label: "Рух БК",
    text: `Отримання — ${totalReceived}, витрата — ${totalSpent}, поточний баланс — ${totalBalance}. Недалекобійних залишків — ${shortRangeBalance}.`
  });

  lines.push({
    label: "Критичність",
    text: `Нульових позицій — ${zeroItems.length}, малих залишків ≤${threshold} — ${lowItems.length}. Частка проблемних позицій — близько ${riskShare}%.`
  });

  if (mostRiskyUnit && mostRiskyUnit.riskCount > 0) {
    lines.push({
      label: "Першочергова увага",
      text: `${mostRiskyUnit.unit}: ${mostRiskyUnit.riskCount} проблемних позицій, з них нульових — ${mostRiskyUnit.zeroCount}, малих — ${mostRiskyUnit.lowCount}.`
    });
  }

  if (bestLongRangeUnit && Number(bestLongRangeUnit.longRangeBalance || 0) > 0) {
    lines.push({
      label: "Далекобійний ресурс",
      text: `Найбільший запас далекобійних має ${bestLongRangeUnit.unit}: ${bestLongRangeUnit.longRangeBalance}.`
    });
  }

  if (mostStockedUnit && Number(mostStockedUnit.totalBalance || 0) > 0) {
    lines.push({
      label: "Найбільший загальний запас",
      text: `${mostStockedUnit.unit}: ${mostStockedUnit.totalBalance} залишку по всіх врахованих позиціях.`
    });
  }

  if (topZeroLong.length) {
    lines.push({
      label: "Нульові далекобійні",
      text: `Перші позиції для перевірки: ${topZeroLong.join("; ")}.`
    });
  }

  const conclusion = [];
  if (zeroItems.length || lowItems.length) {
    conclusion.push(`перевірити проблемні позиції (${zeroItems.length + lowItems.length})`);
  }
  if (longShare < 25 && totalBalance > 0) {
    conclusion.push("окремо контролювати далекобійний ресурс");
  }
  if (mostRiskyUnit && mostRiskyUnit.riskCount > 0) {
    conclusion.push(`почати звірку з ${mostRiskyUnit.unit}`);
  }

  lines.push({
    label: "Короткий висновок",
    text: conclusion.length
      ? `Рекомендовано: ${conclusion.join("; ")}.`
      : "Критичних відхилень за поточними фільтрами не виявлено."
  });

  return lines;
}

function renderCommanderSummary(summary) {
  if (!summary.length) return "";

  return `
    <div class="commander-summary-panel">
      <div class="commander-summary-header">
        <h2>Короткий командирський висновок</h2>
        <span>Стислий підсумок для швидкого огляду</span>
      </div>

      <div class="commander-summary-list">
        ${summary.map(item => `
          <div class="commander-summary-row">
            <div class="commander-summary-label">${escapeHtml(item.label)}</div>
            <div class="commander-summary-text">${escapeHtml(item.text)}</div>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}


function formatDateTime(value) {
  if (!value) return "—";

  try {
    return new Date(value).toLocaleString("uk-UA");
  } catch (error) {
    return "—";
  }
}

function getStockFilterLabel() {
  const labels = {
    all: "всі залишки",
    problem: "вузькі місця",
    zero: "нульові",
    low: "малий залишок"
  };

  return labels[stockFilter?.value] || "всі залишки";
}

function getSortLabel() {
  const labels = {
    default: "як у файлі",
    balanceAsc: "залишок ↑",
    balanceDesc: "залишок ↓",
    rangeDesc: "дальність ↓",
    rangeAsc: "дальність ↑",
    unitAsc: "підрозділ А-Я",
    projectileAsc: "снаряд А-Я"
  };

  return labels[sortFilter?.value] || "як у файлі";
}

function getActiveFilterDescription() {
  const active = [];

  if (unitFilter?.value && unitFilter.value !== "all") {
    active.push(`підрозділ: ${unitFilter.value}`);
  }

  if (longRangeOnly?.checked) {
    active.push("тільки далекобійні");
  }

  if (searchFilter?.value?.trim()) {
    active.push(`пошук: ${searchFilter.value.trim()}`);
  }

  if (stockFilter?.value && stockFilter.value !== "all") {
    active.push(`залишки: ${getStockFilterLabel()}`);
  }

  const actionDescription = getActionFilterDescription();
  if (actionDescription !== "усі дії") {
    active.push(`журнал дій: ${actionDescription}`);
  }

  return active.length ? active.join("; ") : "без додаткових фільтрів";
}

function getActionFilterDescription() {
  const parts = [];
  const priority = actionPriorityFilter?.value || "all";

  if (priority !== "all") {
    parts.push(`пріоритет ${priority}`);
  }

  if (actionLongOnly?.checked) {
    parts.push("тільки далекобійні");
  }

  const status = actionStatusFilter?.value || "all";
  if (status !== "all") {
    const labels = { planned: "заплановано", done: "виконано", rejected: "відхилено" };
    parts.push(`статус: ${labels[status] || status}`);
  }

  return parts.length ? parts.join(", ") : "усі дії";
}

function getReportPassport(items, grouped) {
  const totalRows = window.ArtAmmoState?.unitItems?.length || 0;
  const fileName = window.ArtAmmoState?.fileName || "—";
  const loadedAt = formatDateTime(window.ArtAmmoState?.loadedAt);
  const analyzedAt = formatDateTime(window.ArtAmmoState?.analyzedAt);
  const threshold = getLowBalanceThreshold();

  const meta = getAppMeta();

  return [
    { label: "Файл Excel", value: fileName },
    { label: "Версія системи", value: `${meta.appName} ${meta.version}` },
    { label: "Збірка", value: `${meta.buildLabel} / ${meta.buildDate}` },
    { label: "Профіль логіки", value: meta.logicProfile || "—" },
    { label: "Час завантаження", value: loadedAt },
    { label: "Час аналізу", value: analyzedAt },
    { label: "Показано рядків", value: `${items.length} з ${totalRows}` },
    { label: "Підрозділів у вибірці", value: String(Object.keys(grouped || {}).length) },
    { label: "Активні фільтри", value: getActiveFilterDescription() },
    { label: "Поріг малого залишку", value: `≤${threshold}` },
    { label: "Фільтр залишків", value: getStockFilterLabel() },
    { label: "Сортування", value: getSortLabel() },
    { label: "Фільтр журналу дій", value: getActionFilterDescription() }
  ];
}

function renderReportPassport(passport) {
  if (!passport?.length) return "";

  return `
    <div class="report-passport-panel">
      <div class="report-passport-header">
        <h2>Паспорт звіту</h2>
        <span>Контекст формування поточної вибірки</span>
      </div>

      <div class="report-passport-grid">
        ${passport.map(item => `
          <div class="report-passport-item">
            <div class="report-passport-label">${escapeHtml(item.label)}</div>
            <div class="report-passport-value">${escapeHtml(item.value)}</div>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function getActionStatusIntegrity(plan) {
  const currentPlan = Array.isArray(plan) ? plan : [];
  const statuses = getAllActionStatuses();
  const allowed = new Set(["planned", "done", "rejected"]);
  const currentIds = new Set(currentPlan.map(item => item.actionId));
  const counters = {
    total: currentPlan.length,
    explicit: 0,
    missing: 0,
    planned: 0,
    done: 0,
    rejected: 0,
    orphaned: 0,
    importedTotal: 0
  };

  currentPlan.forEach(item => {
    const hasSaved = Object.prototype.hasOwnProperty.call(statuses, item.actionId);
    const savedStatus = hasSaved && allowed.has(String(statuses[item.actionId]))
      ? String(statuses[item.actionId])
      : "planned";

    if (hasSaved) {
      counters.explicit += 1;
    } else {
      counters.missing += 1;
    }

    if (savedStatus === "done") counters.done += 1;
    else if (savedStatus === "rejected") counters.rejected += 1;
    else counters.planned += 1;
  });

  Object.entries(statuses).forEach(([actionId, status]) => {
    if (!allowed.has(String(status))) return;

    counters.importedTotal += 1;

    if (!currentIds.has(actionId)) {
      counters.orphaned += 1;
    }
  });

  return counters;
}

function renderActionStatusIntegrity(integrity) {
  if (!integrity || !integrity.total) return "";

  const healthClass = integrity.orphaned > 0
    ? "integrity-warning"
    : integrity.missing > 0
      ? "integrity-soft"
      : "integrity-ok";

  const healthText = integrity.orphaned > 0
    ? "Є старі/зайві статуси з імпортованого JSON"
    : integrity.missing > 0
      ? "Частина дій ще без явно збереженого статусу"
      : "Статуси відповідають поточному журналу дій";

  return `
    <div class="action-integrity-panel ${healthClass}">
      <div class="action-integrity-header">
        <h2>Контроль статусів журналу дій</h2>
        <span>${escapeHtml(healthText)}</span>
      </div>

      <div class="action-integrity-grid">
        <div class="action-integrity-card">
          <div class="impact-label">Поточних дій</div>
          <div class="impact-value">${integrity.total}</div>
        </div>

        <div class="action-integrity-card">
          <div class="impact-label">Явно збережено</div>
          <div class="impact-value">${integrity.explicit}</div>
        </div>

        <div class="action-integrity-card ${integrity.missing ? "integrity-card-warning" : ""}">
          <div class="impact-label">Без рішення</div>
          <div class="impact-value">${integrity.missing}</div>
        </div>

        <div class="action-integrity-card">
          <div class="impact-label">Заплановано</div>
          <div class="impact-value">${integrity.planned}</div>
        </div>

        <div class="action-integrity-card">
          <div class="impact-label">Виконано</div>
          <div class="impact-value">${integrity.done}</div>
        </div>

        <div class="action-integrity-card">
          <div class="impact-label">Відхилено</div>
          <div class="impact-value">${integrity.rejected}</div>
        </div>

        <div class="action-integrity-card ${integrity.orphaned ? "integrity-card-danger" : ""}">
          <div class="impact-label">Зайві статуси JSON</div>
          <div class="impact-value">${integrity.orphaned}</div>
        </div>
      </div>

      <div class="action-integrity-note">
        “Без рішення” означає, що дія показана в поточному журналі, але її статус ще не був явно змінений або збережений. “Зайві статуси JSON” означають, що в імпортованому файлі є статуси для дій, яких уже немає в поточних рекомендаціях.
      </div>
    </div>
  `;
}

function groupByCategory(items) {
  const grouped = {};

  items.forEach((item) => {
    const category = item.category || "Без категорії";

    if (!grouped[category]) {
      grouped[category] = {
        category,
        items: [],
        totalReceived: 0,
        totalSpent: 0,
        totalBalance: 0,
        longRangeBalance: 0,
        combinations: new Set()
      };
    }

    grouped[category].items.push(item);
    grouped[category].totalReceived += Number(item.received || 0);
    grouped[category].totalSpent += Number(item.spent || 0);
    grouped[category].totalBalance += Number(item.balance || 0);

    if (item.longRange) {
      grouped[category].longRangeBalance += Number(item.balance || 0);
    }

    grouped[category].combinations.add(item.combination);
  });

  return grouped;
}

function getRangeBand(item) {
  const km = Number(item.rangeKm || 0);

  if (!km) return "Без дальності";
  if (km < ART_AMMO_SCHEMA.LONG_RANGE_KM) return `до ${ART_AMMO_SCHEMA.LONG_RANGE_KM} км`;
  if (km < 25) return `${ART_AMMO_SCHEMA.LONG_RANGE_KM}–25 км`;
  if (km < 30) return `25–30 км`;
  return `30+ км`;
}

function groupByRangeBand(items) {
  const grouped = {};

  items.forEach((item) => {
    const band = getRangeBand(item);

    if (!grouped[band]) {
      grouped[band] = {
        band,
        items: [],
        totalBalance: 0,
        combinations: new Set()
      };
    }

    grouped[band].items.push(item);
    grouped[band].totalBalance += Number(item.balance || 0);
    grouped[band].combinations.add(item.combination);
  });

  const order = [
    `до ${ART_AMMO_SCHEMA.LONG_RANGE_KM} км`,
    `${ART_AMMO_SCHEMA.LONG_RANGE_KM}–25 км`,
    `25–30 км`,
    `30+ км`,
    "Без дальності"
  ];

  return Object.fromEntries(
    Object.entries(grouped).sort((a, b) => order.indexOf(a[0]) - order.indexOf(b[0]))
  );
}

function getCriticalItems(items) {
  return [...items]
    .filter(item => Number(item.balance || 0) === 0)
    .sort(sortCriticalItems);
}

function getLowBalanceItems(items, threshold = 10) {
  return [...items]
    .filter(item => Number(item.balance || 0) > 0 && Number(item.balance || 0) <= threshold)
    .sort(sortCriticalItems);
}


function getTopBalanceItems(items, limit = 5) {
  return [...items]
    .filter(item => Number(item.balance || 0) > 0)
    .sort((a, b) =>
      Number(b.balance || 0) - Number(a.balance || 0) ||
      Number(b.rangeKm || 0) - Number(a.rangeKm || 0) ||
      String(a.unit).localeCompare(String(b.unit), "uk")
    )
    .slice(0, limit);
}

function getTopLongRangeItems(items, limit = 5) {
  return [...items]
    .filter(item => item.longRange && Number(item.balance || 0) > 0)
    .sort((a, b) =>
      Number(b.rangeKm || 0) - Number(a.rangeKm || 0) ||
      Number(b.balance || 0) - Number(a.balance || 0) ||
      String(a.unit).localeCompare(String(b.unit), "uk")
    )
    .slice(0, limit);
}

function getTopRiskItems(items, threshold = 10, limit = 5) {
  return [...items]
    .filter(item => Number(item.balance || 0) <= threshold)
    .sort((a, b) => {
      const aZero = Number(a.balance || 0) === 0 ? 0 : 1;
      const bZero = Number(b.balance || 0) === 0 ? 0 : 1;
      if (aZero !== bZero) return aZero - bZero;

      const aLong = a.longRange ? 0 : 1;
      const bLong = b.longRange ? 0 : 1;
      if (aLong !== bLong) return aLong - bLong;

      return Number(a.balance || 0) - Number(b.balance || 0) ||
        Number(b.rangeKm || 0) - Number(a.rangeKm || 0) ||
        String(a.unit).localeCompare(String(b.unit), "uk");
    })
    .slice(0, limit);
}

function renderInsightList(items, emptyText) {
  if (!items.length) {
    return `<div class="insight-empty">${escapeHtml(emptyText)}</div>`;
  }

  return items.map(item => `
    <div class="insight-row ${getRowClass(item)}">
      <div class="insight-main">
        <strong>${escapeHtml(item.projectile)}</strong>
        <span>${escapeHtml(item.charge)}</span>
      </div>
      <div class="insight-meta">
        <span>${escapeHtml(item.unit)}</span>
        <span>${item.rangeKm ? item.rangeKm.toFixed(1) + " км" : "без дальності"}</span>
        <b>${item.balance}</b>
      </div>
    </div>
  `).join("");
}

function sortCriticalItems(a, b) {
  const aLong = a.longRange ? 0 : 1;
  const bLong = b.longRange ? 0 : 1;

  if (aLong !== bLong) return aLong - bLong;

  const balanceDiff = Number(a.balance || 0) - Number(b.balance || 0);
  if (balanceDiff !== 0) return balanceDiff;

  return String(a.unit).localeCompare(String(b.unit), "uk");
}

function setupFilters(grouped) {
  const previousUnit = unitFilter.value || "all";

  unitFilter.innerHTML = `<option value="all">Всі підрозділи</option>`;

  Object.keys(grouped).forEach(unitName => {
    const option = document.createElement("option");
    option.value = unitName;
    option.textContent = unitName;
    unitFilter.appendChild(option);
  });

  if ([...unitFilter.options].some(option => option.value === previousUnit)) {
    unitFilter.value = previousUnit;
  }

  setupCompareControls(grouped);

  filtersPanel.hidden = false;
}


function setupCompareControls(grouped) {
  if (!filtersPanel) return;

  let comparePanel = document.getElementById("comparePanel");

  if (!comparePanel) {
    comparePanel = document.createElement("div");
    comparePanel.id = "comparePanel";
    comparePanel.className = "compare-panel";
    comparePanel.innerHTML = `
      <div class="compare-title">Порівняння підрозділів</div>
      <select id="compareUnitA" class="filter-control">
        <option value="">Підрозділ A</option>
      </select>
      <select id="compareUnitB" class="filter-control">
        <option value="">Підрозділ B</option>
      </select>
    `;
    filtersPanel.insertAdjacentElement("afterend", comparePanel);
  }

  const compareUnitA = document.getElementById("compareUnitA");
  const compareUnitB = document.getElementById("compareUnitB");

  const previousA = compareUnitA.value;
  const previousB = compareUnitB.value;
  const unitNames = Object.keys(grouped);

  compareUnitA.innerHTML = `<option value="">Підрозділ A</option>`;
  compareUnitB.innerHTML = `<option value="">Підрозділ B</option>`;

  unitNames.forEach(unitName => {
    const optionA = document.createElement("option");
    optionA.value = unitName;
    optionA.textContent = unitName;
    compareUnitA.appendChild(optionA);

    const optionB = document.createElement("option");
    optionB.value = unitName;
    optionB.textContent = unitName;
    compareUnitB.appendChild(optionB);
  });

  if (unitNames.includes(previousA)) compareUnitA.value = previousA;
  if (unitNames.includes(previousB)) compareUnitB.value = previousB;

  compareUnitA.onchange = applyFilters;
  compareUnitB.onchange = applyFilters;
}

function getComparisonData(items) {
  const compareUnitA = document.getElementById("compareUnitA");
  const compareUnitB = document.getElementById("compareUnitB");

  if (!compareUnitA || !compareUnitB) return null;

  const unitA = compareUnitA.value;
  const unitB = compareUnitB.value;

  if (!unitA || !unitB || unitA === unitB) return null;

  const aItems = items.filter(item => item.unit === unitA);
  const bItems = items.filter(item => item.unit === unitB);

  const sum = (arr, key) => arr.reduce((total, item) => total + Number(item[key] || 0), 0);
  const countZero = arr => arr.filter(item => Number(item.balance || 0) === 0).length;
  const countLow = arr => arr.filter(item => {
    const balance = Number(item.balance || 0);
    return balance > 0 && balance <= getLowBalanceThreshold();
  }).length;

  const a = {
    unit: unitA,
    rows: aItems.length,
    received: sum(aItems, "received"),
    spent: sum(aItems, "spent"),
    balance: sum(aItems, "balance"),
    longRange: sum(aItems.filter(item => item.longRange), "balance"),
    zero: countZero(aItems),
    low: countLow(aItems)
  };

  const b = {
    unit: unitB,
    rows: bItems.length,
    received: sum(bItems, "received"),
    spent: sum(bItems, "spent"),
    balance: sum(bItems, "balance"),
    longRange: sum(bItems.filter(item => item.longRange), "balance"),
    zero: countZero(bItems),
    low: countLow(bItems)
  };

  return { a, b };
}

function renderComparisonPanel(items) {
  const data = getComparisonData(items);

  if (!data) return "";

  const rows = [
    ["Рядків", data.a.rows, data.b.rows],
    ["Отримання", data.a.received, data.b.received],
    ["Витрата", data.a.spent, data.b.spent],
    ["Залишок", data.a.balance, data.b.balance],
    ["Далекобійних", data.a.longRange, data.b.longRange],
    ["Нульових позицій", data.a.zero, data.b.zero],
    ["Малих залишків", data.a.low, data.b.low]
  ];

  const strongerBalance = data.a.balance === data.b.balance
    ? "Загальний залишок однаковий."
    : data.a.balance > data.b.balance
      ? `${data.a.unit} має більший загальний залишок на ${data.a.balance - data.b.balance}.`
      : `${data.b.unit} має більший загальний залишок на ${data.b.balance - data.a.balance}.`;

  const strongerLong = data.a.longRange === data.b.longRange
    ? "Далекобійний залишок однаковий."
    : data.a.longRange > data.b.longRange
      ? `${data.a.unit} має більший далекобійний залишок на ${data.a.longRange - data.b.longRange}.`
      : `${data.b.unit} має більший далекобійний залишок на ${data.b.longRange - data.a.longRange}.`;

  const riskA = data.a.zero + data.a.low;
  const riskB = data.b.zero + data.b.low;
  const riskLine = riskA === riskB
    ? "Кількість проблемних позицій однакова."
    : riskA > riskB
      ? `${data.a.unit} має більше проблемних позицій на ${riskA - riskB}.`
      : `${data.b.unit} має більше проблемних позицій на ${riskB - riskA}.`;

  return `
    <div class="comparison-result-panel">
      <div class="comparison-header">
        <h2>Порівняння підрозділів</h2>
        <span>${escapeHtml(data.a.unit)} ↔ ${escapeHtml(data.b.unit)}</span>
      </div>

      <div class="table-wrap compact-wrap">
        <table>
          <thead>
            <tr>
              <th>Показник</th>
              <th>${escapeHtml(data.a.unit)}</th>
              <th>${escapeHtml(data.b.unit)}</th>
              <th>Різниця A-B</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(row => `
              <tr>
                <td>${escapeHtml(row[0])}</td>
                <td>${row[1]}</td>
                <td>${row[2]}</td>
                <td>${Number(row[1]) - Number(row[2])}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>

      <div class="comparison-notes">
        <div>${escapeHtml(strongerBalance)}</div>
        <div>${escapeHtml(strongerLong)}</div>
        <div>${escapeHtml(riskLine)}</div>
      </div>
    </div>
  `;
}

function applyFilters() {
  const filtered = getCurrentFilteredItems();
  const grouped = groupByUnit(filtered);

  updateFilterStatus(filtered);
  renderAnalysis(filtered, filtered, [], grouped);
}

function getCurrentFilteredItems() {
  const items = window.ArtAmmoState?.unitItems || [];
  let filtered = [...items];

  if (unitFilter.value !== "all") {
    filtered = filtered.filter(item => item.unit === unitFilter.value);
  }

  if (longRangeOnly.checked) {
    filtered = filtered.filter(item => item.longRange);
  }

  const searchValue = searchFilter.value.trim().toLowerCase();

  if (searchValue) {
    filtered = filtered.filter(item =>
      String(item.unit).toLowerCase().includes(searchValue) ||
      String(item.projectile).toLowerCase().includes(searchValue) ||
      String(item.charge).toLowerCase().includes(searchValue) ||
      String(item.category).toLowerCase().includes(searchValue) ||
      String(item.note).toLowerCase().includes(searchValue)
    );
  }

  const threshold = getLowBalanceThreshold();

  if (stockFilter.value === "problem") {
    filtered = filtered.filter(item =>
      Number(item.balance || 0) === 0 ||
      (Number(item.balance || 0) > 0 && Number(item.balance || 0) <= threshold)
    );
  }

  if (stockFilter.value === "zero") {
    filtered = filtered.filter(item => Number(item.balance || 0) === 0);
  }

  if (stockFilter.value === "low") {
    filtered = filtered.filter(item =>
      Number(item.balance || 0) > 0 && Number(item.balance || 0) <= threshold
    );
  }

  return sortItems(filtered);
}

function sortItems(items) {
  const sorted = [...items];
  const mode = sortFilter?.value || "default";

  const byText = (a, b, key) =>
    String(a[key] || "").localeCompare(String(b[key] || ""), "uk", { numeric: true });

  const byNumber = (a, b, key) =>
    Number(a[key] || 0) - Number(b[key] || 0);

  if (mode === "balanceAsc") {
    sorted.sort((a, b) => byNumber(a, b, "balance") || byText(a, b, "unit"));
  }

  if (mode === "balanceDesc") {
    sorted.sort((a, b) => byNumber(b, a, "balance") || byText(a, b, "unit"));
  }

  if (mode === "rangeDesc") {
    sorted.sort((a, b) => byNumber(b, a, "rangeKm") || byNumber(b, a, "balance"));
  }

  if (mode === "rangeAsc") {
    sorted.sort((a, b) => byNumber(a, b, "rangeKm") || byNumber(a, b, "balance"));
  }

  if (mode === "unitAsc") {
    sorted.sort((a, b) => byText(a, b, "unit") || byText(a, b, "projectile"));
  }

  if (mode === "projectileAsc") {
    sorted.sort((a, b) => byText(a, b, "projectile") || byText(a, b, "charge"));
  }

  return sorted;
}

function resetFilters() {
  unitFilter.value = "all";
  longRangeOnly.checked = false;
  searchFilter.value = "";
  lowBalanceThreshold.value = "10";
  stockFilter.value = "all";
  sortFilter.value = "default";
  if (actionPriorityFilter) actionPriorityFilter.value = "all";
  if (actionLongOnly) actionLongOnly.checked = false;
  if (actionStatusFilter) actionStatusFilter.value = "all";
  applyFilters();
}

function getLowBalanceThreshold() {
  const value = Number(lowBalanceThreshold?.value || 10);

  if (!Number.isFinite(value) || value <= 0) {
    return 10;
  }

  return Math.floor(value);
}

function updateFilterStatus(items) {
  const total = window.ArtAmmoState?.unitItems?.length || 0;
  const active = [];

  if (unitFilter.value !== "all") active.push(`підрозділ: ${unitFilter.value}`);
  if (longRangeOnly.checked) active.push("тільки далекобійні");
  if (searchFilter.value.trim()) active.push(`пошук: ${searchFilter.value.trim()}`);

  const stockFilterLabels = {
    problem: "тільки вузькі місця",
    zero: "тільки нульові",
    low: "тільки малий залишок"
  };

  if (stockFilter.value !== "all") {
    active.push(stockFilterLabels[stockFilter.value] || "фільтр залишків");
  }

  const sortLabels = {
    balanceAsc: "залишок ↑",
    balanceDesc: "залишок ↓",
    rangeDesc: "дальність ↓",
    rangeAsc: "дальність ↑",
    unitAsc: "підрозділ А-Я",
    projectileAsc: "снаряд А-Я"
  };

  if (sortFilter.value !== "default") {
    active.push(`сортування: ${sortLabels[sortFilter.value] || sortFilter.value}`);
  }

  active.push(`малий залишок ≤${getLowBalanceThreshold()}`);

  filterStatus.hidden = false;
  filterStatus.textContent = active.length
    ? `Показано ${items.length} з ${total}. Активні фільтри: ${active.join(", ")}.`
    : `Показано всі рядки: ${items.length}.`;
}


function getDataQualityReport(items, summaryItems = [], grouped = {}) {
  const threshold = getLowBalanceThreshold();
  const issues = [];
  const seen = new Map();

  items.forEach(item => {
    const location = `${item.unit} · рядок ${item.row}`;

    if (!item.projectile || !item.charge) {
      issues.push({
        severity: "critical",
        type: "Неповна комбінація",
        location,
        details: `Не вдалося визначити снаряд або заряд: ${item.ammoRaw || item.combination || ""}`
      });
    }

    if (!item.rangeKm || Number(item.rangeKm) <= 0) {
      issues.push({
        severity: "warning",
        type: "Немає дальності",
        location,
        details: `${item.combination}: дальність порожня або 0`
      });
    }

    if (Number(item.balance) < 0 || Number(item.received) < 0 || Number(item.spent) < 0) {
      issues.push({
        severity: "critical",
        type: "Від’ємне значення",
        location,
        details: `${item.combination}: отримання ${item.received}, витрата ${item.spent}, залишок ${item.balance}`
      });
    }

    const hasMovement = Number(item.received || 0) || Number(item.spent || 0) || Number(item.balance || 0);
    const expectedBalance = Number(item.received || 0) - Number(item.spent || 0);

    if (hasMovement && Number.isFinite(expectedBalance) && expectedBalance !== Number(item.balance || 0)) {
      issues.push({
        severity: "info",
        type: "Контроль формули",
        location,
        details: `${item.combination}: отримання - витрата = ${expectedBalance}, у таблиці залишок ${item.balance}`
      });
    }

    const key = `${item.unit}|${item.projectile}|${item.charge}|${item.rangeMeters}`;
    if (!seen.has(key)) seen.set(key, []);
    seen.get(key).push(item.row);
  });

  seen.forEach((rows, key) => {
    if (rows.length <= 1) return;
    const [unit, projectile, charge] = key.split("|");
    issues.push({
      severity: "warning",
      type: "Дубль позиції",
      location: `${unit} · рядки ${rows.join(", ")}`,
      details: `${projectile} (${charge}) повторюється ${rows.length} рази`
    });
  });

  const summaryTotal = summaryItems.reduce((sum, item) => sum + Number(item.balance || 0), 0);
  const unitTotal = items.reduce((sum, item) => sum + Number(item.balance || 0), 0);

  if (summaryItems.length && summaryTotal !== unitTotal) {
    issues.push({
      severity: "warning",
      type: "Розбіжність зі зведеним аркушем",
      location: ART_AMMO_SCHEMA.SUMMARY_SHEET_NAME,
      details: `Підрозділи: ${unitTotal}; зведений аркуш: ${summaryTotal}; різниця: ${unitTotal - summaryTotal}`
    });
  }

  const bySeverity = issues.reduce((acc, issue) => {
    acc[issue.severity] = (acc[issue.severity] || 0) + 1;
    return acc;
  }, {});

  const unitsWithoutRows = Object.values(grouped)
    .filter(unit => !unit.items || !unit.items.length)
    .map(unit => unit.unit);

  return {
    totalIssues: issues.length,
    critical: bySeverity.critical || 0,
    warning: bySeverity.warning || 0,
    info: bySeverity.info || 0,
    lowBalanceThreshold: threshold,
    summaryCompared: Boolean(summaryItems.length),
    unitsWithoutRows,
    issues: issues.slice(0, 80)
  };
}

function renderDataQualityPanel(report) {
  if (!report) return "";

  const statusClass = report.critical
    ? "quality-critical"
    : report.warning
      ? "quality-warning"
      : "quality-ok";

  const statusText = report.critical
    ? "Є критичні помилки"
    : report.warning
      ? "Є попередження"
      : "Явних проблем не знайдено";

  const issueRows = report.issues && report.issues.length
    ? report.issues.map(issue => `
        <tr class="quality-row-${escapeHtml(issue.severity)}">
          <td><span class="quality-badge ${escapeHtml(issue.severity)}">${escapeHtml(issue.severity)}</span></td>
          <td>${escapeHtml(issue.type)}</td>
          <td>${escapeHtml(issue.location)}</td>
          <td>${escapeHtml(issue.details)}</td>
        </tr>
      `).join("")
    : `
        <tr>
          <td colspan="4">Після базової перевірки проблем не виявлено.</td>
        </tr>
      `;

  return `
    <div class="data-quality-panel ${statusClass}">
      <div class="data-quality-header">
        <div>
          <h2>Контроль якості Excel</h2>
          <p>${escapeHtml(statusText)}</p>
        </div>
        <div class="data-quality-score">
          <span>${report.totalIssues}</span>
          <small>зауважень</small>
        </div>
      </div>

      <div class="quality-metrics-grid">
        <div class="quality-metric"><b>${report.critical}</b><span>критичних</span></div>
        <div class="quality-metric"><b>${report.warning}</b><span>попереджень</span></div>
        <div class="quality-metric"><b>${report.info}</b><span>інформаційних</span></div>
        <div class="quality-metric"><b>${report.summaryCompared ? "Так" : "Ні"}</b><span>звірка зі зведеним</span></div>
      </div>

      <div class="table-wrap compact-wrap quality-table-wrap">
        <table>
          <thead>
            <tr>
              <th>Рівень</th>
              <th>Тип</th>
              <th>Де</th>
              <th>Деталі</th>
            </tr>
          </thead>
          <tbody>${issueRows}</tbody>
        </table>
      </div>
    </div>
  `;
}


function buildAutoDictionaries(items) {
  const source = items || [];
  const unitsMap = new Map();
  const projectilesMap = new Map();
  const chargesMap = new Map();
  const combinationsMap = new Map();

  source.forEach(item => {
    const unitKey = cleanCell(item.unit || "");
    const projectileKey = cleanCell(item.projectile || "");
    const chargeKey = cleanCell(item.charge || "");
    const combinationKey = `${projectileKey} + ${chargeKey}`;

    if (unitKey && !unitsMap.has(unitKey)) {
      unitsMap.set(unitKey, {
        name: unitKey,
        rows: 0,
        totalBalance: 0,
        longRangeBalance: 0
      });
    }

    if (projectileKey && !projectilesMap.has(projectileKey)) {
      projectilesMap.set(projectileKey, {
        name: projectileKey,
        rows: 0,
        totalBalance: 0
      });
    }

    if (chargeKey && !chargesMap.has(chargeKey)) {
      chargesMap.set(chargeKey, {
        name: chargeKey,
        rows: 0,
        totalBalance: 0
      });
    }

    if (projectileKey && chargeKey && !combinationsMap.has(combinationKey)) {
      combinationsMap.set(combinationKey, {
        projectile: projectileKey,
        charge: chargeKey,
        combination: combinationKey,
        rows: 0,
        totalBalance: 0,
        maxRangeKm: 0,
        longRange: false,
        units: new Set()
      });
    }

    const balance = Number(item.balance || 0);

    if (unitsMap.has(unitKey)) {
      const unit = unitsMap.get(unitKey);
      unit.rows += 1;
      unit.totalBalance += balance;
      if (item.longRange) unit.longRangeBalance += balance;
    }

    if (projectilesMap.has(projectileKey)) {
      const projectile = projectilesMap.get(projectileKey);
      projectile.rows += 1;
      projectile.totalBalance += balance;
    }

    if (chargesMap.has(chargeKey)) {
      const charge = chargesMap.get(chargeKey);
      charge.rows += 1;
      charge.totalBalance += balance;
    }

    if (combinationsMap.has(combinationKey)) {
      const combination = combinationsMap.get(combinationKey);
      combination.rows += 1;
      combination.totalBalance += balance;
      combination.maxRangeKm = Math.max(combination.maxRangeKm, Number(item.rangeKm || 0));
      combination.longRange = combination.longRange || Boolean(item.longRange);
      combination.units.add(unitKey);
    }
  });

  const sortByName = (a, b) => String(a.name || a.combination).localeCompare(String(b.name || b.combination), "uk");

  return {
    units: Array.from(unitsMap.values()).sort(sortByName),
    projectiles: Array.from(projectilesMap.values()).sort(sortByName),
    charges: Array.from(chargesMap.values()).sort(sortByName),
    combinations: Array.from(combinationsMap.values()).map(item => ({
      ...item,
      unitsCount: item.units.size,
      units: Array.from(item.units).filter(Boolean).sort().join(", ")
    })).sort((a, b) => a.combination.localeCompare(b.combination, "uk"))
  };
}

function renderAutoDictionariesPanel(dicts) {
  if (!dicts) return "";

  const topCombinations = [...dicts.combinations]
    .sort((a, b) => Number(b.totalBalance || 0) - Number(a.totalBalance || 0))
    .slice(0, 12);

  return `
    <div class="dictionary-panel">
      <div class="dictionary-header">
        <div>
          <h2>Автодовідники з Excel</h2>
          <p>Система автоматично сформувала довідники з поточного файлу. Це база для майбутнього ручного довідника верхнього рівня.</p>
        </div>
        <div class="dictionary-metrics">
          <span>Підрозділів: <b>${dicts.units.length}</b></span>
          <span>Снарядів: <b>${dicts.projectiles.length}</b></span>
          <span>Зарядів: <b>${dicts.charges.length}</b></span>
          <span>Комбінацій: <b>${dicts.combinations.length}</b></span>
        </div>
      </div>

      <div class="dictionary-grid">
        <div class="dictionary-card">
          <div class="dictionary-title">Підрозділи</div>
          ${renderDictionaryList(dicts.units.map(item => `${item.name} — залишок ${item.totalBalance}`), "Підрозділів не знайдено")}
        </div>
        <div class="dictionary-card">
          <div class="dictionary-title">Снаряди</div>
          ${renderDictionaryList(dicts.projectiles.map(item => `${item.name} — ${item.rows} ряд.`), "Снарядів не знайдено")}
        </div>
        <div class="dictionary-card">
          <div class="dictionary-title">Заряди</div>
          ${renderDictionaryList(dicts.charges.map(item => `${item.name} — ${item.rows} ряд.`), "Зарядів не знайдено")}
        </div>
      </div>

      <div class="table-panel analysis-table dictionary-table-panel">
        <h2>ТОП комбінацій з автодовідника</h2>
        <div class="table-wrap compact-wrap">
          <table>
            <thead>
              <tr>
                <th>Снаряд</th>
                <th>Заряд</th>
                <th>Макс. дальність, км</th>
                <th>Далекобійна</th>
                <th>Підрозділів</th>
                <th>Залишок</th>
              </tr>
            </thead>
            <tbody>
              ${topCombinations.map(item => `
                <tr>
                  <td>${escapeHtml(item.projectile)}</td>
                  <td>${escapeHtml(item.charge)}</td>
                  <td>${item.maxRangeKm ? item.maxRangeKm.toFixed(1) : ""}</td>
                  <td>${item.longRange ? "Так" : "Ні"}</td>
                  <td>${item.unitsCount}</td>
                  <td>${item.totalBalance}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function renderDictionaryList(values, emptyText) {
  if (!values || !values.length) {
    return `<div class="dictionary-empty">${escapeHtml(emptyText)}</div>`;
  }

  return `
    <ul class="dictionary-list">
      ${values.slice(0, 18).map(value => `<li>${escapeHtml(value)}</li>`).join("")}
      ${values.length > 18 ? `<li class="dictionary-more">+${values.length - 18} ще</li>` : ""}
    </ul>
  `;
}

function renderAnalysis(allItems, unitItems, summaryItems, grouped) {
  if (!allItems.length) {
    analysisPanel.innerHTML = `
      <p class="empty-message">
        Не знайдено структурованих рядків у колонці <b>C</b>
        формату <b>Снаряд (Заряд)</b>.
      </p>
    `;
    return;
  }

  const unitTotalBalance = unitItems.reduce((sum, item) => sum + Number(item.balance || 0), 0);
  const unitLongRangeBalance = unitItems
    .filter(item => item.longRange)
    .reduce((sum, item) => sum + Number(item.balance || 0), 0);
  const unitShortRangeBalance = unitTotalBalance - unitLongRangeBalance;
  const summaryTotalBalance = summaryItems.reduce((sum, item) => sum + Number(item.balance || 0), 0);
  const uniqueCombinations = new Set(unitItems.map(item => item.combination)).size;
  const unitCount = Object.keys(grouped).length;
  const groupedByCategory = groupByCategory(unitItems);
  const groupedByRangeBand = groupByRangeBand(unitItems);
  const criticalItems = getCriticalItems(unitItems);
  const lowThreshold = getLowBalanceThreshold();
  const lowBalanceItems = getLowBalanceItems(unitItems, lowThreshold);
  const topBalanceItems = getTopBalanceItems(unitItems, 5);
  const topLongRangeItems = getTopLongRangeItems(unitItems, 5);
  const topRiskItems = getTopRiskItems(unitItems, lowThreshold, 5);
  const topUnitsByBalance = getUnitRanking(grouped, "balance", 6);
  const topUnitsByLongRange = getUnitRanking(grouped, "longRange", 6);
  const topUnitsByRisk = getUnitRanking(grouped, "risk", 6);
  const recommendations = getRecommendations(unitItems, grouped);
  const commanderSummary = getCommanderSummary(unitItems, grouped);
  const reportPassport = getReportPassport(unitItems, grouped);
  const exchangeBaseItems = window.ArtAmmoState?.unitItems || unitItems;
  const exchangeRecommendations = getFilteredExchangeRecommendations(exchangeBaseItems, getLowBalanceThreshold(), 12);
  const exchangeActionPlan = getFilteredExchangeActionPlan(exchangeRecommendations);
  const actionStatusIntegrity = getActionStatusIntegrity(exchangeActionPlan);
  const dataQuality = getDataQualityReport(unitItems, summaryItems, grouped);
  const autoDictionaries = buildAutoDictionaries(window.ArtAmmoState?.unitItems || unitItems);

  window.ArtAmmoState.dataQuality = dataQuality;
  window.ArtAmmoState.autoDictionaries = autoDictionaries;

  let html = `
    <div class="analysis-grid">
      <div class="metric-card">
        <div class="metric-label">Підрозділів</div>
        <div class="metric-value">${unitCount}</div>
      </div>
      <div class="metric-card">
        <div class="metric-label">Комбінацій</div>
        <div class="metric-value">${uniqueCombinations}</div>
      </div>
      <div class="metric-card">
        <div class="metric-label">Залишок</div>
        <div class="metric-value">${unitTotalBalance}</div>
      </div>
      <div class="metric-card">
        <div class="metric-label">Далекобійних</div>
        <div class="metric-value">${unitLongRangeBalance}</div>
      </div>
      <div class="metric-card">
        <div class="metric-label">Недалекобійних</div>
        <div class="metric-value">${unitShortRangeBalance}</div>
      </div>
      <div class="metric-card">
        <div class="metric-label">Нульових позицій</div>
        <div class="metric-value">${criticalItems.length}</div>
      </div>
      <div class="metric-card">
        <div class="metric-label">Малий залишок ≤${lowThreshold}</div>
        <div class="metric-value">${lowBalanceItems.length}</div>
      </div>
    </div>

    <div class="insights-grid">
      <div class="insight-card">
        <div class="insight-title">ТОП залишків</div>
        ${renderInsightList(topBalanceItems, "Позицій із залишком не знайдено")}
      </div>

      <div class="insight-card">
        <div class="insight-title">ТОП далекобійних</div>
        ${renderInsightList(topLongRangeItems, "Далекобійних залишків не знайдено")}
      </div>

      <div class="insight-card">
        <div class="insight-title">Критичний ризик</div>
        ${renderInsightList(topRiskItems, "Критичних позицій не знайдено")}
      </div>
    </div>

    <div class="unit-rank-grid">
      <div class="unit-rank-card">
        <div class="insight-title">Рейтинг підрозділів за залишком</div>
        ${renderUnitRankList(topUnitsByBalance, "totalBalance")}
      </div>

      <div class="unit-rank-card">
        <div class="insight-title">Рейтинг за далекобійними</div>
        ${renderUnitRankList(topUnitsByLongRange, "longRangeBalance")}
      </div>

      <div class="unit-rank-card">
        <div class="insight-title">Рейтинг ризику</div>
        ${renderUnitRankList(topUnitsByRisk, "riskCount", " поз.")}
      </div>
    </div>

    ${renderReportPassport(reportPassport)}
    ${renderDataQualityPanel(dataQuality)}
    ${renderAutoDictionariesPanel(autoDictionaries)}
    ${renderComparisonPanel(window.ArtAmmoState?.unitItems || unitItems)}
    ${renderExchangeRecommendations(exchangeRecommendations)}
    ${renderExchangeImpact(exchangeRecommendations)}
    ${renderExchangeActionPlan(exchangeActionPlan)}
    ${renderActionStatusIntegrity(actionStatusIntegrity)}
    ${renderRecommendations(recommendations)}
    ${renderCommanderSummary(commanderSummary)}
  `;

  if (summaryItems && summaryItems.length) {
    html += `
      <div class="table-panel analysis-table">
        <h2>Контрольна звірка</h2>
        <div class="table-wrap compact-wrap">
          <table>
            <thead>
              <tr><th>Показник</th><th>Значення</th></tr>
            </thead>
            <tbody>
              <tr><td>Залишок по аркушах підрозділів</td><td>${unitTotalBalance}</td></tr>
              <tr><td>Залишок по зведеному аркушу ${ART_AMMO_SCHEMA.SUMMARY_SHEET_NAME}</td><td>${summaryTotalBalance}</td></tr>
              <tr><td>Різниця</td><td>${unitTotalBalance - summaryTotalBalance}</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  html += `
    <div class="table-panel analysis-table">
      <h2>Підсумок по підрозділах</h2>
      <div class="table-wrap compact-wrap">
        <table>
          <thead>
            <tr>
              <th>Підрозділ</th>
              <th>Комбінацій</th>
              <th>Отримання</th>
              <th>Витрата</th>
              <th>Залишок</th>
              <th>Далекобійних</th>
            </tr>
          </thead>
          <tbody>
  `;

  Object.values(grouped).forEach((unitData) => {
    html += `
      <tr>
        <td>${escapeHtml(unitData.unit)}</td>
        <td>${unitData.combinations.size}</td>
        <td>${unitData.totalReceived}</td>
        <td>${unitData.totalSpent}</td>
        <td>${unitData.totalBalance}</td>
        <td>${unitData.longRangeBalance}</td>
      </tr>
    `;
  });

  html += `
          </tbody>
        </table>
      </div>
    </div>

    <div class="table-panel analysis-table">
      <h2>Підсумок по категоріях</h2>
      <div class="table-wrap compact-wrap">
        <table>
          <thead>
            <tr>
              <th>Категорія</th>
              <th>Комбінацій</th>
              <th>Отримання</th>
              <th>Витрата</th>
              <th>Залишок</th>
              <th>Далекобійних</th>
            </tr>
          </thead>
          <tbody>
  `;

  Object.values(groupedByCategory).forEach((categoryData) => {
    html += `
      <tr>
        <td>${escapeHtml(categoryData.category)}</td>
        <td>${categoryData.combinations.size}</td>
        <td>${categoryData.totalReceived}</td>
        <td>${categoryData.totalSpent}</td>
        <td>${categoryData.totalBalance}</td>
        <td>${categoryData.longRangeBalance}</td>
      </tr>
    `;
  });

  html += `
          </tbody>
        </table>
      </div>
    </div>

    <div class="table-panel analysis-table">
      <h2>Розподіл по дальності</h2>
      <div class="table-wrap compact-wrap">
        <table>
          <thead>
            <tr>
              <th>Діапазон</th>
              <th>Комбінацій</th>
              <th>Рядків</th>
              <th>Залишок</th>
            </tr>
          </thead>
          <tbody>
  `;

  Object.values(groupedByRangeBand).forEach((bandData) => {
    html += `
      <tr>
        <td>${escapeHtml(bandData.band)}</td>
        <td>${bandData.combinations.size}</td>
        <td>${bandData.items.length}</td>
        <td>${bandData.totalBalance}</td>
      </tr>
    `;
  });

  html += `
          </tbody>
        </table>
      </div>
    </div>

    <div class="table-panel analysis-table">
      <h2>Вузькі місця</h2>
      <div class="table-wrap compact-wrap">
        <table>
          <thead>
            <tr>
              <th>Статус</th>
              <th>Підрозділ</th>
              <th>Категорія</th>
              <th>Снаряд</th>
              <th>Заряд</th>
              <th>Дальність, км</th>
              <th>Далекобійна</th>
              <th>Залишок</th>
            </tr>
          </thead>
          <tbody>
  `;

  [...criticalItems, ...lowBalanceItems].slice(0, 80).forEach((item) => {
    const status = Number(item.balance || 0) === 0 ? "Нуль" : "Мало";

    html += `
      <tr class="${getRowClass(item)}">
        <td><span class="status-badge ${Number(item.balance || 0) === 0 ? "status-zero" : "status-low"}">${status}</span></td>
        <td>${escapeHtml(item.unit)}</td>
        <td>${escapeHtml(item.category)}</td>
        <td>${escapeHtml(item.projectile)}</td>
        <td>${escapeHtml(item.charge)}</td>
        <td>${item.rangeKm ? item.rangeKm.toFixed(1) : ""}</td>
        <td>${item.longRange ? "Так" : "Ні"}</td>
        <td>${item.balance}</td>
      </tr>
    `;
  });

  if (![...criticalItems, ...lowBalanceItems].length) {
    html += `
      <tr>
        <td colspan="8">Критичних позицій за поточними фільтрами не знайдено.</td>
      </tr>
    `;
  }

  html += `
          </tbody>
        </table>
      </div>
    </div>

    <div class="table-panel analysis-table">
      <h2>Детальний аналіз по підрозділах</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Підрозділ</th>
              <th>Рядок</th>
              <th>Категорія</th>
              <th>Снаряд</th>
              <th>Заряд</th>
              <th>Примітка</th>
              <th>Дальність, м</th>
              <th>Дальність, км</th>
              <th>Далекобійна</th>
              <th>Отримання</th>
              <th>Витрата</th>
              <th>Залишок</th>
            </tr>
          </thead>
          <tbody>
  `;

  unitItems.forEach((item) => {
    html += `
      <tr class="${getRowClass(item)}">
        <td>${escapeHtml(item.unit)}</td>
        <td>${item.row}</td>
        <td>${escapeHtml(item.category)}</td>
        <td>${escapeHtml(item.projectile)}</td>
        <td>${escapeHtml(item.charge)}</td>
        <td>${escapeHtml(item.note)}</td>
        <td>${item.rangeMeters || ""}</td>
        <td>${item.rangeKm ? item.rangeKm.toFixed(1) : ""}</td>
        <td>${item.longRange ? "Так" : "Ні"}</td>
        <td>${item.received}</td>
        <td>${item.spent}</td>
        <td>${item.balance}</td>
      </tr>
    `;
  });

  html += `
          </tbody>
        </table>
      </div>
    </div>
  `;

  analysisPanel.innerHTML = html;
}


function getExchangeRecommendations(items, threshold = 10, limit = 12) {
  const activeItems = [...(items || [])];
  const byCombination = {};

  activeItems.forEach(item => {
    const key = item.combination || `${item.projectile} + ${item.charge}`;

    if (!byCombination[key]) {
      byCombination[key] = {
        key,
        projectile: item.projectile,
        charge: item.charge,
        rangeKm: item.rangeKm,
        longRange: item.longRange,
        units: []
      };
    }

    byCombination[key].units.push({
      unit: item.unit,
      balance: Number(item.balance || 0),
      item
    });
  });

  const recommendations = [];

  Object.values(byCombination).forEach(group => {
    const receivers = group.units
      .filter(row => row.balance >= 0 && row.balance <= threshold)
      .sort((a, b) => a.balance - b.balance || String(a.unit).localeCompare(String(b.unit), "uk"));

    const donors = group.units
      .filter(row => row.balance > threshold)
      .sort((a, b) => b.balance - a.balance || String(a.unit).localeCompare(String(b.unit), "uk"));

    receivers.forEach(receiver => {
      if (!donors.length) return;

      const donor = donors[0];
      if (!donor || donor.unit === receiver.unit) return;

      const targetLevel = threshold + 1;
      const receiverNeed = Math.max(1, targetLevel - receiver.balance);
      const donorSurplus = Math.max(0, donor.balance - threshold);
      const recommendedQty = Math.min(receiverNeed, donorSurplus);

      if (recommendedQty <= 0) return;

      recommendations.push({
        priority: group.longRange ? 1 : 2,
        type: receiver.balance === 0 ? "Нуль" : "Мало",
        projectile: group.projectile,
        charge: group.charge,
        combination: group.key,
        rangeKm: group.rangeKm,
        longRange: group.longRange,
        fromUnit: donor.unit,
        toUnit: receiver.unit,
        donorBalance: donor.balance,
        receiverBalance: receiver.balance,
        recommendedQty,
        expectedReceiverBalance: receiver.balance + recommendedQty,
        reason: receiver.balance === 0
          ? "закриття нульового залишку"
          : "підняття малого залишку вище порогу"
      });
    });
  });

  return recommendations
    .sort((a, b) =>
      a.priority - b.priority ||
      a.receiverBalance - b.receiverBalance ||
      b.recommendedQty - a.recommendedQty ||
      Number(b.rangeKm || 0) - Number(a.rangeKm || 0)
    )
    .slice(0, limit);
}

function renderExchangeRecommendations(recommendations) {
  if (!recommendations || !recommendations.length) {
    return `
      <div class="exchange-panel">
        <div class="exchange-header">
          <h2>Рекомендації щодо обміну</h2>
          <span>Потенційних обмінів за поточними даними не знайдено</span>
        </div>
      </div>
    `;
  }

  return `
    <div class="exchange-panel">
      <div class="exchange-header">
        <h2>Рекомендації щодо обміну</h2>
        <span>Пошук однакових комбінацій: де дефіцит в одному підрозділі і надлишок в іншому</span>
      </div>

      <div class="table-wrap compact-wrap">
        <table>
          <thead>
            <tr>
              <th>Пріоритет</th>
              <th>Комбінація</th>
              <th>Дальність</th>
              <th>Передати з</th>
              <th>Передати до</th>
              <th>Рекомендовано</th>
              <th>Пояснення</th>
            </tr>
          </thead>
          <tbody>
            ${recommendations.map(item => `
              <tr class="${item.longRange ? "row-long" : "row-normal"}">
                <td><span class="status-badge ${item.type === "Нуль" ? "status-zero" : "status-low"}">${item.type}</span></td>
                <td>${escapeHtml(item.projectile)} (${escapeHtml(item.charge)})</td>
                <td>${item.rangeKm ? item.rangeKm.toFixed(1) + " км" : ""}</td>
                <td>${escapeHtml(item.fromUnit)} <span class="exchange-muted">(${item.donorBalance})</span></td>
                <td>${escapeHtml(item.toUnit)} <span class="exchange-muted">(${item.receiverBalance})</span></td>
                <td><strong>${item.recommendedQty}</strong></td>
                <td>${escapeHtml(item.reason)} → очікувано ${item.expectedReceiverBalance}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function getExchangeImpactSummary(recommendations) {
  const threshold = getLowBalanceThreshold();
  const recs = Array.isArray(recommendations) ? recommendations : [];

  const totalRecommended = recs.reduce((sum, item) => sum + Number(item.recommendedQty || 0), 0);
  const longRangeRecommended = recs
    .filter(item => item.longRange)
    .reduce((sum, item) => sum + Number(item.recommendedQty || 0), 0);

  const zeroClosed = recs.filter(item =>
    Number(item.receiverBalance || 0) === 0 && Number(item.expectedReceiverBalance || 0) > 0
  ).length;

  const raisedAboveThreshold = recs.filter(item =>
    Number(item.receiverBalance || 0) <= threshold &&
    Number(item.expectedReceiverBalance || 0) > threshold
  ).length;

  const donorWarnings = recs.filter(item =>
    Number(item.donorBalance || 0) - Number(item.recommendedQty || 0) <= threshold
  ).length;

  return {
    count: recs.length,
    totalRecommended,
    longRangeRecommended,
    zeroClosed,
    raisedAboveThreshold,
    donorWarnings,
    threshold
  };
}

function renderExchangeImpact(recommendations) {
  if (!recommendations || !recommendations.length) return "";

  const impact = getExchangeImpactSummary(recommendations);

  return `
    <div class="exchange-impact-panel">
      <div class="exchange-impact-header">
        <h2>Очікуваний ефект обміну</h2>
        <span>Оцінка після виконання рекомендованих передач</span>
      </div>

      <div class="exchange-impact-grid">
        <div class="exchange-impact-card">
          <div class="impact-label">Рекомендацій</div>
          <div class="impact-value">${impact.count}</div>
        </div>

        <div class="exchange-impact-card">
          <div class="impact-label">Передати загалом</div>
          <div class="impact-value">${impact.totalRecommended}</div>
        </div>

        <div class="exchange-impact-card">
          <div class="impact-label">Далекобійних у передачі</div>
          <div class="impact-value">${impact.longRangeRecommended}</div>
        </div>

        <div class="exchange-impact-card">
          <div class="impact-label">Закрито нульових</div>
          <div class="impact-value">${impact.zeroClosed}</div>
        </div>

        <div class="exchange-impact-card">
          <div class="impact-label">Піднято вище порогу</div>
          <div class="impact-value">${impact.raisedAboveThreshold}</div>
        </div>

        <div class="exchange-impact-card ${impact.donorWarnings ? "impact-warning" : ""}">
          <div class="impact-label">Донорів біля порогу</div>
          <div class="impact-value">${impact.donorWarnings}</div>
        </div>
      </div>

      <div class="exchange-impact-note">
        Поріг малого залишку: ≤${impact.threshold}. Розрахунок є попереднім і не враховує фізичну доступність, час переміщення та рішення відповідальних осіб.
      </div>
    </div>
  `;
}


function getFilteredExchangeRecommendations(items, threshold = 10, limit = 12) {
  let recommendations = getExchangeRecommendations(items, threshold, limit);

  if (actionLongOnly?.checked) {
    recommendations = recommendations.filter(item => item.longRange);
  }

  return recommendations;
}

function getFilteredExchangeActionPlan(recommendations) {
  let plan = getExchangeActionPlan(recommendations);
  const priority = actionPriorityFilter?.value || "all";
  const status = actionStatusFilter?.value || "all";

  if (priority !== "all") {
    plan = plan.filter(item => item.priority === priority);
  }

  if (status !== "all") {
    plan = plan.filter(item => item.status === status);
  }

  return plan;
}

function getExchangeActionPlan(recommendations) {
  const recs = Array.isArray(recommendations) ? recommendations : [];

  return recs.map((item, index) => {
    const donorAfter = Number(item.donorBalance || 0) - Number(item.recommendedQty || 0);
    const receiverAfter = Number(item.expectedReceiverBalance || 0);
    const threshold = getLowBalanceThreshold();

    let priority = "III";
    let risk = "Планово";

    if (item.longRange && Number(item.receiverBalance || 0) === 0) {
      priority = "I";
      risk = "Критично";
    } else if (item.longRange || Number(item.receiverBalance || 0) === 0) {
      priority = "II";
      risk = "Високий";
    }

    if (donorAfter <= threshold) {
      risk += " / контроль донора";
    }

    const actionId = getActionId(item);

    return {
      order: index + 1,
      actionId,
      status: getActionStatus(actionId),
      priority,
      risk,
      action: `Передати ${item.recommendedQty} од. ${item.projectile} (${item.charge})`,
      fromUnit: item.fromUnit,
      toUnit: item.toUnit,
      before: `${item.toUnit}: ${item.receiverBalance}; ${item.fromUnit}: ${item.donorBalance}`,
      after: `${item.toUnit}: ${receiverAfter}; ${item.fromUnit}: ${donorAfter}`,
      longRange: item.longRange,
      rangeKm: item.rangeKm,
      reason: item.reason
    };
  });
}

function getActionId(item) {
  return [
    item.projectile,
    item.charge,
    item.fromUnit,
    item.toUnit,
    item.recommendedQty
  ].map(part => String(part ?? "").trim()).join("|");
}


function getAllActionStatuses() {
  try {
    return JSON.parse(localStorage.getItem("artAmmoActionStatuses") || "{}");
  } catch (error) {
    return {};
  }
}

function saveAllActionStatuses(statuses) {
  try {
    localStorage.setItem("artAmmoActionStatuses", JSON.stringify(statuses || {}));
  } catch (error) {
    console.warn("Не вдалося зберегти статуси дій", error);
  }
}

function exportActionStatusesToJson() {
  const statuses = getAllActionStatuses();
  const passport = window.ArtAmmoState?.reportPassport || {};
  const payload = {
    app: "Art Ammo",
    type: "action-statuses",
    version: 1,
    exportedAt: new Date().toISOString(),
    sourceFileName: passport.fileName || window.ArtAmmoState?.fileName || "",
    statuses
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json;charset=utf-8"
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const fileDate = new Date().toISOString().slice(0, 10);
  link.href = url;
  link.download = `art_ammo_action_statuses_${fileDate}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function importActionStatusesFromJson(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const payload = JSON.parse(String(e.target.result || "{}"));
      const incoming = payload.statuses || payload;

      if (!incoming || typeof incoming !== "object" || Array.isArray(incoming)) {
        throw new Error("Невірна структура JSON");
      }

      const allowed = new Set(["planned", "done", "rejected"]);
      const current = getAllActionStatuses();
      let importedCount = 0;

      Object.entries(incoming).forEach(([actionId, status]) => {
        if (allowed.has(String(status))) {
          current[actionId] = String(status);
          importedCount += 1;
        }
      });

      saveAllActionStatuses(current);
      alert(`Імпортовано статусів: ${importedCount}`);
      applyFilters();
    } catch (error) {
      alert("Не вдалося імпортувати статуси. Перевір JSON-файл.");
      console.error(error);
    } finally {
      event.target.value = "";
    }
  };

  reader.readAsText(file, "utf-8");
}

window.exportActionStatusesToJson = exportActionStatusesToJson;
window.importActionStatusesFromJson = importActionStatusesFromJson;

function getActionStatus(actionId) {
  const statuses = getAllActionStatuses();
  return statuses[actionId] || "planned";
}

function setActionStatus(actionId, status) {
  const statuses = getAllActionStatuses();
  statuses[actionId] = status;
  saveAllActionStatuses(statuses);
}

function updateActionStatus(actionId, status) {
  setActionStatus(actionId, status);
  applyFilters();
}

window.updateActionStatus = updateActionStatus;

function renderExchangeActionPlan(plan) {
  if (!plan || !plan.length) return "";

  return `
    <div class="action-plan-panel">
      <div class="action-plan-header">
        <h2>Журнал рекомендованих дій</h2>
        <span>Практичний список кроків за результатами обміну · ${escapeHtml(getActionFilterDescription())}</span>
      </div>

      <div class="table-wrap compact-wrap">
        <table>
          <thead>
            <tr>
              <th>№</th>
              <th>Пріоритет</th>
              <th>Дія</th>
              <th>Звідки</th>
              <th>Куди</th>
              <th>Було</th>
              <th>Буде</th>
              <th>Ризик</th>
              <th>Статус</th>
            </tr>
          </thead>
          <tbody>
            ${plan.map(item => `
              <tr class="${item.priority === "I" ? "row-zero" : item.priority === "II" ? "row-low" : "row-normal"}">
                <td>${item.order}</td>
                <td><span class="action-priority action-priority-${item.priority}">${item.priority}</span></td>
                <td>${escapeHtml(item.action)}${item.longRange ? ` <span class="status-badge status-long">Далекобійна</span>` : ""}</td>
                <td>${escapeHtml(item.fromUnit)}</td>
                <td>${escapeHtml(item.toUnit)}</td>
                <td>${escapeHtml(item.before)}</td>
                <td>${escapeHtml(item.after)}</td>
                <td>${escapeHtml(item.risk)}</td>
                <td>
                  <select class="action-status-select status-${item.status}" data-action-id="${escapeHtml(item.actionId)}" onchange="updateActionStatus(this.dataset.actionId, this.value)">
                    <option value="planned" ${item.status === "planned" ? "selected" : ""}>Заплановано</option>
                    <option value="done" ${item.status === "done" ? "selected" : ""}>Виконано</option>
                    <option value="rejected" ${item.status === "rejected" ? "selected" : ""}>Відхилено</option>
                  </select>
                </td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>

      <div class="action-plan-note">
        Це попередній журнал дій для планування. Перед виконанням потрібно підтвердити фізичну наявність, доступність транспорту, час переміщення та рішення відповідальних осіб.
      </div>
    </div>
  `;
}


function buildDecisionPackagePayload() {
  const items = typeof getCurrentFilteredItems === "function"
    ? getCurrentFilteredItems()
    : (window.ArtAmmoState?.unitItems || []);

  const grouped = groupByUnit(items);
  const threshold = getLowBalanceThreshold();
  const passport = window.ArtAmmoState?.reportPassport || getReportPassport(items);
  const commanderSummary = getCommanderSummary(items, grouped);
  const recommendations = getRecommendations(items, grouped);
  const exchangeRecommendations = getFilteredExchangeRecommendations(
    window.ArtAmmoState?.unitItems || items,
    threshold,
    50
  );
  const actionPlan = getFilteredExchangeActionPlan(exchangeRecommendations);
  const actionIntegrity = getActionStatusIntegrity(actionPlan);
  const dataQuality = getDataQualityReport(items, [], grouped);

  const totals = {
    units: Object.keys(grouped).length,
    rows: items.length,
    combinations: new Set(items.map(item => item.combination)).size,
    received: items.reduce((sum, item) => sum + Number(item.received || 0), 0),
    spent: items.reduce((sum, item) => sum + Number(item.spent || 0), 0),
    balance: items.reduce((sum, item) => sum + Number(item.balance || 0), 0),
    longRangeBalance: items.filter(item => item.longRange).reduce((sum, item) => sum + Number(item.balance || 0), 0),
    zeroPositions: items.filter(item => Number(item.balance || 0) === 0).length,
    lowPositions: items.filter(item => Number(item.balance || 0) > 0 && Number(item.balance || 0) <= threshold).length,
    lowBalanceThreshold: threshold
  };

  const meta = getAppMeta();

  return {
    app: meta.appName || "Art Ammo",
    type: "decision-package",
    packageVersion: 2,
    appVersion: meta.version,
    buildLabel: meta.buildLabel,
    buildDate: meta.buildDate,
    logicProfile: meta.logicProfile,
    exportedAt: new Date().toISOString(),
    passport,
    totals,
    commanderSummary,
    recommendations,
    exchangeRecommendations,
    actionPlan,
    actionStatusIntegrity: actionIntegrity,
    dataQuality,
    actionStatuses: getAllActionStatuses(),
    filteredRows: items.map(item => ({
      unit: item.unit,
      row: item.row,
      category: item.category,
      projectile: item.projectile,
      charge: item.charge,
      note: item.note,
      rangeMeters: item.rangeMeters,
      rangeKm: item.rangeKm,
      longRange: item.longRange,
      received: item.received,
      spent: item.spent,
      balance: item.balance,
      combination: item.combination
    }))
  };
}


function getImportedDecisionPackageInfo() {
  const payload = window.ArtAmmoState?.importedDecisionPackage;

  if (!payload) return null;

  const exportedAt = payload.exportedAt
    ? new Date(payload.exportedAt).toLocaleString("uk-UA")
    : "—";

  const statusesCount = payload.actionStatuses
    ? Object.keys(payload.actionStatuses).length
    : payload.statuses
      ? Object.keys(payload.statuses).length
      : 0;

  return {
    type: payload.type || "невідомий тип",
    exportedAt,
    sourceFileName: payload.passport?.fileName || payload.sourceFileName || "—",
    actionCount: Array.isArray(payload.actionPlan) ? payload.actionPlan.length : 0,
    statusesCount
  };
}

function renderImportedDecisionPackageInfo() {
  const info = getImportedDecisionPackageInfo();

  if (!info) return "";

  return `
    <div class="decision-import-panel">
      <div class="decision-import-header">
        <h2>Імпортований пакет рішення</h2>
        <span>${escapeHtml(info.type)}</span>
      </div>

      <div class="decision-import-grid">
        <div class="decision-import-card">
          <div class="impact-label">Файл-джерело</div>
          <div class="impact-value small-text">${escapeHtml(info.sourceFileName)}</div>
        </div>

        <div class="decision-import-card">
          <div class="impact-label">Сформовано</div>
          <div class="impact-value small-text">${escapeHtml(info.exportedAt)}</div>
        </div>

        <div class="decision-import-card">
          <div class="impact-label">Дій у пакеті</div>
          <div class="impact-value">${info.actionCount}</div>
        </div>

        <div class="decision-import-card">
          <div class="impact-label">Імпортовано статусів</div>
          <div class="impact-value">${info.statusesCount}</div>
        </div>
      </div>
    </div>
  `;
}

async function readDecisionPackagePayload(file) {
  const name = String(file?.name || "").toLowerCase();

  if (name.endsWith(".zip")) {
    if (!window.JSZip) {
      throw new Error("JSZip не завантажено. Неможливо прочитати ZIP-пакет.");
    }

    const zip = await JSZip.loadAsync(file);
    const preferred = zip.file("decision_package.json") || zip.file("action_statuses.json");

    if (!preferred) {
      throw new Error("У ZIP не знайдено decision_package.json або action_statuses.json");
    }

    const text = await preferred.async("string");
    return JSON.parse(text);
  }

  const text = await file.text();
  return JSON.parse(text);
}

function extractStatusesFromDecisionPayload(payload) {
  if (!payload || typeof payload !== "object") return {};

  if (payload.type === "decision-package") {
    return payload.actionStatuses || {};
  }

  if (payload.type === "action-statuses") {
    return payload.statuses || {};
  }

  return payload.statuses || payload.actionStatuses || payload;
}

function mergeImportedActionStatuses(incoming) {
  const allowed = new Set(["planned", "done", "rejected"]);
  const current = getAllActionStatuses();
  let importedCount = 0;

  Object.entries(incoming || {}).forEach(([actionId, status]) => {
    if (allowed.has(String(status))) {
      current[actionId] = String(status);
      importedCount += 1;
    }
  });

  saveAllActionStatuses(current);
  return importedCount;
}

async function importDecisionPackageFromFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    const payload = await readDecisionPackagePayload(file);
    const statuses = extractStatusesFromDecisionPayload(payload);
    const importedCount = mergeImportedActionStatuses(statuses);

    window.ArtAmmoState.importedDecisionPackage = payload;
    window.ArtAmmoState.lastDecisionPackageImport = {
      fileName: file.name,
      importedAt: new Date().toISOString(),
      importedCount
    };

    if (typeof setStatus === "function") {
      setStatus(`Імпортовано пакет: ${importedCount} статусів`, "ok");
    }

    alert(`Пакет імпортовано. Статусів: ${importedCount}`);
    applyFilters();
  } catch (error) {
    alert("Не вдалося імпортувати пакет рішення. Перевір JSON/ZIP-файл.");
    console.error(error);
  } finally {
    event.target.value = "";
  }
}

window.importDecisionPackageFromFile = importDecisionPackageFromFile;

function decisionPackageText(payload) {
  const lines = [];

  lines.push("ART AMMO — ПАКЕТ РІШЕННЯ");
  lines.push("================================");
  lines.push(`Версія системи: ${payload.app || "Art Ammo"} ${payload.appVersion || "—"}`);
  lines.push(`Збірка: ${payload.buildLabel || "—"} / ${payload.buildDate || "—"}`);
  lines.push(`Профіль логіки: ${payload.logicProfile || "—"}`);
  lines.push(`Сформовано: ${new Date(payload.exportedAt).toLocaleString("uk-UA")}`);
  lines.push(`Файл: ${payload.passport?.fileName || "—"}`);
  lines.push(`Активні фільтри: ${payload.passport?.filters || "—"}`);
  lines.push(`Поріг малого залишку: ${payload.totals.lowBalanceThreshold}`);
  lines.push("");
  lines.push("КЛЮЧОВІ ПОКАЗНИКИ");
  lines.push(`Підрозділів: ${payload.totals.units}`);
  lines.push(`Рядків: ${payload.totals.rows}`);
  lines.push(`Комбінацій: ${payload.totals.combinations}`);
  lines.push(`Отримання: ${payload.totals.received}`);
  lines.push(`Витрата: ${payload.totals.spent}`);
  lines.push(`Залишок: ${payload.totals.balance}`);
  lines.push(`Далекобійних: ${payload.totals.longRangeBalance}`);
  lines.push(`Нульових позицій: ${payload.totals.zeroPositions}`);
  lines.push(`Малих залишків: ${payload.totals.lowPositions}`);
  lines.push("");
  lines.push("КОРОТКИЙ КОМАНДИРСЬКИЙ ВИСНОВОК");
  (payload.commanderSummary || []).forEach(item => {
    lines.push(`- ${item.label}: ${item.text}`);
  });
  lines.push("");
  lines.push("ЖУРНАЛ РЕКОМЕНДОВАНИХ ДІЙ");
  if (!payload.actionPlan?.length) {
    lines.push("Рекомендовані дії відсутні за поточними фільтрами.");
  } else {
    payload.actionPlan.forEach(item => {
      lines.push(`${item.order}. [${item.priority}] ${item.action} | ${item.fromUnit} → ${item.toUnit} | ${item.statusLabel || item.status}`);
      lines.push(`   Було: ${item.before}; буде: ${item.after}; ризик: ${item.risk}`);
    });
  }

  return lines.join("\n");
}

function toCsv(rows, columns) {
  const escape = value => `"${String(value ?? "").replace(/"/g, '""')}"`;
  return [
    columns.map(col => escape(col.label)).join(";"),
    ...rows.map(row => columns.map(col => escape(row[col.key])).join(";"))
  ].join("\n");
}

async function exportDecisionPackage() {
  const items = typeof getCurrentFilteredItems === "function"
    ? getCurrentFilteredItems()
    : (window.ArtAmmoState?.unitItems || []);

  if (!items.length) {
    alert("Немає даних для пакета рішення. Спочатку завантаж і проаналізуй Excel.");
    return;
  }

  const payload = buildDecisionPackagePayload();
  const fileDate = new Date().toISOString().slice(0, 10);
  const baseName = `art_ammo_decision_package_${fileDate}`;

  const actionCsv = toCsv(payload.actionPlan || [], [
    { key: "order", label: "№" },
    { key: "priority", label: "Пріоритет" },
    { key: "action", label: "Дія" },
    { key: "fromUnit", label: "Звідки" },
    { key: "toUnit", label: "Куди" },
    { key: "before", label: "Було" },
    { key: "after", label: "Буде" },
    { key: "risk", label: "Ризик" },
    { key: "status", label: "Статус" }
  ]);

  if (window.JSZip) {
    const zip = new JSZip();
    zip.file("decision_package.json", JSON.stringify(payload, null, 2));
    zip.file("commander_summary.txt", decisionPackageText(payload));
    zip.file("action_log.csv", "\ufeff" + actionCsv);
    zip.file("action_statuses.json", JSON.stringify({
      app: payload.app || "Art Ammo",
      type: "action-statuses",
      appVersion: payload.appVersion,
      buildLabel: payload.buildLabel,
      buildDate: payload.buildDate,
      exportedAt: payload.exportedAt,
      statuses: payload.actionStatuses
    }, null, 2));

    const blob = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${baseName}.zip`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    return;
  }

  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json;charset=utf-8"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${baseName}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

window.exportDecisionPackage = exportDecisionPackage;

function getRowClass(item) {
  const balance = Number(item?.balance || 0);
  const threshold = getLowBalanceThreshold();

  if (balance === 0) return "row-zero";
  if (balance > 0 && balance <= threshold) return "row-low";
  if (item?.longRange) return "row-long";

  return "row-normal";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
