const analyzeBtn = document.getElementById("analyzeBtn");
const analysisPanel = document.getElementById("analysisPanel");
const filtersPanel = document.getElementById("filtersPanel");
const unitFilter = document.getElementById("unitFilter");
const longRangeOnly = document.getElementById("longRangeOnly");
const searchFilter = document.getElementById("searchFilter");
const resetFiltersBtn = document.getElementById("resetFiltersBtn");
const filterStatus = document.getElementById("filterStatus");
const lowBalanceThreshold = document.getElementById("lowBalanceThreshold");
const stockFilter = document.getElementById("stockFilter");
const sortFilter = document.getElementById("sortFilter");

analyzeBtn.addEventListener("click", analyzeWorkbook);
unitFilter.addEventListener("change", applyFilters);
longRangeOnly.addEventListener("change", applyFilters);
searchFilter.addEventListener("input", applyFilters);
resetFiltersBtn.addEventListener("click", resetFilters);
lowBalanceThreshold.addEventListener("input", applyFilters);
stockFilter.addEventListener("change", applyFilters);
sortFilter.addEventListener("change", applyFilters);

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

  return active.length ? active.join("; ") : "без додаткових фільтрів";
}

function getReportPassport(items, grouped) {
  const totalRows = window.ArtAmmoState?.unitItems?.length || 0;
  const fileName = window.ArtAmmoState?.fileName || "—";
  const loadedAt = formatDateTime(window.ArtAmmoState?.loadedAt);
  const analyzedAt = formatDateTime(window.ArtAmmoState?.analyzedAt);
  const threshold = getLowBalanceThreshold();

  return [
    { label: "Файл Excel", value: fileName },
    { label: "Час завантаження", value: loadedAt },
    { label: "Час аналізу", value: analyzedAt },
    { label: "Показано рядків", value: `${items.length} з ${totalRows}` },
    { label: "Підрозділів у вибірці", value: String(Object.keys(grouped || {}).length) },
    { label: "Активні фільтри", value: getActiveFilterDescription() },
    { label: "Поріг малого залишку", value: `≤${threshold}` },
    { label: "Фільтр залишків", value: getStockFilterLabel() },
    { label: "Сортування", value: getSortLabel() }
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
    ${renderComparisonPanel(window.ArtAmmoState?.unitItems || unitItems)}
    ${renderExchangeRecommendations(getExchangeRecommendations(window.ArtAmmoState?.unitItems || unitItems, getLowBalanceThreshold(), 12))}
    ${renderExchangeImpact(getExchangeRecommendations(window.ArtAmmoState?.unitItems || unitItems, getLowBalanceThreshold(), 12))}
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
