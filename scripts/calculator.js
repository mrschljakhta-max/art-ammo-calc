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

  filtersPanel.hidden = false;
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
