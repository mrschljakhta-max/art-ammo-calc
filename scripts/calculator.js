const analyzeBtn = document.getElementById("analyzeBtn");
const analysisPanel = document.getElementById("analysisPanel");

analyzeBtn.addEventListener("click", analyzeWorkbook);

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

    const parsedItems = parseStructuredSheet(sheetName, rows);
    allItems.push(...parsedItems);
  });

  const unitItems = allItems.filter(item => !item.isSummarySheet);
  const summaryItems = allItems.filter(item => item.isSummarySheet);

  const grouped = groupByUnit(unitItems);

  window.ArtAmmoState.analysisItems = allItems;
  window.ArtAmmoState.unitItems = unitItems;
  window.ArtAmmoState.summaryItems = summaryItems;
  window.ArtAmmoState.groupedByUnit = grouped;

  document.getElementById("exportExcelBtn").disabled = false;
  document.getElementById("exportPdfBtn").disabled = false;

  renderAnalysis(allItems, unitItems, summaryItems, grouped);
}

function parseStructuredSheet(sheetName, rows) {
  const items = [];

  const isSummarySheet = sheetName.trim().toLowerCase() === "аг3+ар";

  rows.forEach((row, rowIndex) => {
    if (rowIndex < 5) return;

    const category = cleanCell(row[0]);
    const ammoText = cleanCell(row[2]);
    const rangeMeters = toNumber(row[3]);
    const received = toNumber(row[4]);
    const spent = toNumber(row[5]);
    const balance = toNumber(row[6]);

    if (!ammoText) return;

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

      longRange: rangeKm >= 18,

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
    projectile: match[1].trim(),
    charge: match[2].trim(),
    note: cleanCell(match[3])
  };
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

function renderAnalysis(allItems, unitItems, summaryItems, grouped) {
  if (!allItems.length) {
    analysisPanel.innerHTML = `
      <p>
        Не знайдено структурованих рядків у колонці <b>C</b>
        формату <b>Снаряд (Заряд)</b>.
      </p>
    `;
    return;
  }

  const unitTotalBalance = unitItems.reduce(
    (sum, item) => sum + Number(item.balance || 0),
    0
  );

  const unitLongRangeBalance = unitItems
    .filter(item => item.longRange)
    .reduce((sum, item) => sum + Number(item.balance || 0), 0);

  const summaryTotalBalance = summaryItems.reduce(
    (sum, item) => sum + Number(item.balance || 0),
    0
  );

  const uniqueCombinations = new Set(
    unitItems.map(item => item.combination)
  ).size;

  const unitCount = Object.keys(grouped).length;

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
        <div class="metric-label">Залишок по підрозділах</div>
        <div class="metric-value">${unitTotalBalance}</div>
      </div>

      <div class="metric-card">
        <div class="metric-label">Далекобійних</div>
        <div class="metric-value">${unitLongRangeBalance}</div>
      </div>

    </div>

    <div class="table-panel analysis-table">
      <h2>Контрольна звірка</h2>

      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Показник</th>
              <th>Значення</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Залишок по аркушах підрозділів</td>
              <td>${unitTotalBalance}</td>
            </tr>
            <tr>
              <td>Залишок по зведеному аркушу АГ3+АР</td>
              <td>${summaryTotalBalance}</td>
            </tr>
            <tr>
              <td>Різниця</td>
              <td>${unitTotalBalance - summaryTotalBalance}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="table-panel analysis-table">
      <h2>Підсумок по підрозділах</h2>

      <div class="table-wrap">
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
        <td>${unitData.unit}</td>
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
      <tr>
        <td>${item.unit}</td>
        <td>${item.row}</td>
        <td>${item.category}</td>
        <td>${item.projectile}</td>
        <td>${item.charge}</td>
        <td>${item.note}</td>
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
