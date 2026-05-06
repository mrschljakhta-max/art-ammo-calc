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

    const parsedItems = parseUnitSheet(sheetName, rows);
    allItems.push(...parsedItems);
  });

  const grouped = groupByUnit(allItems);

  window.ArtAmmoState.analysisItems = allItems;
  window.ArtAmmoState.groupedByUnit = grouped;

  document.getElementById("exportExcelBtn").disabled = false;
  document.getElementById("exportPdfBtn").disabled = false;

  renderAnalysis(allItems, grouped);
}

function parseUnitSheet(sheetName, rows) {
  const items = [];

  rows.forEach((row, rowIndex) => {
    const cells = row.map(cell => String(cell || "").trim());

    cells.forEach((cell) => {
      const parsed = parseProjectileCharge(cell);

      if (!parsed) return;

      const range = findRange(cells);
      const quantity = findQuantity(cells);

      items.push({
        unit: sheetName,
        row: rowIndex + 1,
        projectile: parsed.projectile,
        charge: parsed.charge,
        combination: `${parsed.projectile} + ${parsed.charge}`,
        range,
        longRange: range >= 18,
        quantity
      });
    });
  });

  return items;
}

function parseProjectileCharge(text) {
  const match = String(text || "").trim().match(/^(.+?)\s*\((.+?)\)$/);

  if (!match) return null;

  return {
    projectile: match[1].trim(),
    charge: match[2].trim()
  };
}

function findRange(cells) {
  for (const cell of cells) {
    const text = String(cell || "")
      .replace(",", ".")
      .trim();

    const kmMatch = text.match(/(\d+(\.\d+)?)\s*км/i);

    if (kmMatch) {
      return Number(kmMatch[1]);
    }

    const number = Number(text);

    if (!Number.isNaN(number) && number > 0 && number <= 70) {
      return number;
    }
  }

  return null;
}

function findQuantity(cells) {
  let lastNumber = 0;

  cells.forEach((cell) => {
    const value = Number(String(cell).replace(",", "."));

    if (!Number.isNaN(value)) {
      lastNumber = value;
    }
  });

  return lastNumber;
}

function groupByUnit(items) {
  const grouped = {};

  items.forEach((item) => {
    if (!grouped[item.unit]) {
      grouped[item.unit] = {
        unit: item.unit,
        items: [],
        totalQuantity: 0,
        longRangeQuantity: 0,
        combinations: new Set()
      };
    }

    grouped[item.unit].items.push(item);
    grouped[item.unit].totalQuantity += Number(item.quantity || 0);

    if (item.longRange) {
      grouped[item.unit].longRangeQuantity += Number(item.quantity || 0);
    }

    grouped[item.unit].combinations.add(item.combination);
  });

  return grouped;
}

function renderAnalysis(items, grouped) {
  if (!items.length) {
    analysisPanel.innerHTML = `
      <p>
        Не знайдено комбінацій типу:
        <b>M483 (M3A1)</b>
      </p>
    `;
    return;
  }

  const totalQuantity = items.reduce(
    (sum, item) => sum + Number(item.quantity || 0),
    0
  );

  const longRangeQuantity = items
    .filter(item => item.longRange)
    .reduce((sum, item) => sum + Number(item.quantity || 0), 0);

  const uniqueCombinations = new Set(
    items.map(item => item.combination)
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
        <div class="metric-label">Загальний залишок</div>
        <div class="metric-value">${totalQuantity}</div>
      </div>

      <div class="metric-card">
        <div class="metric-label">Далекобійних</div>
        <div class="metric-value">${longRangeQuantity}</div>
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
              <th>Загальний залишок</th>
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
        <td>${unitData.totalQuantity}</td>
        <td>${unitData.longRangeQuantity}</td>
      </tr>
    `;
  });

  html += `
          </tbody>
        </table>
      </div>
    </div>

    <div class="table-panel analysis-table">
      <h2>Детальний аналіз</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Підрозділ</th>
              <th>Рядок</th>
              <th>Снаряд</th>
              <th>Заряд</th>
              <th>Комбінація</th>
              <th>Дальність</th>
              <th>Далекобійна</th>
              <th>Кількість</th>
            </tr>
          </thead>
          <tbody>
  `;

  items.forEach((item) => {
    html += `
      <tr>
        <td>${item.unit}</td>
        <td>${item.row}</td>
        <td>${item.projectile}</td>
        <td>${item.charge}</td>
        <td>${item.combination}</td>
        <td>${item.range ?? ""}</td>
        <td>${item.longRange ? "Так" : "Ні"}</td>
        <td>${item.quantity}</td>
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
