const analyzeBtn = document.getElementById("analyzeBtn");
const analysisPanel = document.getElementById("analysisPanel");

analyzeBtn.addEventListener("click", analyzeWorkbook);

function analyzeWorkbook() {
  const workbook = window.ArtAmmoState?.workbook;

  if (!workbook) {
    analysisPanel.innerHTML = `<p>Спочатку завантаж Excel-файл.</p>`;
    return;
  }

  const result = [];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: ""
    });

    const parsed = parseBalanceSheet(sheetName, rows);

    result.push(...parsed);
  });

  renderAnalysis(result);
}

function parseBalanceSheet(sheetName, rows) {
  const items = [];

  rows.forEach((row) => {
    const cells = row.map(value => String(value || "").trim());

    cells.forEach((cell, index) => {
      const match = cell.match(/^(.+?)\s*\((.+?)\)$/);

      if (!match) return;

      const projectile = match[1].trim();
      const charge = match[2].trim();

      const range = findRangeInRow(cells);
      const quantity = findLastNumberInRow(cells);

      items.push({
        sheetName,
        projectile,
        charge,
        range,
        longRange: Number(range) >= 18,
        quantity
      });
    });
  });

  return items;
}

function findRangeInRow(cells) {
  for (const cell of cells) {
    const text = String(cell || "").replace(",", ".");

    const kmMatch = text.match(/(\d+(\.\d+)?)\s*км/i);
    if (kmMatch) return Number(kmMatch[1]);

    const plainNumber = Number(text);
    if (!Number.isNaN(plainNumber) && plainNumber > 0 && plainNumber <= 70) {
      return plainNumber;
    }
  }

  return null;
}

function findLastNumberInRow(cells) {
  let last = 0;

  cells.forEach((cell) => {
    const value = Number(String(cell).replace(",", "."));

    if (!Number.isNaN(value)) {
      last = value;
    }
  });

  return last;
}

function renderAnalysis(items) {
  if (!items.length) {
    analysisPanel.innerHTML = `
      <p>Не знайшла рядків формату <b>Снаряд (Заряд)</b>.</p>
    `;
    return;
  }

  const totalQuantity = items.reduce((sum, item) => sum + Number(item.quantity || 0), 0);
  const longRangeQuantity = items
    .filter(item => item.longRange)
    .reduce((sum, item) => sum + Number(item.quantity || 0), 0);

  const uniqueCombinations = new Set(
    items.map(item => `${item.projectile}|${item.charge}`)
  ).size;

  let html = `
    <div class="analysis-grid">
      <div class="metric-card">
        <div class="metric-label">Знайдено рядків</div>
        <div class="metric-value">${items.length}</div>
      </div>

      <div class="metric-card">
        <div class="metric-label">Комбінацій</div>
        <div class="metric-value">${uniqueCombinations}</div>
      </div>

      <div class="metric-card">
        <div class="metric-label">Загальний залишок</div>
        <div class="metric-value">${totalQuantity}</div>
      </div>
    </div>

    <div class="table-panel analysis-table">
      <h2>Первинний аналіз</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Аркуш</th>
              <th>Снаряд</th>
              <th>Заряд</th>
              <th>Дальність</th>
              <th>Далекобійна</th>
              <th>Кількість</th>
            </tr>
          </thead>
          <tbody>
  `;

  items.forEach(item => {
    html += `
      <tr>
        <td>${item.sheetName}</td>
        <td>${item.projectile}</td>
        <td>${item.charge}</td>
        <td>${item.range ?? ""}</td>
        <td>${item.longRange ? "Так" : "Ні"}</td>
        <td>${item.quantity ?? ""}</td>
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
