const analyzeBtn = document.getElementById("analyzeBtn");

const analysisPanel = document.getElementById("analysisPanel");

analyzeBtn.addEventListener("click", analyzeWorkbook);

function analyzeWorkbook() {

  const workbook = window.ArtAmmoState?.workbook;

  if (!workbook) {

    analysisPanel.innerHTML = `
      <p>Спочатку завантаж Excel.</p>
    `;

    return;
  }

  const result = [];

  workbook.SheetNames.forEach((sheetName) => {

    const sheet = workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: ""
    });

    const parsedItems = parseSheet(sheetName, rows);

    result.push(...parsedItems);

  });

  window.ArtAmmoState.analysisItems = result;

  document.getElementById("exportExcelBtn").disabled = false;

  document.getElementById("exportPdfBtn").disabled = false;

  renderAnalysis(result);

}

function parseSheet(sheetName, rows) {

  const items = [];

  rows.forEach((row) => {

    const cells = row.map(cell =>
      String(cell || "").trim()
    );

    cells.forEach((cell) => {

      const match = cell.match(/^(.+?)\s*\((.+?)\)$/);

      if (!match) return;

      const projectile = match[1].trim();

      const charge = match[2].trim();

      const range = findRange(cells);

      const quantity = findQuantity(cells);

      items.push({
        sheetName,
        projectile,
        charge,
        range,
        quantity,
        longRange: range >= 18
      });

    });

  });

  return items;

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

    if (
      !Number.isNaN(number) &&
      number > 0 &&
      number <= 70
    ) {
      return number;
    }

  }

  return null;

}

function findQuantity(cells) {

  let lastNumber = 0;

  cells.forEach((cell) => {

    const value = Number(
      String(cell).replace(",", ".")
    );

    if (!Number.isNaN(value)) {
      lastNumber = value;
    }

  });

  return lastNumber;

}

function renderAnalysis(items) {

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
    (sum, item) =>
      sum + Number(item.quantity || 0),
    0
  );

  const longRangeQuantity = items
    .filter(item => item.longRange)
    .reduce(
      (sum, item) =>
        sum + Number(item.quantity || 0),
      0
    );

  const uniqueCombinations = new Set(
    items.map(item =>
      `${item.projectile}|${item.charge}`
    )
  ).size;

  let html = `
    <div class="analysis-grid">

      <div class="metric-card">
        <div class="metric-label">
          Комбінацій
        </div>

        <div class="metric-value">
          ${uniqueCombinations}
        </div>
      </div>

      <div class="metric-card">
        <div class="metric-label">
          Загальний залишок
        </div>

        <div class="metric-value">
          ${totalQuantity}
        </div>
      </div>

      <div class="metric-card">
        <div class="metric-label">
          Далекобійних
        </div>

        <div class="metric-value">
          ${longRangeQuantity}
        </div>
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

  items.forEach((item) => {

    html += `
      <tr>

        <td>${item.sheetName}</td>

        <td>${item.projectile}</td>

        <td>${item.charge}</td>

        <td>${item.range ?? ""}</td>

        <td>
          ${item.longRange ? "Так" : "Ні"}
        </td>

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
