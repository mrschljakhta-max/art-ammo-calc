const excelInput = document.getElementById("excelFile");
const fileInfo = document.getElementById("fileInfo");
const sheetList = document.getElementById("sheetList");
const appStatus = document.getElementById("appStatus");

let currentWorkbook = null;

window.ArtAmmoState = {
  workbook: null,
  analysisItems: [],
  unitItems: [],
  summaryItems: [],
  groupedByUnit: {},
  fileName: "",
  loadedAt: null,
  analyzedAt: null
};

excelInput.addEventListener("change", handleExcel);

function handleExcel(event) {
  const file = event.target.files[0];
  if (!file) return;

  resetWorkbookState(file.name);

  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);

      currentWorkbook = XLSX.read(data, {
        type: "array",
        cellDates: true
      });

      window.ArtAmmoState.workbook = currentWorkbook;
      window.ArtAmmoState.loadedAt = new Date().toISOString();

      const analyzeButton = document.getElementById("analyzeBtn");
      if (analyzeButton) analyzeButton.disabled = false;
      setStatus("Файл завантажено", "ok");

      renderSheets(currentWorkbook);

      const settings = window.BastionSettings || {};
      const shouldAnalyze = settings.autoAnalyze !== false;
      const shouldOpenAnalytics = settings.autoOpenAnalytics !== false;

      if (shouldOpenAnalytics && window.BastionNavigation?.activateView) {
        window.BastionNavigation.activateView("analytics");
      }

      if (shouldAnalyze && typeof analyzeWorkbook === "function") {
        analyzeWorkbook();
      } else {
        setStatus("Файл завантажено. Автоаналіз вимкнено.", "ok");
      }
    } catch (error) {
      console.error(error);
      setStatus("Помилка читання Excel", "error");
      alert("Не вдалося прочитати Excel-файл. Перевір формат файлу.");
    }
  };

  reader.onerror = function () {
    setStatus("Помилка завантаження", "error");
    alert("Не вдалося завантажити файл.");
  };

  reader.readAsArrayBuffer(file);
}

function resetWorkbookState(fileName) {
  fileInfo.innerHTML = `<p>Файл: ${escapeHtml(fileName)}</p>`;
  sheetList.innerHTML = "";

  const analysisPanel = document.getElementById("analysisPanel");
  if (analysisPanel) analysisPanel.innerHTML = "";

  const filterStatus = document.getElementById("filterStatus");
  if (filterStatus) {
    filterStatus.hidden = true;
    filterStatus.textContent = "";
  }

  const analyzeBtnNode = document.getElementById("analyzeBtn");
  if (analyzeBtnNode) analyzeBtnNode.disabled = true;
  const exportExcelBtnNode = document.getElementById("exportExcelBtn");
  if (exportExcelBtnNode) exportExcelBtnNode.disabled = true;
  const exportPdfBtnNode = document.getElementById("exportPdfBtn");
  if (exportPdfBtnNode) exportPdfBtnNode.disabled = true;

  const filtersPanel = document.getElementById("filtersPanel");
  if (filtersPanel) filtersPanel.hidden = true;

  const unitFilter = document.getElementById("unitFilter");
  const longRangeOnly = document.getElementById("longRangeOnly");
  const searchFilter = document.getElementById("searchFilter");

  if (unitFilter) unitFilter.innerHTML = `<option value="all">Всі підрозділи</option>`;
  if (longRangeOnly) longRangeOnly.checked = false;
  if (searchFilter) searchFilter.value = "";

  window.ArtAmmoState.workbook = null;
  window.ArtAmmoState.analysisItems = [];
  window.ArtAmmoState.unitItems = [];
  window.ArtAmmoState.summaryItems = [];
  window.ArtAmmoState.groupedByUnit = {};
  window.ArtAmmoState.fileName = fileName || "";
  window.ArtAmmoState.loadedAt = null;
  window.ArtAmmoState.analyzedAt = null;

  setStatus("Читання файлу...", "wait");
}

function renderSheets(workbook) {
  sheetList.innerHTML = "";

  workbook.SheetNames.forEach((sheetName) => {
    const div = document.createElement("button");
    div.type = "button";
    div.className = "sheet-item";
    div.textContent = sheetName;

    div.addEventListener("click", () => {
      renderSheetTable(sheetName);
    });

    sheetList.appendChild(div);
  });
}

function renderSheetTable(sheetName) {
  if (!currentWorkbook) return;

  const sheet = currentWorkbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  const oldPanel = document.querySelector(".sheet-preview-panel");
  if (oldPanel) oldPanel.remove();

  if (!rows.length) {
    sheetList.insertAdjacentHTML("afterend", `<p class="sheet-preview-panel">Аркуш порожній</p>`);
    return;
  }

  const previewRows = rows.slice(0, 80);

  let html = `
    <div class="table-panel sheet-preview-panel">
      <h2>${escapeHtml(sheetName)}</h2>
      <div class="table-wrap">
        <table>
  `;

  previewRows.forEach((row, rowIndex) => {
    html += "<tr>";

    row.forEach((cell) => {
      const safeCell = escapeHtml(cell);
      html += rowIndex === 0 ? `<th>${safeCell}</th>` : `<td>${safeCell}</td>`;
    });

    html += "</tr>";
  });

  html += `
        </table>
      </div>
    </div>
  `;

  sheetList.insertAdjacentHTML("afterend", html);
}

function setStatus(text, type = "wait") {
  if (!appStatus) return;

  appStatus.textContent = text;
  appStatus.dataset.type = type;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
