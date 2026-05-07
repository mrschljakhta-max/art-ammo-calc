const excelInput = document.getElementById("excelFile");
const fileInfo = document.getElementById("fileInfo");
const sheetList = document.getElementById("sheetList");

let currentWorkbook = null;

window.ArtAmmoState = {
  workbook: null,
  analysisItems: [],
  unitItems: [],
  summaryItems: [],
  groupedByUnit: {}
};

excelInput.addEventListener("change", handleExcel);

function handleExcel(event) {
  const file = event.target.files[0];

  if (!file) return;

  fileInfo.innerHTML = `<p>Файл: ${file.name}</p>`;

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);

    currentWorkbook = XLSX.read(data, {
      type: "array",
      cellDates: true
    });

    window.ArtAmmoState.workbook = currentWorkbook;

    document.getElementById("analyzeBtn").disabled = false;

    renderSheets(currentWorkbook);
  };

  reader.readAsArrayBuffer(file);
}

function renderSheets(workbook) {
  sheetList.innerHTML = "";

  workbook.SheetNames.forEach((sheetName) => {
    const div = document.createElement("div");

    div.className = "sheet-item";
    div.textContent = sheetName;

    div.addEventListener("click", () => {
      renderSheetTable(sheetName);
    });

    sheetList.appendChild(div);
  });
}

function renderSheetTable(sheetName) {
  const sheet = currentWorkbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  if (!rows.length) {
    sheetList.innerHTML += `<p>Аркуш порожній</p>`;
    return;
  }

  let html = `
    <div class="table-panel">
      <h2>${sheetName}</h2>
      <div class="table-wrap">
        <table>
  `;

  rows.forEach((row, rowIndex) => {
    html += "<tr>";

    row.forEach((cell) => {
      if (rowIndex === 0) {
        html += `<th>${cell}</th>`;
      } else {
        html += `<td>${cell}</td>`;
      }
    });

    html += "</tr>";
  });

  html += `
        </table>
      </div>
    </div>
  `;

  const oldPanel = document.querySelector(".table-panel");

  if (oldPanel) oldPanel.remove();

  document
    .querySelector(".upload-card")
    .insertAdjacentHTML("beforeend", html);
}
