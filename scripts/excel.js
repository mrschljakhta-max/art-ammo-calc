const excelInput = document.getElementById("excelFile");

const fileInfo = document.getElementById("fileInfo");

const sheetList = document.getElementById("sheetList");

excelInput.addEventListener("change", handleExcel);

function handleExcel(event) {

  const file = event.target.files[0];

  if (!file) return;

  fileInfo.innerHTML = `
    <p>Файл: ${file.name}</p>
  `;

  const reader = new FileReader();

  reader.onload = function(e) {

    const data = new Uint8Array(e.target.result);

    const workbook = XLSX.read(data, {
      type: "array"
    });

    renderSheets(workbook);

  };

  reader.readAsArrayBuffer(file);
}

function renderSheets(workbook) {

  sheetList.innerHTML = "";

  workbook.SheetNames.forEach(sheetName => {

    const div = document.createElement("div");

    div.className = "sheet-item";

    div.innerHTML = `
      ${sheetName}
    `;

    sheetList.appendChild(div);

  });

}
