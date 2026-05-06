const exportExcelBtn = document.getElementById("exportExcelBtn");

exportExcelBtn.addEventListener("click", exportAnalysisToExcel);

function exportAnalysisToExcel() {
  const items = window.ArtAmmoState?.analysisItems || [];

  if (!items.length) {
    alert("Немає даних для експорту. Спочатку натисни «Аналізувати файл».");
    return;
  }

  const rows = items.map(item => ({
    "Аркуш": item.sheetName,
    "Снаряд": item.projectile,
    "Заряд": item.charge,
    "Дальність, км": item.range ?? "",
    "Далекобійна": item.longRange ? "Так" : "Ні",
    "Кількість": item.quantity
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows);

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, worksheet, "Первинний аналіз");

  XLSX.writeFile(workbook, "art_ammo_analysis.xlsx");
}
