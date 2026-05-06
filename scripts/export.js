const exportExcelBtn = document.getElementById("exportExcelBtn");

exportExcelBtn.addEventListener("click", exportAnalysisToExcel);

function exportAnalysisToExcel() {
  const items = window.ArtAmmoState?.unitItems || [];
  const grouped = window.ArtAmmoState?.groupedByUnit || {};

  if (!items.length) {
    alert("Немає даних для експорту. Спочатку натисни «Аналізувати файл».");
    return;
  }

  const workbook = XLSX.utils.book_new();

  const summaryRows = Object.values(grouped).map(unit => ({
    "Підрозділ": unit.unit,
    "Комбінацій": unit.combinations.size,
    "Отримання": unit.totalReceived,
    "Витрата": unit.totalSpent,
    "Залишок": unit.totalBalance,
    "Далекобійних": unit.longRangeBalance
  }));

  const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
  XLSX.utils.book_append_sheet(workbook, summarySheet, "Підсумок");

  const detailRows = items.map(item => ({
    "Підрозділ": item.unit,
    "Рядок": item.row,
    "Категорія": item.category,
    "Снаряд": item.projectile,
    "Заряд": item.charge,
    "Примітка": item.note,
    "Дальність, м": item.rangeMeters || "",
    "Дальність, км": item.rangeKm ? item.rangeKm.toFixed(1) : "",
    "Далекобійна": item.longRange ? "Так" : "Ні",
    "Отримання": item.received,
    "Витрата": item.spent,
    "Залишок": item.balance
  }));

  const detailSheet = XLSX.utils.json_to_sheet(detailRows);
  XLSX.utils.book_append_sheet(workbook, detailSheet, "Детально");

  Object.values(grouped).forEach(unit => {
    const unitRows = unit.items.map(item => ({
      "Рядок": item.row,
      "Категорія": item.category,
      "Снаряд": item.projectile,
      "Заряд": item.charge,
      "Примітка": item.note,
      "Дальність, м": item.rangeMeters || "",
      "Дальність, км": item.rangeKm ? item.rangeKm.toFixed(1) : "",
      "Далекобійна": item.longRange ? "Так" : "Ні",
      "Отримання": item.received,
      "Витрата": item.spent,
      "Залишок": item.balance
    }));

    const safeSheetName = unit.unit
      .replace(/[\\/?*\[\]:]/g, "")
      .slice(0, 31);

    const unitSheet = XLSX.utils.json_to_sheet(unitRows);
    XLSX.utils.book_append_sheet(workbook, unitSheet, safeSheetName);
  });

  const now = new Date();
  const fileDate = now.toISOString().slice(0, 10);

  XLSX.writeFile(workbook, `art_ammo_export_${fileDate}.xlsx`);
}
