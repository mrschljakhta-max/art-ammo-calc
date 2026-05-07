(function () {
  const btn = document.getElementById("exportExcelBtn");

  if (!btn) return;

  btn.addEventListener("click", exportAnalysisToExcel);

  function exportAnalysisToExcel() {
    const items = typeof getCurrentFilteredItems === "function"
      ? getCurrentFilteredItems()
      : (window.ArtAmmoState?.unitItems || []);

    if (!items.length) {
      alert("Немає даних для експорту. Спочатку натисни «Аналізувати файл».");
      return;
    }

    const grouped = groupByUnit(items);
    const groupedCategories = typeof groupByCategory === "function"
      ? groupByCategory(items)
      : {};
    const groupedRanges = typeof groupByRangeBand === "function"
      ? groupByRangeBand(items)
      : {};

    const workbook = XLSX.utils.book_new();

    const summaryRows = Object.values(grouped).map(unit => ({
      "Підрозділ": unit.unit,
      "Комбінацій": unit.combinations.size,
      "Отримання": unit.totalReceived,
      "Витрата": unit.totalSpent,
      "Залишок": unit.totalBalance,
      "Далекобійних": unit.longRangeBalance
    }));

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(summaryRows),
      "Підсумок"
    );

    const categoryRows = Object.values(groupedCategories).map(category => ({
      "Категорія": category.category,
      "Комбінацій": category.combinations.size,
      "Отримання": category.totalReceived,
      "Витрата": category.totalSpent,
      "Залишок": category.totalBalance,
      "Далекобійних": category.longRangeBalance
    }));

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(categoryRows),
      "Категорії"
    );

    const rangeRows = Object.values(groupedRanges).map(range => ({
      "Діапазон": range.band,
      "Комбінацій": range.combinations.size,
      "Рядків": range.items.length,
      "Залишок": range.totalBalance
    }));

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(rangeRows),
      "Дальності"
    );

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

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(detailRows),
      "Детально"
    );

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

      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.json_to_sheet(unitRows),
        safeSheetName || "Підрозділ"
      );
    });

    const now = new Date();
    const fileDate = now.toISOString().slice(0, 10);

    XLSX.writeFile(workbook, `art_ammo_export_${fileDate}.xlsx`);
  }
})();
