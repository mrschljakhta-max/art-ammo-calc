(function () {
  const exportExcelBtn = document.getElementById("exportExcelBtn");

  if (!exportExcelBtn) return;

  exportExcelBtn.addEventListener("click", exportAnalysisToExcel);

  function exportAnalysisToExcel() {
    const items = typeof getCurrentFilteredItems === "function"
      ? getCurrentFilteredItems()
      : (window.ArtAmmoState?.unitItems || []);

    if (!items.length) {
      alert("Немає даних для експорту. Спочатку завантаж Excel-файл.");
      return;
    }

    const grouped = typeof groupByUnit === "function" ? groupByUnit(items) : {};
    const workbook = XLSX.utils.book_new();

    const commanderSummary = typeof getCommanderSummary === "function"
      ? getCommanderSummary(items, grouped)
      : [];

    const reportPassport = typeof getReportPassport === "function"
      ? getReportPassport(items, grouped)
      : [];

    if (reportPassport.length) {
      const passportRows = reportPassport.map(item => ({
        "Параметр": item.label,
        "Значення": item.value
      }));

      const passportSheet = XLSX.utils.json_to_sheet(passportRows);
      XLSX.utils.book_append_sheet(workbook, passportSheet, "Паспорт звіту");
    }

    if (commanderSummary.length) {
      const commanderRows = commanderSummary.map(item => ({
        "Блок": item.label,
        "Висновок": item.text
      }));

      const commanderSheet = XLSX.utils.json_to_sheet(commanderRows);
      XLSX.utils.book_append_sheet(workbook, commanderSheet, "Командирський висновок");
    }

    const exchangeRecommendations = typeof getExchangeRecommendations === "function"
      ? getExchangeRecommendations(window.ArtAmmoState?.unitItems || items, typeof getLowBalanceThreshold === "function" ? getLowBalanceThreshold() : 10, 50)
      : [];

    if (exchangeRecommendations.length) {
      const exchangeRows = exchangeRecommendations.map(item => ({
        "Пріоритет": item.type,
        "Снаряд": item.projectile,
        "Заряд": item.charge,
        "Дальність, км": item.rangeKm ? item.rangeKm.toFixed(1) : "",
        "Далекобійна": item.longRange ? "Так" : "Ні",
        "Передати з": item.fromUnit,
        "Залишок донора": item.donorBalance,
        "Передати до": item.toUnit,
        "Залишок отримувача": item.receiverBalance,
        "Рекомендовано передати": item.recommendedQty,
        "Очікуваний залишок отримувача": item.expectedReceiverBalance,
        "Пояснення": item.reason
      }));

      const exchangeSheet = XLSX.utils.json_to_sheet(exchangeRows);
      XLSX.utils.book_append_sheet(workbook, exchangeSheet, "Рекомендації обміну");

      if (typeof getExchangeImpactSummary === "function") {
        const impact = getExchangeImpactSummary(exchangeRecommendations);
        const impactRows = [
          { "Показник": "Кількість рекомендацій", "Значення": impact.count },
          { "Показник": "Рекомендовано передати загалом", "Значення": impact.totalRecommended },
          { "Показник": "Далекобійних у передачі", "Значення": impact.longRangeRecommended },
          { "Показник": "Закрито нульових позицій", "Значення": impact.zeroClosed },
          { "Показник": "Піднято вище порогу", "Значення": impact.raisedAboveThreshold },
          { "Показник": "Донорів біля порогу після передачі", "Значення": impact.donorWarnings },
          { "Показник": "Поріг малого залишку", "Значення": `≤${impact.threshold}` }
        ];

        const impactSheet = XLSX.utils.json_to_sheet(impactRows);
        XLSX.utils.book_append_sheet(workbook, impactSheet, "Ефект обміну");
      }
    }

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

      const safeSheetName = String(unit.unit || "Підрозділ")
        .replace(/[\\/?*\[\]:]/g, "")
        .slice(0, 31) || "Підрозділ";

      const unitSheet = XLSX.utils.json_to_sheet(unitRows);
      XLSX.utils.book_append_sheet(workbook, unitSheet, safeSheetName);
    });

    const now = new Date();
    const fileDate = now.toISOString().slice(0, 10);

    XLSX.writeFile(workbook, `art_ammo_export_${fileDate}.xlsx`);
  }
})();
