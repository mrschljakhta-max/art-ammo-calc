(function () {
  const btn = document.getElementById("exportPdfBtn");
  if (!btn) return;

  btn.addEventListener("click", exportAnalysisToPdf);

  function exportAnalysisToPdf() {
    const items = typeof getCurrentFilteredItems === "function"
      ? getCurrentFilteredItems()
      : (window.ArtAmmoState?.unitItems || []);

    if (!items.length) {
      alert("Немає даних для PDF. Спочатку натисни «Аналізувати файл».");
      return;
    }

    const grouped = groupByUnit(items);
    const groupedCategories = typeof groupByCategory === "function" ? groupByCategory(items) : {};
    const groupedRanges = typeof groupByRangeBand === "function" ? groupByRangeBand(items) : {};
    const criticalItems = typeof getCriticalItems === "function"
      ? getCriticalItems(items)
      : items.filter(item => Number(item.balance || 0) === 0);
    const lowBalanceItems = typeof getLowBalanceItems === "function"
      ? getLowBalanceItems(items, 10)
      : items.filter(item => Number(item.balance || 0) > 0 && Number(item.balance || 0) <= 10);

    const totalBalance = items.reduce((sum, item) => sum + Number(item.balance || 0), 0);
    const totalReceived = items.reduce((sum, item) => sum + Number(item.received || 0), 0);
    const totalSpent = items.reduce((sum, item) => sum + Number(item.spent || 0), 0);
    const longRangeBalance = items
      .filter(item => item.longRange)
      .reduce((sum, item) => sum + Number(item.balance || 0), 0);
    const shortRangeBalance = totalBalance - longRangeBalance;
    const uniqueCombinations = new Set(items.map(item => item.combination)).size;

    const now = new Date();
    const dateText = now.toLocaleDateString("uk-UA");
    const fileDate = now.toISOString().slice(0, 10);

    const unitSummaryBody = [
      [
        { text: "Підрозділ", bold: true },
        { text: "Комбінацій", bold: true },
        { text: "Отримання", bold: true },
        { text: "Витрата", bold: true },
        { text: "Залишок", bold: true },
        { text: "Далекобійних", bold: true }
      ]
    ];

    Object.values(grouped).forEach(unit => {
      unitSummaryBody.push([
        unit.unit,
        String(unit.combinations.size),
        String(unit.totalReceived),
        String(unit.totalSpent),
        String(unit.totalBalance),
        String(unit.longRangeBalance)
      ]);
    });

    const categoryBody = [
      [
        { text: "Категорія", bold: true },
        { text: "Комбінацій", bold: true },
        { text: "Отримання", bold: true },
        { text: "Витрата", bold: true },
        { text: "Залишок", bold: true },
        { text: "Далекобійних", bold: true }
      ]
    ];

    Object.values(groupedCategories).forEach(category => {
      categoryBody.push([
        category.category,
        String(category.combinations.size),
        String(category.totalReceived),
        String(category.totalSpent),
        String(category.totalBalance),
        String(category.longRangeBalance)
      ]);
    });

    const rangeBody = [
      [
        { text: "Діапазон", bold: true },
        { text: "Комбінацій", bold: true },
        { text: "Рядків", bold: true },
        { text: "Залишок", bold: true }
      ]
    ];

    Object.values(groupedRanges).forEach(range => {
      rangeBody.push([
        range.band,
        String(range.combinations.size),
        String(range.items.length),
        String(range.totalBalance)
      ]);
    });

    const bottleneckBody = [
      [
        { text: "Статус", bold: true },
        { text: "Підрозділ", bold: true },
        { text: "Снаряд", bold: true },
        { text: "Заряд", bold: true },
        { text: "Дальність, км", bold: true },
        { text: "Далекобійна", bold: true },
        { text: "Залишок", bold: true }
      ]
    ];

    [...criticalItems, ...lowBalanceItems].slice(0, 80).forEach(item => {
      bottleneckBody.push([
        Number(item.balance || 0) === 0 ? "Нуль" : "Мало",
        item.unit,
        item.projectile,
        item.charge,
        item.rangeKm ? item.rangeKm.toFixed(1) : "",
        item.longRange ? "Так" : "Ні",
        String(item.balance)
      ]);
    });

    if (bottleneckBody.length === 1) {
      bottleneckBody.push(["—", "Критичних позицій не знайдено", "", "", "", "", ""]);
    }

    const detailsBody = [
      [
        { text: "Підрозділ", bold: true },
        { text: "Снаряд", bold: true },
        { text: "Заряд", bold: true },
        { text: "Дальність, км", bold: true },
        { text: "Далекобійна", bold: true },
        { text: "Залишок", bold: true }
      ]
    ];

    items.forEach(item => {
      detailsBody.push([
        item.unit,
        item.projectile,
        item.charge,
        item.rangeKm ? item.rangeKm.toFixed(1) : "",
        item.longRange ? "Так" : "Ні",
        String(item.balance)
      ]);
    });

    const docDefinition = {
      pageSize: "A4",
      pageOrientation: "landscape",
      pageMargins: [24, 28, 24, 28],
      defaultStyle: {
        font: "Roboto",
        fontSize: 8
      },
      content: [
        { text: "ART AMMO — звіт аналізу залишків", style: "header" },
        { text: `Дата формування: ${dateText}`, margin: [0, 0, 0, 12] },
        {
          columns: [
            { text: `Підрозділів: ${Object.keys(grouped).length}`, style: "metric" },
            { text: `Комбінацій: ${uniqueCombinations}`, style: "metric" },
            { text: `Отримання: ${totalReceived}`, style: "metric" },
            { text: `Витрата: ${totalSpent}`, style: "metric" },
            { text: `Залишок: ${totalBalance}`, style: "metric" },
            { text: `Далекобійних: ${longRangeBalance}`, style: "metric" },
            { text: `Недалекобійних: ${shortRangeBalance}`, style: "metric" },
            { text: `Нульових: ${criticalItems.length}`, style: "metric" },
            { text: `Малий залишок: ${lowBalanceItems.length}`, style: "metric" }
          ],
          margin: [0, 0, 0, 16]
        },
        { text: "Підсумок по підрозділах", style: "section" },
        {
          table: {
            headerRows: 1,
            widths: ["*", "auto", "auto", "auto", "auto", "auto"],
            body: unitSummaryBody
          },
          layout: "lightHorizontalLines",
          margin: [0, 0, 0, 14]
        },
        { text: "Підсумок по категоріях", style: "section" },
        {
          table: {
            headerRows: 1,
            widths: ["*", "auto", "auto", "auto", "auto", "auto"],
            body: categoryBody
          },
          layout: "lightHorizontalLines",
          margin: [0, 0, 0, 14]
        },
        { text: "Розподіл по дальності", style: "section" },
        {
          table: {
            headerRows: 1,
            widths: ["*", "auto", "auto", "auto"],
            body: rangeBody
          },
          layout: "lightHorizontalLines",
          margin: [0, 0, 0, 14]
        },
        { text: "Вузькі місця", style: "section" },
        {
          table: {
            headerRows: 1,
            widths: ["auto", "*", "*", "*", "auto", "auto", "auto"],
            body: bottleneckBody
          },
          layout: "lightHorizontalLines",
          margin: [0, 0, 0, 14]
        },
        { text: "Детальний аналіз", style: "section" },
        {
          table: {
            headerRows: 1,
            widths: ["*", "*", "*", "auto", "auto", "auto"],
            body: detailsBody
          },
          layout: "lightHorizontalLines"
        }
      ],
      styles: {
        header: { fontSize: 16, bold: true, margin: [0, 0, 0, 8] },
        section: { fontSize: 12, bold: true, margin: [0, 8, 0, 8] },
        metric: { fontSize: 8, bold: true }
      }
    };

    pdfMake.createPdf(docDefinition).download(`art_ammo_report_${fileDate}.pdf`);
  }
})();
