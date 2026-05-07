(function () {
  const exportPdfBtn = document.getElementById("exportPdfBtn");

  if (!exportPdfBtn) return;

  exportPdfBtn.addEventListener("click", exportAnalysisToPdf);

  function exportAnalysisToPdf() {
    const items = typeof getCurrentFilteredItems === "function"
      ? getCurrentFilteredItems()
      : (window.ArtAmmoState?.unitItems || []);

    if (!items.length) {
      alert("Немає даних для PDF. Спочатку завантаж Excel-файл.");
      return;
    }

    const grouped = typeof groupByUnit === "function" ? groupByUnit(items) : {};
    const commanderSummary = typeof getCommanderSummary === "function"
      ? getCommanderSummary(items, grouped)
      : [];

    const reportPassport = typeof getReportPassport === "function"
      ? getReportPassport(items, grouped)
      : [];

    const totalBalance = items.reduce((sum, item) => sum + Number(item.balance || 0), 0);
    const totalReceived = items.reduce((sum, item) => sum + Number(item.received || 0), 0);
    const totalSpent = items.reduce((sum, item) => sum + Number(item.spent || 0), 0);
    const longRangeBalance = items
      .filter(item => item.longRange)
      .reduce((sum, item) => sum + Number(item.balance || 0), 0);
    const uniqueCombinations = new Set(items.map(item => item.combination)).size;

    const now = new Date();
    const dateText = now.toLocaleDateString("uk-UA");
    const fileDate = now.toISOString().slice(0, 10);

    const content = [
      { text: "BASTION — звіт аналізу залишків", style: "header" },
      { text: `Дата формування: ${dateText}`, margin: [0, 0, 0, 12] },
      {
        columns: [
          { text: `Підрозділів: ${Object.keys(grouped).length}`, style: "metric" },
          { text: `Комбінацій: ${uniqueCombinations}`, style: "metric" },
          { text: `Отримання: ${totalReceived}`, style: "metric" },
          { text: `Витрата: ${totalSpent}`, style: "metric" },
          { text: `Залишок: ${totalBalance}`, style: "metric" },
          { text: `Далекобійних: ${longRangeBalance}`, style: "metric" }
        ],
        margin: [0, 0, 0, 16]
      }
    ];

    if (reportPassport.length) {
      content.push(
        { text: "Паспорт звіту", style: "section" },
        {
          table: {
            headerRows: 1,
            widths: ["auto", "*"],
            body: [
              [
                { text: "Параметр", bold: true },
                { text: "Значення", bold: true }
              ],
              ...reportPassport.map(item => [item.label, item.value])
            ]
          },
          layout: "lightHorizontalLines",
          margin: [0, 0, 0, 18]
        }
      );
    }

    if (commanderSummary.length) {
      content.push(
        { text: "Короткий командирський висновок", style: "section" },
        {
          table: {
            headerRows: 1,
            widths: ["auto", "*"],
            body: [
              [
                { text: "Блок", bold: true },
                { text: "Висновок", bold: true }
              ],
              ...commanderSummary.map(item => [item.label, item.text])
            ]
          },
          layout: "lightHorizontalLines",
          margin: [0, 0, 0, 18]
        }
      );
    }

    const exchangeRecommendations = typeof getFilteredExchangeRecommendations === "function"
      ? getFilteredExchangeRecommendations(window.ArtAmmoState?.unitItems || items, typeof getLowBalanceThreshold === "function" ? getLowBalanceThreshold() : 10, 20)
      : (typeof getExchangeRecommendations === "function"
        ? getExchangeRecommendations(window.ArtAmmoState?.unitItems || items, typeof getLowBalanceThreshold === "function" ? getLowBalanceThreshold() : 10, 20)
        : []);

    if (exchangeRecommendations.length) {
      content.push(
        { text: "Рекомендації щодо обміну", style: "section" },
        {
          table: {
            headerRows: 1,
            widths: ["auto", "*", "auto", "*", "*", "auto", "*"],
            body: [
              [
                { text: "Пріоритет", bold: true },
                { text: "Комбінація", bold: true },
                { text: "Дальність", bold: true },
                { text: "Передати з", bold: true },
                { text: "Передати до", bold: true },
                { text: "К-сть", bold: true },
                { text: "Пояснення", bold: true }
              ],
              ...exchangeRecommendations.map(item => [
                item.type,
                `${item.projectile} (${item.charge})`,
                item.rangeKm ? item.rangeKm.toFixed(1) : "",
                `${item.fromUnit} (${item.donorBalance})`,
                `${item.toUnit} (${item.receiverBalance})`,
                String(item.recommendedQty),
                `${item.reason}; очікувано ${item.expectedReceiverBalance}`
              ])
            ]
          },
          layout: "lightHorizontalLines",
          margin: [0, 0, 0, 12]
        }
      );

      if (typeof getExchangeActionPlan === "function") {
        const actionPlan = typeof getFilteredExchangeActionPlan === "function"
          ? getFilteredExchangeActionPlan(exchangeRecommendations)
          : getExchangeActionPlan(exchangeRecommendations);

        if (actionPlan.length) {
          content.push(
            { text: "Журнал рекомендованих дій", style: "section" },
            {
              table: {
                headerRows: 1,
                widths: ["auto", "auto", "*", "*", "*", "*", "*"],
                body: [
                  [
                    { text: "№", bold: true },
                    { text: "Пріоритет", bold: true },
                    { text: "Дія", bold: true },
                    { text: "Звідки", bold: true },
                    { text: "Куди", bold: true },
                    { text: "Буде", bold: true },
                    { text: "Ризик", bold: true }
                  ],
                  ...actionPlan.map(item => [
                    String(item.order),
                    item.priority,
                    item.action,
                    item.fromUnit,
                    item.toUnit,
                    item.after,
                    item.risk
                  ])
                ]
              },
              layout: "lightHorizontalLines",
              margin: [0, 0, 0, 18]
            }
          );
        }
      }

      if (typeof getExchangeImpactSummary === "function") {
        const impact = getExchangeImpactSummary(exchangeRecommendations);
        content.push(
          { text: "Очікуваний ефект обміну", style: "section" },
          {
            table: {
              headerRows: 1,
              widths: ["*", "auto"],
              body: [
                [{ text: "Показник", bold: true }, { text: "Значення", bold: true }],
                ["Кількість рекомендацій", String(impact.count)],
                ["Рекомендовано передати загалом", String(impact.totalRecommended)],
                ["Далекобійних у передачі", String(impact.longRangeRecommended)],
                ["Закрито нульових позицій", String(impact.zeroClosed)],
                ["Піднято вище порогу", String(impact.raisedAboveThreshold)],
                ["Донорів біля порогу після передачі", String(impact.donorWarnings)],
                ["Поріг малого залишку", `≤${impact.threshold}`]
              ]
            },
            layout: "lightHorizontalLines",
            margin: [0, 0, 0, 18]
          }
        );
      }
    }

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

    content.push(
      { text: "Підсумок по підрозділах", style: "section" },
      {
        table: {
          headerRows: 1,
          widths: ["*", "auto", "auto", "auto", "auto", "auto"],
          body: unitSummaryBody
        },
        layout: "lightHorizontalLines",
        margin: [0, 0, 0, 18]
      }
    );

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

    content.push(
      { text: "Детальний аналіз", style: "section" },
      {
        table: {
          headerRows: 1,
          widths: ["*", "*", "*", "auto", "auto", "auto"],
          body: detailsBody
        },
        layout: "lightHorizontalLines"
      }
    );

    const docDefinition = {
      pageSize: "A4",
      pageOrientation: "landscape",
      pageMargins: [24, 28, 24, 28],
      defaultStyle: {
        font: "Roboto",
        fontSize: 8
      },
      content,
      styles: {
        header: {
          fontSize: 16,
          bold: true,
          margin: [0, 0, 0, 8]
        },
        section: {
          fontSize: 12,
          bold: true,
          margin: [0, 8, 0, 8]
        },
        metric: {
          fontSize: 8,
          bold: true
        }
      }
    };

    const settings = window.BastionSettings || {};
    const prefix = String(settings.exportPrefix || "bastion").trim() || "bastion";
    pdfMake.createPdf(docDefinition).download(`${prefix}_report_${fileDate}.pdf`);
  }
})();
