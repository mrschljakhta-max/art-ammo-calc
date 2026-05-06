const exportPdfBtn = document.getElementById("exportPdfBtn");

exportPdfBtn.addEventListener("click", exportAnalysisToPdf);

function exportAnalysisToPdf() {
  const items = window.ArtAmmoState?.analysisItems || [];

  if (!items.length) {
    alert("Немає даних для PDF. Спочатку натисни «Аналізувати файл».");
    return;
  }

  const { jsPDF } = window.jspdf;

  const doc = new jsPDF({
    orientation: "landscape",
    unit: "mm",
    format: "a4"
  });

  const totalQuantity = items.reduce(
    (sum, item) => sum + Number(item.quantity || 0),
    0
  );

  const longRangeQuantity = items
    .filter(item => item.longRange)
    .reduce((sum, item) => sum + Number(item.quantity || 0), 0);

  const uniqueCombinations = new Set(
    items.map(item => `${item.projectile}|${item.charge}`)
  ).size;

  const now = new Date();
  const dateText = now.toLocaleDateString("uk-UA");

  doc.setFontSize(16);
  doc.text("ART AMMO — ЗВІТ ПЕРВИННОГО АНАЛІЗУ", 14, 16);

  doc.setFontSize(10);
  doc.text(`Дата формування: ${dateText}`, 14, 24);
  doc.text(`Комбінацій: ${uniqueCombinations}`, 14, 31);
  doc.text(`Загальний залишок: ${totalQuantity}`, 70, 31);
  doc.text(`Далекобійних: ${longRangeQuantity}`, 140, 31);

  const tableRows = items.map(item => [
    item.sheetName,
    item.projectile,
    item.charge,
    item.range ?? "",
    item.longRange ? "Так" : "Ні",
    item.quantity ?? ""
  ]);

  doc.autoTable({
    startY: 40,
    head: [[
      "Аркуш",
      "Снаряд",
      "Заряд",
      "Дальність, км",
      "Далекобійна",
      "Кількість"
    ]],
    body: tableRows,
    styles: {
      fontSize: 8,
      cellPadding: 2
    },
    headStyles: {
      fillColor: [37, 43, 58]
    }
  });

  const fileDate = now.toISOString().slice(0, 10);
  doc.save(`art_ammo_report_${fileDate}.pdf`);
}
