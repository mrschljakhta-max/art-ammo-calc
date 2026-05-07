(function () {
  const CYR_TO_LAT = {
    "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H", "О": "O", "Р": "P", "С": "C", "Т": "T", "Х": "X",
    "а": "A", "в": "B", "е": "E", "к": "K", "м": "M", "н": "H", "о": "O", "р": "P", "с": "C", "т": "T", "х": "X"
  };

  const DASH_RE = /[‐‑‒–—―−]/g;

  function transliterateLookalikes(value) {
    return String(value ?? "").replace(/[АВЕКМНОРСТХавекмнорстх]/g, char => CYR_TO_LAT[char] || char);
  }

  function normalizeText(value) {
    return transliterateLookalikes(value)
      .replace(DASH_RE, "-")
      .replace(/\u00A0/g, " ")
      .replace(/[“”«»]/g, '"')
      .replace(/\s+/g, " ")
      .trim();
  }

  function normalizeAmmoName(value) {
    let text = normalizeText(value)
      .replace(/\s*\/\s*/g, "/")
      .replace(/\s*,\s*/g, ", ")
      .replace(/\s*-\s*/g, "-")
      .replace(/\s*\+\s*/g, "+")
      .trim();

    text = text.replace(/^M\s+(?=\d)/i, "M");
    text = text.replace(/^DM\s+(?=\d)/i, "DM");
    text = text.replace(/^L\s+(?=\d)/i, "L");
    text = text.replace(/^ER\s+(?=\d)/i, "ER");

    return text;
  }

  function canonicalKey(value) {
    return normalizeAmmoName(value)
      .toUpperCase()
      .replace(/[\s\-_/.,()]+/g, "")
      .trim();
  }

  window.ArtAmmoNormalizer = {
    normalizeText,
    normalizeAmmoName,
    canonicalKey
  };
})();
