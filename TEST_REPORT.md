# BASTION v0.41 — контрольна збірка

Перевірено:
- структура файлів зібрана в один повний build;
- всі JS-файли проходять синтаксичну перевірку `node --check`;
- зовнішні бібліотеки підключаються через CDN;
- шляхи в `index.html` ведуть у `scripts/` та `styles/`;
- зайві root-дублікати `app.css` / `navigation.js` прибрані.

Що перевірити вручну в браузері:
1. Завантаження Excel.
2. Автоаналіз після завантаження.
3. Перемикання режимів у navbar.
4. PDF / Excel / Decision Package export.
5. History Center / Settings Center.
6. Data Quality Center.
