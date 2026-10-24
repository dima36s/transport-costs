// Общие помощники: формат денег, ленивая загрузка ExcelJS, скачивание файлов.

const rub = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 });
const rub2 = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 2 });

export function money(n) {
  return `${rub.format(Math.round(n || 0))} ₽`;
}
export function moneyExact(n) {
  return `${rub2.format(n || 0)} ₽`;
}
export function num(n) {
  return rub.format(n || 0);
}

let excelPromise = null;
export function loadExcelJS() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  if (!excelPromise) {
    excelPromise = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'vendor/exceljs.min.js';
      s.onload = () => resolve(window.ExcelJS);
      s.onerror = () => {
        excelPromise = null;
        reject(new Error('Не удалось загрузить модуль Excel. Проверьте интернет и попробуйте ещё раз.'));
      };
      document.head.appendChild(s);
    });
  }
  return excelPromise;
}

export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

export function escapeHTML(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/** Ключ для группировки имён: без учёта регистра и ё/е */
export function nameKey(s) {
  return String(s || '').trim().toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' ');
}
