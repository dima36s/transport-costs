// Импорт журнала рейсов из Excel (формат «N | ДАТА | РЕЙС | ВОДИТЕЛЬ | СТОИМОСТЬ | КОМЕНТАРИЙ»,
// плюс необязательные колонки «Расходы» и «За что»). Каждый лист читается отдельно.

import { loadExcelJS, nameKey } from './lib.js';
import { toISO, formatDate } from './periods.js';

const HEADERS = {
  num: [/^n$/, /^№/],
  date: [/дата/],
  route: [/рейс/, /маршрут/],
  driver: [/водитель/],
  cost: [/стоимость/, /^сумма/],
  comment: [/коментарий/, /комментарий/, /оплата \(/, /нал\/безнал/],
  expAmount: [/расходы/],
  expWhat: [/за что/],
};

export async function parseWorkbook(file) {
  const ExcelJS = await loadExcelJS();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());

  const result = { trips: [], expenses: [], notes: [], sheets: [] };

  wb.eachSheet((ws) => {
    const rows = [];
    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const vals = [];
      row.eachCell({ includeEmpty: true }, (cell, col) => {
        // в объединённых ячейках ExcelJS повторяет значение во всех клетках — берём только главную
        const slave = cell.isMerged && cell.master && cell.master.address !== cell.address;
        vals[col] = slave ? null : cellValue(cell.value);
      });
      rows[rowNumber] = vals;
    });
    const parsed = parseSheet(ws.name, rows);
    if (parsed) {
      result.trips.push(...parsed.trips);
      result.expenses.push(...parsed.expenses);
      result.notes.push(...parsed.notes);
      result.sheets.push({ name: ws.name, trips: parsed.trips.length, expenses: parsed.expenses.length });
    }
  });

  assignKeys(result);
  return result;
}

function cellValue(v) {
  if (v == null) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if ('result' in v) return cellValue(v.result);
    if (Array.isArray(v.richText)) return v.richText.map((r) => r.text).join('');
    if ('text' in v) return v.text;
    if ('error' in v) return null;
  }
  return v;
}

const text = (v) => (v == null ? '' : v instanceof Date ? '' : String(v).replace(/\s+/g, ' ').trim());
const norm = (v) => text(v).toLowerCase().replace(/ё/g, 'е');

function findColumns(rows) {
  for (let r = 1; r < Math.min(rows.length, 30); r++) {
    const row = rows[r];
    if (!row) continue;
    if (!row.some((v) => /дата/.test(norm(v)))) continue;
    const cols = {};
    let used = 1;
    for (const rr of [r, r + 1]) {
      const hr = rows[rr] || [];
      let hit = false;
      hr.forEach((v, c) => {
        const t = norm(v);
        if (!t) return;
        for (const [key, pats] of Object.entries(HEADERS)) {
          if (cols[key] == null && pats.some((p) => p.test(t))) {
            cols[key] = c;
            hit = true;
          }
        }
      });
      if (rr === r + 1 && hit) used = 2;
    }
    if (cols.date != null && (cols.route != null || cols.cost != null)) {
      return { cols, firstDataRow: r + used };
    }
  }
  return null;
}

function parseDate(v) {
  if (v instanceof Date) {
    // ExcelJS отдаёт даты в UTC; число «20» в ячейке с форматом даты превращается в 1900 год — это не дата
    const y = v.getUTCFullYear();
    if (y < 2000 || y > 2100) return null;
    return toISO(new Date(y, v.getUTCMonth(), v.getUTCDate()));
  }
  if (typeof v === 'number') {
    if (v < 30000 || v > 80000) return null; // явно не дата (например «20»)
    const ms = Math.round((v - 25569) * 86400000);
    const d = new Date(ms);
    return toISO(new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  }
  const s = text(v);
  let m = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (m) {
    const y = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
    return toISO(new Date(y, Number(m[2]) - 1, Number(m[1])));
  }
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  return null;
}

function parseAmount(v) {
  if (typeof v === 'number') return v;
  const s = text(v).replace(/\s/g, '').replace(',', '.');
  if (!s) return null;
  const n = Number(s.replace(/[^\d.-]/g, ''));
  return /\d/.test(s) && Number.isFinite(n) ? n : NaN;
}

export function detectPayment(comment, costRaw) {
  const c = norm(comment);
  const raw = norm(costRaw);
  if (/наш/.test(raw) || /наш/.test(c)) return 'own';
  if (/б\s*[/\\]\s*н|безнал/.test(c)) return 'bn';
  if (/\bнал\b|^нал/.test(c)) return 'cash';
  return '';
}

export function detectCategory(what) {
  const t = norm(what);
  if (/топлив|бензин|дизел|дт\b|солярк|газ\b/.test(t)) return 'Топливо';
  if (/платн|платон|трасс|м-?4|дорог/.test(t)) return 'Платные дороги';
  if (/\bто\b|ремонт|шин|запчаст|сервис|масл/.test(t)) return 'ТО и ремонт';
  if (/зп|зарплат|оклад|аванс/.test(t)) return 'Зарплата';
  if (/штра?ф|гаи|гибдд/.test(t)) return 'Штрафы';
  if (/налог|ндс|ндфл/.test(t)) return 'Налоги';
  if (/офис|камера|связь|аренд|интернет/.test(t)) return 'Офис и связь';
  return 'Прочее';
}

function parseSheet(sheetName, rows) {
  const found = findColumns(rows);
  if (!found) return null;
  const { cols, firstDataRow } = found;
  const trips = [];
  const rawExpenses = [];
  const notes = [];
  const note = (row, msg) => notes.push({ sheet: sheetName, row, text: msg });
  const get = (row, key) => (cols[key] == null ? null : (rows[row] || [])[cols[key]]);

  let lastDate = null;
  const consumed = new Set();

  for (let r = firstDataRow; r < rows.length; r++) {
    if (consumed.has(r) || !rows[r]) continue;
    const route = text(get(r, 'route'));
    const driver = text(get(r, 'driver'));
    const costRaw = get(r, 'cost');
    const comment = text(get(r, 'comment'));
    const num = text(get(r, 'num'));

    // расходы в той же строке
    const expRaw = get(r, 'expAmount');
    const expWhat = text(get(r, 'expWhat'));
    const expAmount = parseAmount(expRaw);
    if (expAmount && expAmount > 0 && !expWhat && !route && !driver) {
      // итоговая строка расходов — пропускаем
    } else if (expAmount && expAmount > 0) {
      rawExpenses.push({ row: r, amount: expAmount, what: expWhat });
    } else if (expWhat && (expAmount == null || Number.isNaN(expAmount))) {
      note(r, `«${expWhat}» — сумма расхода не указана, строка расхода пропущена`);
    }

    if (!route && !driver) continue; // пустые строки, итоги, номера без данных

    let date = parseDate(get(r, 'date'));
    if (!date) {
      const rawDate = get(r, 'date');
      // дата переехала на следующую строку
      const next = r + 1;
      const nextDate = parseDate((rows[next] || [])[cols.date]);
      const nextHasData = text((rows[next] || [])[cols.route]) || text((rows[next] || [])[cols.driver]);
      if (nextDate && !nextHasData) {
        date = nextDate;
        consumed.add(next);
        note(r, `«${route}» — дата стояла строкой ниже, взята ${formatDate(date)}`);
      } else if (lastDate) {
        date = lastDate;
        const shown = rawDate == null || rawDate === ''
          ? 'не указана'
          : `«${rawDate instanceof Date ? Math.round((rawDate.getTime() - Date.UTC(1899, 11, 30)) / 86400000) : text(rawDate)}» некорректна`;
        note(r, `«${route}» — дата ${shown}, подставлена дата предыдущей строки ${formatDate(date)}. Проверьте`);
      } else {
        note(r, `«${route}» — нет даты, строка пропущена`);
        continue;
      }
    }
    lastDate = date;

    let cost = parseAmount(costRaw);
    let payment = detectPayment(comment, costRaw);
    if (cost == null) {
      cost = 0;
      if (!payment) note(r, `«${route}» — стоимость не указана, записано 0`);
    } else if (Number.isNaN(cost)) {
      if (payment !== 'own') note(r, `«${route}» — стоимость «${text(costRaw)}» не число, записано 0`);
      cost = 0;
    }
    if (!payment && cost === 0) note(r, `«${route}» — стоимость 0 и вид оплаты не указан`);

    trips.push({ date, route, driver, cost, payment, comment: commentLeftover(comment, payment), _sheet: sheetName, _row: r });
  }

  // Расходы без своей даты привязываем к первой дате рейсов листа
  const dates = trips.map((t) => t.date).sort();
  const expenseDate = dates[0] || null;
  const expenses = [];
  for (const e of rawExpenses) {
    if (!expenseDate) {
      note(e.row, `Расход «${e.what}» пропущен: на листе нет ни одной даты`);
      continue;
    }
    expenses.push({
      date: expenseDate,
      category: detectCategory(e.what),
      amount: e.amount,
      comment: e.what,
      _sheet: sheetName,
      _row: e.row,
    });
  }
  if (expenses.length) {
    notes.push({
      sheet: sheetName,
      row: null,
      text: `Расходы (${expenses.length} шт.) записаны на ${formatDate(expenseDate)} — первую дату рейсов листа, потому что своих дат у них нет`,
    });
  }

  return { trips, expenses, notes };
}

function commentLeftover(comment, payment) {
  // «б/н», «нал», «наш груз» уходят в поле «Оплата», остальное остаётся комментарием
  const c = comment.replace(/^(б\s*[/\\]\s*н|безнал|нал|наш груз|наш)$/i, '').trim();
  return c;
}

function assignKeys(result) {
  const seen = new Map();
  const dupNotes = new Map();
  for (const t of result.trips) {
    const base = `t|${t.date}|${nameKey(t.route)}|${nameKey(t.driver)}|${t.cost}`;
    const n = (seen.get(base) || 0) + 1;
    seen.set(base, n);
    t.importKey = `${base}|${n}`;
    if (n === 2) dupNotes.set(base, t);
  }
  for (const t of dupNotes.values()) {
    result.notes.push({
      sheet: t._sheet,
      row: t._row,
      text: `Повторяется рейс «${t.route}», ${t.driver}, ${formatDate(t.date)} — проверьте, не дубль ли (оба импортированы)`,
    });
  }
  const seenE = new Map();
  for (const e of result.expenses) {
    const base = `e|${e.date}|${e.category}|${nameKey(e.comment)}|${e.amount}`;
    const n = (seenE.get(base) || 0) + 1;
    seenE.set(base, n);
    e.importKey = `${base}|${n}`;
  }
}
