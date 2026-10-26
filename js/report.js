// Отчёт за период в Excel: «Сводка» + листы с записями.
// В режиме «Всё» — все направления, в режиме направления — только оно и его расходы.

import { loadExcelJS, downloadBlob } from './lib.js';
import { PAYMENT } from './db.js';
import { TYPE, INCOME_TYPES, DIRECTION } from './kinds.js';
import { formatDate, periodTitle, periodSubtitle } from './periods.js';

// ExcelJS пишет даты по UTC — иначе в Москве дата съезжает на день назад
const xlDate = (iso) => {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d));
};

const MONEY = '#,##0 "₽";[Red]-#,##0 "₽"';
const DATE = 'dd.mm.yyyy';
const INK = 'FF1A2128';
const MUTED = 'FF5B6670';
const SIGN = 'FF1B4F9C';
const LINE = 'FFD3D9DF';
const HEAD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EDF3' } };

const REPORT_NAME = { all: 'Отчёт', trip: 'Отчёт по рейсам', sale: 'Отчёт по продажам', storage: 'Отчёт по хранению' };
const SCOPE_NAME = { all: 'все направления', trip: 'рейсы', sale: 'продажи', storage: 'хранение' };

export async function buildReport({ period, scope, data, stats }) {
  const ExcelJS = await loadExcelJS();
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Учёт рейсов, продаж и хранения';
  wb.created = new Date();

  const types = scope === 'all' ? INCOME_TYPES : [scope];
  const expenses = scope === 'all' ? data.expense : data.expense.filter((e) => e.direction === scope);

  summarySheet(wb, { period, scope, stats });
  for (const t of types) {
    if (scope !== 'all' || data[t].length) SHEETS[t](wb, data[t]);
  }
  SHEETS.expense(wb, expenses);

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  downloadBlob(blob, `${REPORT_NAME[scope]} ${period.from} — ${period.to}.xlsx`);
}

// ---------------------------------------------------------------------
// Сводка
// ---------------------------------------------------------------------
function summarySheet(wb, { period, scope, stats: s }) {
  const ws = wb.addWorksheet('Сводка', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 38 }, { width: 12 }, { width: 18 }, { width: 18 }, { width: 18 }];

  const sub = periodSubtitle(period);
  ws.addRow([`${REPORT_NAME[scope]}: ${periodTitle(period)}`]).font = { size: 16, bold: true, color: { argb: INK } };
  ws.addRow([`${formatDate(period.from)} – ${formatDate(period.to)}${sub ? `, ${sub}` : ''}; ${SCOPE_NAME[scope]}`]).font = { color: { argb: MUTED } };
  ws.addRow([`Сформирован ${formatDate(isoNow())}`]).font = { color: { argb: MUTED }, size: 9 };
  ws.addRow([]);

  const vatRow = [`в т.ч. НДС ${s.vatRate}% с безналичных поступлений`, '', s.vat];
  const cols3 = [{ title: 'Показатель' }, { title: 'Записей', fmt: 'int' }, { title: 'Сумма', fmt: 'money' }];

  if (scope === 'all') {
    block(ws, 'Итоги', cols3, [
      ...INCOME_TYPES.map((t) => [TYPE[t].many, s.countBy[t], s.incomeBy[t]]),
      ['Доходы всего', s.incomeCount, s.income],
      ['Расходы', s.expensesCount, s.expensesTotal],
      ['Остаток (доходы − расходы)', '', s.balance],
      vatRow,
    ], { boldRows: [3, 5] });

    block(ws, 'По направлениям', [
      { title: 'Направление' }, { title: 'Доходы', fmt: 'money' }, { title: 'Расходы', fmt: 'money' }, { title: 'Остаток', fmt: 'money' },
    ], s.byDirection.map((d) => [d.label, d.income, d.expenses, d.balance]), { total: true });
  } else {
    const rows = [
      [`${TYPE[scope].many}: доход`, s.incomeCount, s.income],
      ['Расходы направления', s.expensesCount, s.expensesTotal],
      ['Остаток', '', s.balance],
      vatRow,
    ];
    if (scope === 'trip') rows.push(['Средняя стоимость оплачиваемого рейса', s.paidTrips, s.avg.trip]);
    else rows.push([`Средняя сумма`, s.incomeCount, s.avg[scope]]);
    if (s.generalTotal) rows.push(['Общие расходы (входят только в итог по всем направлениям)', s.generalCount, s.generalTotal]);
    block(ws, 'Итоги', cols3, rows, { boldRows: [2] });
  }

  const pay = Object.entries(s.byPayment).filter(([, v]) => v.count);
  if (pay.length) {
    block(ws, 'По виду оплаты', [{ title: 'Оплата' }, { title: 'Записей', fmt: 'int' }, { title: 'Сумма', fmt: 'money' }],
      pay.map(([k, v]) => [PAYMENT[k], v.count, v.sum]), { total: true });
  }

  for (const t of s.incomeTypes) {
    const list = s.parties[t];
    if (!list.length) continue;
    const cols = [{ title: TYPE[t].partyCol }, { title: TYPE[t].countCol, fmt: 'int' }, { title: 'Сумма', fmt: 'money' }];
    if (t === 'trip') cols.push({ title: 'Свой груз', fmt: 'int' });
    block(ws, TYPE[t].partyTitle, cols, list.map((p) => (t === 'trip' ? [p.name, p.count, p.sum, p.own] : [p.name, p.count, p.sum])), { total: true });
  }

  if (s.byCategory.length) {
    block(ws, 'Расходы по статьям', [{ title: 'Статья' }, { title: 'Записей', fmt: 'int' }, { title: 'Сумма', fmt: 'money' }],
      s.byCategory.map((c) => [c.name, c.count, c.sum]), { total: true });
  }

  const unitName = { day: 'по дням', week: 'по неделям', month: 'по месяцам' }[s.unit];
  const colName = { day: 'День', week: 'Неделя', month: 'Месяц' }[s.unit];
  if (scope === 'all') {
    block(ws, `Динамика ${unitName}`, [
      { title: colName }, ...INCOME_TYPES.map((t) => ({ title: TYPE[t].many, fmt: 'money' })), { title: 'Расходы', fmt: 'money' },
    ], s.series.map((b) => [b.title, ...INCOME_TYPES.map((t) => b[t]), b.expenses]), { total: true });
  } else {
    block(ws, `Динамика ${unitName}`, [
      { title: colName }, { title: 'Записей', fmt: 'int' }, { title: 'Доход', fmt: 'money' }, { title: 'Расходы', fmt: 'money' },
    ], s.series.map((b) => [b.title, b.incomeCount, b.income, b.expenses]), { total: true });
  }
}

const COL_LETTER = (i) => String.fromCharCode(65 + i);

function block(ws, heading, cols, rows, opts = {}) {
  ws.addRow([heading]).font = { bold: true, size: 12, color: { argb: SIGN } };
  styleHeader(ws.addRow(cols.map((c) => c.title)), cols.length);
  const first = ws.lastRow.number + 1;
  rows.forEach((r, i) => {
    const row = ws.addRow(r);
    cols.forEach((c, j) => {
      if (c.fmt === 'money') row.getCell(j + 1).numFmt = MONEY;
    });
    for (let j = 1; j <= cols.length; j++) row.getCell(j).border = { bottom: { style: 'hair', color: { argb: LINE } } };
    if (opts.boldRows && opts.boldRows.includes(i)) row.font = { bold: true };
  });
  if (opts.total && rows.length) {
    const last = first + rows.length - 1;
    const t = ws.addRow(cols.map((c, j) => {
      if (j === 0) return 'Итого';
      if (c.fmt === 'money' || c.fmt === 'int') return { formula: `SUM(${COL_LETTER(j)}${first}:${COL_LETTER(j)}${last})` };
      return '';
    }));
    t.font = { bold: true };
    cols.forEach((c, j) => {
      if (c.fmt === 'money') t.getCell(j + 1).numFmt = MONEY;
      t.getCell(j + 1).border = { top: { style: 'thin', color: { argb: INK } } };
    });
  }
  ws.addRow([]);
}

// ---------------------------------------------------------------------
// Листы с записями
// ---------------------------------------------------------------------
function listSheet(wb, name, columns, records, totalKey, totalLabelKey) {
  const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = [{ header: '№', key: 'n', width: 6 }, ...columns.map(({ header, key, width }) => ({ header, key, width }))];
  styleHeader(ws.getRow(1), columns.length + 1);
  const sorted = [...records].sort((a, b) => a.date.localeCompare(b.date) || a.id - b.id);
  sorted.forEach((r, i) => {
    const values = { n: i + 1 };
    for (const c of columns) values[c.key] = c.value(r);
    const row = ws.addRow(values);
    for (const c of columns) {
      if (c.fmt === 'money') row.getCell(c.key).numFmt = MONEY;
      if (c.fmt === 'date') row.getCell(c.key).numFmt = DATE;
    }
  });
  if (sorted.length) {
    const last = sorted.length + 1;
    const idx = columns.findIndex((c) => c.key === totalKey) + 1; // +1 за колонку «№»
    const letter = COL_LETTER(idx);
    const t = ws.addRow({ [totalLabelKey]: 'Итого', [totalKey]: { formula: `SUM(${letter}2:${letter}${last})` } });
    t.font = { bold: true };
    t.getCell(totalKey).numFmt = MONEY;
    ws.autoFilter = { from: 'A1', to: `${COL_LETTER(columns.length)}${last}` };
  }
}

const dateCol = { header: 'Дата', key: 'date', width: 12, fmt: 'date', value: (r) => xlDate(r.date) };
const payCol = { header: 'Оплата', key: 'payment', width: 12, value: (r) => PAYMENT[r.payment] || '' };
const commentCol = (header = 'Комментарий', width = 32) => ({ header, key: 'comment', width, value: (r) => r.comment || '' });

const SHEETS = {
  trip: (wb, list) => listSheet(wb, 'Рейсы', [
    dateCol,
    { header: 'Рейс', key: 'route', width: 42, value: (r) => r.route },
    { header: 'Водитель', key: 'driver', width: 16, value: (r) => r.driver },
    { header: 'Стоимость', key: 'cost', width: 14, fmt: 'money', value: (r) => r.cost },
    payCol,
    commentCol(),
  ], list, 'cost', 'route'),

  sale: (wb, list) => listSheet(wb, 'Продажи', [
    dateCol,
    { header: 'Покупатель / за что', key: 'buyer', width: 36, value: (r) => r.buyer },
    { header: 'Сумма', key: 'amount', width: 14, fmt: 'money', value: (r) => r.amount },
    payCol,
    commentCol(),
  ], list, 'amount', 'buyer'),

  storage: (wb, list) => listSheet(wb, 'Хранение', [
    dateCol,
    { header: 'Клиент', key: 'client', width: 28, value: (r) => r.client },
    { header: 'Груз', key: 'cargo', width: 24, value: (r) => r.cargo || '' },
    { header: 'Хранение с', key: 'termFrom', width: 12, fmt: 'date', value: (r) => (r.termFrom ? xlDate(r.termFrom) : '') },
    { header: 'Хранение по', key: 'termTo', width: 12, fmt: 'date', value: (r) => (r.termTo ? xlDate(r.termTo) : '') },
    { header: 'Сумма', key: 'amount', width: 14, fmt: 'money', value: (r) => r.amount },
    payCol,
    commentCol(),
  ], list, 'amount', 'client'),

  expense: (wb, list) => listSheet(wb, 'Расходы', [
    dateCol,
    { header: 'Статья', key: 'category', width: 22, value: (r) => r.category },
    { header: 'Направление', key: 'direction', width: 14, value: (r) => DIRECTION[r.direction] || '' },
    { header: 'Сумма', key: 'amount', width: 14, fmt: 'money', value: (r) => r.amount },
    commentCol('За что', 40),
  ], list, 'amount', 'category'),
};

function styleHeader(row, n) {
  for (let i = 1; i <= n; i++) {
    const c = row.getCell(i);
    c.font = { bold: true, color: { argb: INK } };
    c.fill = HEAD_FILL;
    c.border = { bottom: { style: 'thin', color: { argb: INK } } };
    c.alignment = { vertical: 'middle' };
  }
}

function isoNow() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
