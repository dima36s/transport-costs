// Отчёт за период в Excel: «Сводка», «Рейсы», «Расходы».

import { loadExcelJS, downloadBlob } from './lib.js';
import { PAYMENT } from './db.js';
import { formatDate, periodTitle, periodSubtitle } from './periods.js';

// ExcelJS пишет даты по UTC — иначе в Москве дата съезжает на день назад
const xlDate = (iso) => { const [y, m, d] = iso.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };

const MONEY = '#,##0 "₽";[Red]-#,##0 "₽"';
const INK = 'FF1A2128';
const SIGN = 'FF1B4F9C';
const LINE = 'FFD3D9DF';
const HEAD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EDF3' } };

export async function buildReport({ period, trips, expenses, stats }) {
  const ExcelJS = await loadExcelJS();
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Учёт рейсов';
  wb.created = new Date();

  const title = periodTitle(period);
  const sub = periodSubtitle(period);
  const range = `${formatDate(period.from)} – ${formatDate(period.to)}`;

  summarySheet(wb, { title, sub, range, stats });
  tripsSheet(wb, trips);
  expensesSheet(wb, expenses);

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  downloadBlob(blob, `Отчёт по рейсам ${period.from} — ${period.to}.xlsx`);
}

function summarySheet(wb, { title, sub, range, stats }) {
  const ws = wb.addWorksheet('Сводка', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 34 }, { width: 14 }, { width: 18 }, { width: 14 }];

  ws.addRow([`Отчёт по рейсам и расходам: ${title}`]).font = { size: 16, bold: true, color: { argb: INK } };
  ws.addRow([`${range}${sub ? `, ${sub}` : ''}`]).font = { color: { argb: 'FF5B6670' } };
  ws.addRow([`Сформирован ${formatDate(isoNow())}`]).font = { color: { argb: 'FF5B6670' }, size: 9 };
  ws.addRow([]);

  block(ws, 'Итоги', ['Показатель', '', 'Сумма'], [
    ['Сумма рейсов', stats.tripsCount, stats.tripsTotal],
    ['Расходы', stats.expensesCount, stats.expensesTotal],
    ['Остаток (рейсы − расходы)', '', stats.balance],
    [`в т.ч. НДС ${stats.vatRate}% с безналичных рейсов`, '', stats.vat],
    ['Средняя стоимость оплачиваемого рейса', stats.paidTrips, stats.avgTrip],
  ], { boldFirst: true, countHeader: 'Кол-во' });

  block(ws, 'По виду оплаты', ['Оплата', 'Рейсов', 'Сумма'],
    Object.entries(stats.byPayment)
      .filter(([, v]) => v.count)
      .map(([k, v]) => [PAYMENT[k], v.count, v.sum]),
    { total: true });

  block(ws, 'По водителям', ['Водитель', 'Рейсов', 'Сумма', 'Свой груз'],
    stats.byDriver.map((d) => [d.name, d.count, d.sum, d.own]),
    { total: true });

  if (stats.byCategory.length) {
    block(ws, 'Расходы по статьям', ['Статья', 'Записей', 'Сумма'],
      stats.byCategory.map((c) => [c.name, c.count, c.sum]), { total: true });
  }

  const unitName = { day: 'По дням', week: 'По неделям', month: 'По месяцам' }[stats.unit];
  const colName = { day: 'День', week: 'Неделя', month: 'Месяц' }[stats.unit];
  block(ws, `Динамика: ${unitName.toLowerCase()}`, [colName, 'Рейсов', 'Сумма рейсов', 'Расходы'],
    stats.series.map((s) => [s.title, s.tripsCount, s.trips, s.expenses]), { total: true });
}

function block(ws, heading, header, rows, opts = {}) {
  const h = ws.addRow([heading]);
  h.font = { bold: true, size: 12, color: { argb: SIGN } };
  const hr = ws.addRow(opts.countHeader ? [header[0], opts.countHeader, header[2]] : header);
  styleHeader(hr, header.length);
  const first = hr.number + 1;
  for (const r of rows) {
    const row = ws.addRow(r);
    row.getCell(3).numFmt = MONEY;
    if (header[3] === 'Расходы') row.getCell(4).numFmt = MONEY;
    row.eachCell((c) => (c.border = { bottom: { style: 'hair', color: { argb: LINE } } }));
  }
  if (opts.total && rows.length) {
    const last = first + rows.length - 1;
    const t = ws.addRow(['Итого', { formula: `SUM(B${first}:B${last})` }, { formula: `SUM(C${first}:C${last})` }]);
    if (header.length > 3) {
      t.getCell(4).value = { formula: `SUM(D${first}:D${last})` };
      if (header[3] === 'Расходы') t.getCell(4).numFmt = MONEY;
    }
    t.font = { bold: true };
    t.getCell(3).numFmt = MONEY;
    t.eachCell((c) => (c.border = { top: { style: 'thin', color: { argb: INK } } }));
  }
  ws.addRow([]);
}

function tripsSheet(wb, trips) {
  const ws = wb.addWorksheet('Рейсы', { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = [
    { header: '№', key: 'n', width: 6 },
    { header: 'Дата', key: 'date', width: 12 },
    { header: 'Рейс', key: 'route', width: 42 },
    { header: 'Водитель', key: 'driver', width: 16 },
    { header: 'Стоимость', key: 'cost', width: 14 },
    { header: 'Оплата', key: 'payment', width: 12 },
    { header: 'Комментарий', key: 'comment', width: 30 },
  ];
  styleHeader(ws.getRow(1), 7);
  const sorted = [...trips].sort((a, b) => a.date.localeCompare(b.date) || a.id - b.id);
  sorted.forEach((t, i) => {
    const row = ws.addRow({
      n: i + 1,
      date: xlDate(t.date),
      route: t.route,
      driver: t.driver,
      cost: t.cost,
      payment: PAYMENT[t.payment] || '',
      comment: t.comment,
    });
    row.getCell('date').numFmt = 'dd.mm.yyyy';
    row.getCell('cost').numFmt = MONEY;
  });
  if (sorted.length) {
    const last = sorted.length + 1;
    const t = ws.addRow({ route: 'Итого', cost: { formula: `SUM(E2:E${last})` } });
    t.font = { bold: true };
    t.getCell('cost').numFmt = MONEY;
    ws.autoFilter = { from: 'A1', to: `G${last}` };
  }
}

function expensesSheet(wb, expenses) {
  const ws = wb.addWorksheet('Расходы', { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = [
    { header: '№', key: 'n', width: 6 },
    { header: 'Дата', key: 'date', width: 12 },
    { header: 'Статья', key: 'category', width: 18 },
    { header: 'Сумма', key: 'amount', width: 14 },
    { header: 'Комментарий', key: 'comment', width: 40 },
  ];
  styleHeader(ws.getRow(1), 5);
  const sorted = [...expenses].sort((a, b) => a.date.localeCompare(b.date) || a.id - b.id);
  sorted.forEach((e, i) => {
    const row = ws.addRow({ n: i + 1, date: xlDate(e.date), category: e.category, amount: e.amount, comment: e.comment });
    row.getCell('date').numFmt = 'dd.mm.yyyy';
    row.getCell('amount').numFmt = MONEY;
  });
  if (sorted.length) {
    const last = sorted.length + 1;
    const t = ws.addRow({ category: 'Итого', amount: { formula: `SUM(D2:D${last})` } });
    t.font = { bold: true };
    t.getCell('amount').numFmt = MONEY;
    ws.autoFilter = { from: 'A1', to: `E${last}` };
  }
}

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
