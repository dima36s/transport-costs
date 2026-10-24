// Расчёт итогов за период — используется и экраном, и отчётом.

import { buckets } from './periods.js';
import { nameKey } from './lib.js';

export function compute(trips, expenses, period, vatRate) {
  const tripsTotal = sum(trips, 'cost');
  const expensesTotal = sum(expenses, 'amount');

  const byPayment = { bn: agg(), cash: agg(), own: agg(), '': agg() };
  for (const t of trips) {
    const p = byPayment[t.payment] ? t.payment : '';
    byPayment[p].count++;
    byPayment[p].sum += t.cost;
  }

  const rate = Number(vatRate) || 0;
  const vat = rate > 0 ? (byPayment.bn.sum * rate) / (100 + rate) : 0;

  // водители: «Лёша» и «Леша» — один человек
  const drivers = new Map();
  for (const t of trips) {
    const k = nameKey(t.driver) || '—';
    if (!drivers.has(k)) drivers.set(k, { names: new Map(), count: 0, sum: 0, own: 0 });
    const d = drivers.get(k);
    const display = t.driver.trim() || 'Без водителя';
    d.names.set(display, (d.names.get(display) || 0) + 1);
    d.count++;
    d.sum += t.cost;
    if (t.payment === 'own') d.own++;
  }
  const byDriver = [...drivers.values()]
    .map((d) => ({
      name: [...d.names.entries()].sort((a, b) => b[1] - a[1])[0][0],
      variants: [...d.names.keys()],
      count: d.count,
      sum: d.sum,
      own: d.own,
    }))
    .sort((a, b) => b.sum - a.sum || b.count - a.count);

  const cats = new Map();
  for (const e of expenses) {
    if (!cats.has(e.category)) cats.set(e.category, { name: e.category, count: 0, sum: 0 });
    const c = cats.get(e.category);
    c.count++;
    c.sum += e.amount;
  }
  const byCategory = [...cats.values()].sort((a, b) => b.sum - a.sum);

  const { unit, items } = buckets(period);
  const series = items.map((b) => ({
    ...b,
    trips: sum(trips.filter((t) => t.date >= b.from && t.date <= b.to), 'cost'),
    tripsCount: trips.filter((t) => t.date >= b.from && t.date <= b.to).length,
    expenses: sum(expenses.filter((e) => e.date >= b.from && e.date <= b.to), 'amount'),
  }));

  return {
    tripsTotal,
    tripsCount: trips.length,
    expensesTotal,
    expensesCount: expenses.length,
    balance: tripsTotal - expensesTotal,
    paidTrips: trips.filter((t) => t.cost > 0).length,
    avgTrip: trips.filter((t) => t.cost > 0).length ? tripsTotal / trips.filter((t) => t.cost > 0).length : 0,
    byPayment,
    vat,
    vatRate: rate,
    byDriver,
    byCategory,
    unit,
    series,
  };
}

function agg() {
  return { count: 0, sum: 0 };
}
function sum(arr, key) {
  return arr.reduce((s, x) => s + (Number(x[key]) || 0), 0);
}
