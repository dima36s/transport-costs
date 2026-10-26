// Расчёт итогов за период — используется и экраном, и отчётом.
//
// data:  { trip: [], sale: [], storage: [], expense: [] } — записи за период
// scope: 'all' — все направления вместе; 'trip' | 'sale' | 'storage' — одно направление
//        со своими доходами и своими расходами (общие расходы в него не входят)

import { buckets } from './periods.js';
import { nameKey } from './lib.js';
import { INCOME_TYPES, TYPE, amountOf } from './kinds.js';

export function compute(data, period, vatRate, scope = 'all') {
  const incomeTypes = scope === 'all' ? INCOME_TYPES : [scope];
  const expenses = scope === 'all' ? data.expense : data.expense.filter((e) => e.direction === scope);
  const general = data.expense.filter((e) => e.direction === 'general');

  const incomeBy = {};
  const countBy = {};
  for (const t of INCOME_TYPES) {
    incomeBy[t] = sumOf(t, data[t]);
    countBy[t] = data[t].length;
  }
  const income = incomeTypes.reduce((s, t) => s + incomeBy[t], 0);
  const incomeCount = incomeTypes.reduce((s, t) => s + countBy[t], 0);
  const expensesTotal = sumOf('expense', expenses);

  // оплата
  const byPayment = { bn: agg(), cash: agg(), own: agg(), '': agg() };
  for (const t of incomeTypes) {
    for (const r of data[t]) {
      const p = byPayment[r.payment] ? r.payment : '';
      byPayment[p].count++;
      byPayment[p].sum += amountOf(t, r);
    }
  }
  const rate = Number(vatRate) || 0;
  const vat = rate > 0 ? (byPayment.bn.sum * rate) / (100 + rate) : 0;

  // водители, покупатели, клиенты
  const parties = {};
  for (const t of incomeTypes) parties[t] = groupParties(t, data[t]);

  // все направления — для итога «Всё»
  const byDirection = INCOME_TYPES.map((t) => {
    const exp = data.expense.filter((e) => e.direction === t);
    const expSum = sumOf('expense', exp);
    return { key: t, label: TYPE[t].many, count: countBy[t], income: incomeBy[t], expCount: exp.length, expenses: expSum, balance: incomeBy[t] - expSum };
  });
  const generalTotal = sumOf('expense', general);
  byDirection.push({ key: 'general', label: 'Общие расходы', count: 0, income: 0, expCount: general.length, expenses: generalTotal, balance: -generalTotal });

  // статьи расходов
  const cats = new Map();
  for (const e of expenses) {
    if (!cats.has(e.category)) cats.set(e.category, { name: e.category, count: 0, sum: 0 });
    const c = cats.get(e.category);
    c.count++;
    c.sum += e.amount;
  }
  const byCategory = [...cats.values()].sort((a, b) => b.sum - a.sum);

  // средние
  const paidTrips = data.trip.filter((t) => t.cost > 0).length;
  const avg = {
    trip: paidTrips ? incomeBy.trip / paidTrips : 0,
    sale: countBy.sale ? incomeBy.sale / countBy.sale : 0,
    storage: countBy.storage ? incomeBy.storage / countBy.storage : 0,
  };

  // динамика
  const { unit, items } = buckets(period);
  const series = items.map((b) => {
    const inside = (r) => r.date >= b.from && r.date <= b.to;
    const row = { ...b, income: 0, incomeCount: 0 };
    for (const t of INCOME_TYPES) {
      const list = data[t].filter(inside);
      row[t] = sumOf(t, list);
      row[`${t}Count`] = list.length;
      if (incomeTypes.includes(t)) {
        row.income += row[t];
        row.incomeCount += list.length;
      }
    }
    row.expenses = sumOf('expense', expenses.filter(inside));
    return row;
  });

  return {
    scope,
    incomeTypes,
    incomeBy,
    countBy,
    income,
    incomeCount,
    expensesTotal,
    expensesCount: expenses.length,
    balance: income - expensesTotal,
    generalTotal: scope === 'all' ? 0 : generalTotal,
    generalCount: scope === 'all' ? 0 : general.length,
    byPayment,
    vat,
    vatRate: rate,
    parties,
    byDirection,
    byCategory,
    paidTrips,
    ownTrips: data.trip.filter((t) => t.payment === 'own').length,
    avg,
    unit,
    series,
  };
}

/** «Лёша» и «Леша» — один человек */
function groupParties(type, list) {
  const key = TYPE[type].partyKey;
  const map = new Map();
  for (const r of list) {
    const raw = String(r[key] || '').trim();
    const k = nameKey(raw) || '—';
    if (!map.has(k)) map.set(k, { names: new Map(), count: 0, sum: 0, own: 0 });
    const g = map.get(k);
    const display = raw || TYPE[type].partyEmpty;
    g.names.set(display, (g.names.get(display) || 0) + 1);
    g.count++;
    g.sum += amountOf(type, r);
    if (r.payment === 'own') g.own++;
  }
  return [...map.values()]
    .map((g) => ({
      name: [...g.names.entries()].sort((a, b) => b[1] - a[1])[0][0],
      variants: [...g.names.keys()],
      count: g.count,
      sum: g.sum,
      own: g.own,
    }))
    .sort((a, b) => b.sum - a.sum || b.count - a.count);
}

function agg() {
  return { count: 0, sum: 0 };
}
function sumOf(type, list) {
  return list.reduce((s, r) => s + amountOf(type, r), 0);
}
