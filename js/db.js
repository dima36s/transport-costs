// Хранилище: IndexedDB (через Dexie). Все записи живут в базе браузера.
// Слой изолирован: чтобы перейти на серверную БД, достаточно переписать этот файл.

import { TYPES, DIRECTION, guessDirection, importIdentity } from './kinds.js';

const db = new Dexie('transport-costs');

db.version(1).stores({
  trips: '++id, date, driver, payment, importKey',
  expenses: '++id, date, category, importKey',
  settings: 'key',
});

// v2: продажи, хранение и направление у расходов
db.version(2)
  .stores({
    trips: '++id, date, driver, payment, importKey',
    expenses: '++id, date, category, importKey',
    settings: 'key',
    sales: '++id, date, buyer',
    storage: '++id, date, client',
  })
  .upgrade((tx) => tx.table('expenses').toCollection().modify((e) => {
    if (!e.direction) e.direction = guessDirection(e.category);
  }));

const TABLE = { trip: 'trips', sale: 'sales', storage: 'storage', expense: 'expenses' };
const table = (type) => db[TABLE[type]];

export const PAYMENT = {
  bn: 'Безнал',
  cash: 'Нал',
  own: 'Наш груз',
  '': 'Не указано',
};

export const CATEGORIES = [
  'Топливо',
  'Платные дороги',
  'ТО и ремонт',
  'Зарплата',
  'Закупка товара',
  'Аренда и коммунальные',
  'Штрафы',
  'Налоги',
  'Офис и связь',
  'Прочее',
];

export async function init() {
  try {
    if (navigator.storage && navigator.storage.persist) {
      await navigator.storage.persist();
    }
  } catch (_) { /* не критично */ }
  await db.open();
}

// ---------- очистка полей ----------
const money = (v) => Math.round((Number(v) || 0) * 100) / 100;
const str = (v) => String(v ?? '').trim();
const isoOrEmpty = (v) => (/^\d{4}-\d{2}-\d{2}$/.test(v || '') ? v : '');
const withKey = (r) => (r.importKey ? { importKey: r.importKey } : {});

const CLEAN = {
  trip: (t) => ({
    date: t.date,
    route: str(t.route),
    driver: str(t.driver),
    cost: money(t.cost),
    payment: t.payment in PAYMENT ? t.payment : '',
    comment: str(t.comment),
    ...withKey(t),
  }),
  sale: (s) => ({
    date: s.date,
    buyer: str(s.buyer),
    amount: money(s.amount),
    payment: s.payment === 'cash' ? 'cash' : 'bn',
    comment: str(s.comment),
    ...withKey(s),
  }),
  storage: (s) => {
    let termFrom = isoOrEmpty(s.termFrom);
    let termTo = isoOrEmpty(s.termTo);
    if (termFrom && termTo && termFrom > termTo) [termFrom, termTo] = [termTo, termFrom];
    return {
      date: s.date,
      client: str(s.client),
      cargo: str(s.cargo),
      termFrom,
      termTo,
      amount: money(s.amount),
      payment: s.payment === 'cash' ? 'cash' : 'bn',
      comment: str(s.comment),
      ...withKey(s),
    };
  },
  expense: (e) => {
    const category = CATEGORIES.includes(e.category) ? e.category : 'Прочее';
    return {
      date: e.date,
      category,
      direction: e.direction in DIRECTION ? e.direction : guessDirection(category),
      amount: money(e.amount),
      comment: str(e.comment),
      ...withKey(e),
    };
  },
};

// ---------- записи любого типа ----------
export function add(type, rec) {
  return table(type).add({ ...CLEAN[type](rec), createdAt: Date.now() });
}
export function update(type, id, rec) {
  return table(type).update(id, CLEAN[type](rec));
}
export function remove(type, id) {
  return table(type).delete(id);
}
export function put(type, rec) {
  return table(type).put(rec);
}
export function all(type) {
  return table(type).toArray();
}

/** Все записи за период: { trip, sale, storage, expense } */
export async function allBetween(from, to) {
  const lists = await Promise.all(TYPES.map((t) => table(t).where('date').between(from, to, true, true).toArray()));
  const out = {};
  TYPES.forEach((t, i) => (out[t] = lists[i]));
  out.expense = out.expense.map((e) => (e.direction ? e : { ...e, direction: guessDirection(e.category) }));
  return out;
}

// ---------- справочники для подсказок ----------
export async function suggestions() {
  const [trips, sales, storage] = await Promise.all([db.trips.toArray(), db.sales.toArray(), db.storage.toArray()]);
  const count = (arr) => {
    const m = new Map();
    for (const v of arr) {
      const k = (v || '').trim();
      if (k) m.set(k, (m.get(k) || 0) + 1);
    }
    return [...m.entries()].sort((a, b) => b[1] - a[1]).map(([k]) => k);
  };
  return {
    drivers: count(trips.map((t) => t.driver)).slice(0, 50),
    routes: count(trips.map((t) => t.route)).slice(0, 200),
    buyers: count(sales.map((s) => s.buyer)).slice(0, 100),
    clients: count(storage.map((s) => s.client)).slice(0, 100),
    cargos: count(storage.map((s) => s.cargo)).slice(0, 100),
  };
}

// ---------- импорт ----------
export async function existingImportKeys() {
  const [a, b] = await Promise.all([db.trips.toArray(), db.expenses.toArray()]);
  return new Set([...a, ...b].map((r) => importIdentity(r.importKey)).filter(Boolean));
}

export async function bulkImport(trips, expenses) {
  await db.transaction('rw', db.trips, db.expenses, async () => {
    const now = Date.now();
    if (trips.length) await db.trips.bulkAdd(trips.map((t) => ({ ...CLEAN.trip(t), createdAt: now })));
    if (expenses.length) await db.expenses.bulkAdd(expenses.map((e) => ({ ...CLEAN.expense(e), createdAt: now })));
  });
}

// ---------- настройки ----------
export async function getSetting(key, fallback) {
  const row = await db.settings.get(key);
  return row ? row.value : fallback;
}
export function setSetting(key, value) {
  return db.settings.put({ key, value });
}

// ---------- резервная копия ----------
export async function exportAll() {
  const [trips, sales, storage, expenses, settings] = await Promise.all([
    db.trips.toArray(),
    db.sales.toArray(),
    db.storage.toArray(),
    db.expenses.toArray(),
    db.settings.toArray(),
  ]);
  return { app: 'transport-costs', version: 2, exportedAt: new Date().toISOString(), trips, sales, storage, expenses, settings };
}

export async function restoreAll(data) {
  if (!data || data.app !== 'transport-costs' || !Array.isArray(data.trips)) {
    throw new Error('Это не файл резервной копии этого приложения');
  }
  // копии первой версии: без продаж и хранения, у расходов нет направления
  const expenses = (data.expenses || []).map((e) => (e.direction ? e : { ...e, direction: guessDirection(e.category) }));
  await db.transaction('rw', [db.trips, db.sales, db.storage, db.expenses, db.settings], async () => {
    await Promise.all([db.trips.clear(), db.sales.clear(), db.storage.clear(), db.expenses.clear(), db.settings.clear()]);
    await db.trips.bulkPut(data.trips);
    await db.sales.bulkPut(data.sales || []);
    await db.storage.bulkPut(data.storage || []);
    await db.expenses.bulkPut(expenses);
    await db.settings.bulkPut(data.settings || []);
  });
}

export async function wipeAll() {
  await db.transaction('rw', [db.trips, db.sales, db.storage, db.expenses], async () => {
    await Promise.all([db.trips.clear(), db.sales.clear(), db.storage.clear(), db.expenses.clear()]);
  });
}

export async function counts() {
  const [trip, sale, storage, expense] = await Promise.all(TYPES.map((t) => table(t).count()));
  return { trip, sale, storage, expense };
}
