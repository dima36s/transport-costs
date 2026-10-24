// Хранилище: IndexedDB (через Dexie). Все записи живут в базе браузера.
// Слой изолирован: чтобы перейти на серверную БД, достаточно переписать этот файл.

const db = new Dexie('transport-costs');

db.version(1).stores({
  trips: '++id, date, driver, payment, importKey',
  expenses: '++id, date, category, importKey',
  settings: 'key',
});

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

// ---------- рейсы ----------
export function addTrip(t) {
  return db.trips.add({ ...cleanTrip(t), createdAt: Date.now() });
}
export function updateTrip(id, t) {
  return db.trips.update(id, cleanTrip(t));
}
export function deleteTrip(id) {
  return db.trips.delete(id);
}
export function putTrip(t) {
  return db.trips.put(t);
}
export function tripsBetween(from, to) {
  return db.trips.where('date').between(from, to, true, true).toArray();
}
export function allTrips() {
  return db.trips.toArray();
}

function cleanTrip(t) {
  return {
    date: t.date,
    route: (t.route || '').trim(),
    driver: (t.driver || '').trim(),
    cost: Math.round((Number(t.cost) || 0) * 100) / 100,
    payment: t.payment in PAYMENT ? t.payment : '',
    comment: (t.comment || '').trim(),
    ...(t.importKey ? { importKey: t.importKey } : {}),
  };
}

// ---------- расходы ----------
export function addExpense(e) {
  return db.expenses.add({ ...cleanExpense(e), createdAt: Date.now() });
}
export function updateExpense(id, e) {
  return db.expenses.update(id, cleanExpense(e));
}
export function deleteExpense(id) {
  return db.expenses.delete(id);
}
export function putExpense(e) {
  return db.expenses.put(e);
}
export function expensesBetween(from, to) {
  return db.expenses.where('date').between(from, to, true, true).toArray();
}
export function allExpenses() {
  return db.expenses.toArray();
}

function cleanExpense(e) {
  return {
    date: e.date,
    category: CATEGORIES.includes(e.category) ? e.category : 'Прочее',
    amount: Math.round((Number(e.amount) || 0) * 100) / 100,
    comment: (e.comment || '').trim(),
    ...(e.importKey ? { importKey: e.importKey } : {}),
  };
}

// ---------- справочники для подсказок ----------
export async function suggestions() {
  const trips = await db.trips.toArray();
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
  };
}

// ---------- импорт ----------
export async function existingImportKeys() {
  const [a, b] = await Promise.all([db.trips.toArray(), db.expenses.toArray()]);
  return new Set([...a, ...b].map((r) => r.importKey).filter(Boolean));
}

export async function bulkImport(trips, expenses) {
  await db.transaction('rw', db.trips, db.expenses, async () => {
    const now = Date.now();
    if (trips.length) await db.trips.bulkAdd(trips.map((t) => ({ ...cleanTrip(t), createdAt: now })));
    if (expenses.length) await db.expenses.bulkAdd(expenses.map((e) => ({ ...cleanExpense(e), createdAt: now })));
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
  const [trips, expenses, settings] = await Promise.all([
    db.trips.toArray(),
    db.expenses.toArray(),
    db.settings.toArray(),
  ]);
  return { app: 'transport-costs', version: 1, exportedAt: new Date().toISOString(), trips, expenses, settings };
}

export async function restoreAll(data) {
  if (!data || data.app !== 'transport-costs' || !Array.isArray(data.trips)) {
    throw new Error('Это не файл резервной копии этого приложения');
  }
  await db.transaction('rw', db.trips, db.expenses, db.settings, async () => {
    await Promise.all([db.trips.clear(), db.expenses.clear(), db.settings.clear()]);
    await db.trips.bulkPut(data.trips);
    await db.expenses.bulkPut(data.expenses || []);
    await db.settings.bulkPut(data.settings || []);
  });
}

export async function wipeAll() {
  await db.transaction('rw', db.trips, db.expenses, async () => {
    await Promise.all([db.trips.clear(), db.expenses.clear()]);
  });
}

export async function counts() {
  const [t, e] = await Promise.all([db.trips.count(), db.expenses.count()]);
  return { trips: t, expenses: e };
}
