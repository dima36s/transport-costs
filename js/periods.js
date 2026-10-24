// Работа с датами и периодами. Дата хранится строкой 'YYYY-MM-DD' (локальная, без часового пояса).

export const MONTHS = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
export const MONTHS_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
export const MONTHS_SHORT = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
const WEEKDAYS_SHORT = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс'];

const pad = (n) => String(n).padStart(2, '0');

export function toISO(d) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
export function fromISO(s) {
  const [y, m, d] = s.split('-').map(Number);
  return new Date(y, m - 1, d);
}
export function today() {
  return toISO(new Date());
}
export function addDays(iso, n) {
  const d = fromISO(iso);
  d.setDate(d.getDate() + n);
  return toISO(d);
}
export function formatDate(iso) {
  const [y, m, d] = iso.split('-');
  return `${d}.${m}.${y}`;
}
export function daysBetween(a, b) {
  return Math.round((fromISO(b) - fromISO(a)) / 86400000);
}

// ISO-неделя: понедельник — первый день
export function weekStart(iso) {
  const d = fromISO(iso);
  const wd = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - wd);
  return toISO(d);
}
export function isoWeekNumber(iso) {
  const d = fromISO(iso);
  const target = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const dayNr = (target.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  const firstThursday = new Date(target.getFullYear(), 0, 4);
  const diff = (target - firstThursday) / 86400000;
  return 1 + Math.round((diff - 3 + ((firstThursday.getDay() + 6) % 7)) / 7);
}

/**
 * Период: { kind: 'week'|'month'|'year'|'custom', from, to }
 */
export function periodFor(kind, anchorISO) {
  const d = fromISO(anchorISO);
  if (kind === 'week') {
    const from = weekStart(anchorISO);
    return { kind, from, to: addDays(from, 6) };
  }
  if (kind === 'month') {
    const from = toISO(new Date(d.getFullYear(), d.getMonth(), 1));
    const to = toISO(new Date(d.getFullYear(), d.getMonth() + 1, 0));
    return { kind, from, to };
  }
  if (kind === 'year') {
    return { kind, from: `${d.getFullYear()}-01-01`, to: `${d.getFullYear()}-12-31` };
  }
  throw new Error('unknown period kind ' + kind);
}

export function shiftPeriod(p, dir) {
  if (p.kind === 'week') return periodFor('week', addDays(p.from, 7 * dir));
  if (p.kind === 'month') {
    const d = fromISO(p.from);
    return periodFor('month', toISO(new Date(d.getFullYear(), d.getMonth() + dir, 1)));
  }
  if (p.kind === 'year') {
    const y = Number(p.from.slice(0, 4)) + dir;
    return periodFor('year', `${y}-01-01`);
  }
  // произвольный: сдвиг на собственную длину
  const len = daysBetween(p.from, p.to) + 1;
  return { kind: 'custom', from: addDays(p.from, len * dir), to: addDays(p.to, len * dir) };
}

export function previousPeriod(p) {
  return shiftPeriod(p, -1);
}

export function containsDate(p, iso) {
  return iso >= p.from && iso <= p.to;
}

export function periodTitle(p) {
  const a = fromISO(p.from);
  const b = fromISO(p.to);
  if (p.kind === 'month') {
    const name = MONTHS[a.getMonth()];
    return `${name[0].toUpperCase()}${name.slice(1)} ${a.getFullYear()}`;
  }
  if (p.kind === 'year') return `${a.getFullYear()} год`;
  return rangeText(a, b);
}

export function periodSubtitle(p) {
  if (p.kind === 'week') return `${isoWeekNumber(p.from)}-я неделя`;
  if (p.kind === 'custom') {
    const n = daysBetween(p.from, p.to) + 1;
    return `${n} ${plural(n, 'день', 'дня', 'дней')}`;
  }
  return '';
}

/** Название периода в родительном падеже — для фраз «итого за …» */
export function periodPhrase(p) {
  const a = fromISO(p.from);
  if (p.kind === 'month') return `${MONTHS[a.getMonth()]} ${a.getFullYear()}`;
  if (p.kind === 'year') return `${a.getFullYear()} год`;
  if (p.kind === 'week') return `неделю ${rangeText(a, fromISO(p.to))}`;
  return `период ${rangeText(a, fromISO(p.to))}`;
}

function rangeText(a, b) {
  if (a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth()) {
    return `${a.getDate()}–${b.getDate()} ${MONTHS_GEN[a.getMonth()]} ${a.getFullYear()}`;
  }
  if (a.getFullYear() === b.getFullYear()) {
    return `${a.getDate()} ${MONTHS_GEN[a.getMonth()]} – ${b.getDate()} ${MONTHS_GEN[b.getMonth()]} ${a.getFullYear()}`;
  }
  return `${formatDate(toISO(a))} – ${formatDate(toISO(b))}`;
}

export function previousLabel(p) {
  return {
    week: 'прошлой неделе',
    month: 'прошлому месяцу',
    year: 'прошлому году',
    custom: 'предыдущему периоду той же длины',
  }[p.kind];
}

/**
 * Разбиение периода на столбики графика.
 * Неделя и месяц — по дням, год — по месяцам, произвольный — по дням / неделям / месяцам по длине.
 */
export function buckets(p) {
  const len = daysBetween(p.from, p.to) + 1;
  let unit;
  if (p.kind === 'year') unit = 'month';
  else if (p.kind === 'week' || p.kind === 'month') unit = 'day';
  else unit = len <= 45 ? 'day' : len <= 190 ? 'week' : 'month';

  const out = [];
  if (unit === 'day') {
    for (let i = 0; i < len; i++) {
      const iso = addDays(p.from, i);
      const d = fromISO(iso);
      out.push({
        from: iso,
        to: iso,
        label: p.kind === 'week' ? WEEKDAYS_SHORT[(d.getDay() + 6) % 7] : String(d.getDate()),
        title: `${d.getDate()} ${MONTHS_GEN[d.getMonth()]}, ${WEEKDAYS_SHORT[(d.getDay() + 6) % 7]}`,
        weekend: d.getDay() === 0 || d.getDay() === 6,
      });
    }
  } else if (unit === 'week') {
    let cur = weekStart(p.from);
    while (cur <= p.to) {
      const end = addDays(cur, 6);
      const a = fromISO(cur);
      out.push({
        from: cur < p.from ? p.from : cur,
        to: end > p.to ? p.to : end,
        label: `${a.getDate()}.${pad(a.getMonth() + 1)}`,
        title: `Неделя с ${a.getDate()} ${MONTHS_GEN[a.getMonth()]}`,
      });
      cur = addDays(cur, 7);
    }
  } else {
    let d = fromISO(p.from);
    d = new Date(d.getFullYear(), d.getMonth(), 1);
    while (toISO(d) <= p.to) {
      const start = toISO(d);
      const end = toISO(new Date(d.getFullYear(), d.getMonth() + 1, 0));
      out.push({
        from: start < p.from ? p.from : start,
        to: end > p.to ? p.to : end,
        label: MONTHS_SHORT[d.getMonth()],
        title: `${MONTHS[d.getMonth()]} ${d.getFullYear()}`,
      });
      d = new Date(d.getFullYear(), d.getMonth() + 1, 1);
    }
  }
  return { unit, items: out };
}

export function plural(n, one, few, many) {
  const a = Math.abs(n) % 100;
  const b = a % 10;
  if (a > 10 && a < 20) return many;
  if (b > 1 && b < 5) return few;
  if (b === 1) return one;
  return many;
}
