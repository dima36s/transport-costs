import * as db from './db.js';
import { PAYMENT, CATEGORIES } from './db.js';
import * as P from './periods.js';
import { compute } from './stats.js';
import { money, escapeHTML, downloadBlob, nameKey } from './lib.js';

const $ = (s, root = document) => root.querySelector(s);
const $$ = (s, root = document) => [...root.querySelectorAll(s)];

const state = {
  kind: 'month',
  period: P.periodFor('month', P.today()),
  customFrom: null,
  customTo: null,
  tab: 'trips',
  search: '',
  formType: 'trip',
  editing: null, // { type: 'trip'|'expense', id }
  vatRate: 22,
  data: { trips: [], expenses: [], stats: null },
  highlight: null, // { type, id }
  printing: false,
};

// ======================================================================
// Запуск
// ======================================================================
async function start() {
  try {
    await db.init();
  } catch (err) {
    showFatal('Браузер не дал доступ к базе данных. Откройте приложение в обычном (не приватном) окне Chrome, Edge, Firefox или Safari.');
    console.error(err);
    return;
  }
  state.vatRate = Number(await db.getSetting('vatRate', 22));
  const saved = loadView();
  if (saved) Object.assign(state, saved);

  fillCategories();
  resetTripForm();
  resetExpenseForm();
  bindForms();
  bindPeriod();
  bindJournal();
  bindDialogs();
  bindSheet();
  window.addEventListener('beforeprint', () => { state.printing = true; renderJournal(); });
  window.addEventListener('afterprint', () => { state.printing = false; renderJournal(); });

  await refreshSuggestions();
  await render();

  // Пустая база — подсказываем, что можно загрузить журнал
  const c = await db.counts();
  if (!c.trips && !c.expenses) {
    $('#reportNote').textContent = 'База пока пустая. Добавьте первый рейс в форме «Новая запись» или загрузите свой журнал из Excel.';
  }
}

function showFatal(msg) {
  document.querySelector('main').innerHTML = `<div class="panel" style="grid-column:1/-1"><h2 class="h2">База данных недоступна</h2><p>${escapeHTML(msg)}</p></div>`;
}

function saveView() {
  try {
    localStorage.setItem('tc-view', JSON.stringify({ kind: state.kind, customFrom: state.customFrom, customTo: state.customTo }));
  } catch (_) { /* приватный режим */ }
}
function loadView() {
  try {
    const v = JSON.parse(localStorage.getItem('tc-view') || 'null');
    if (!v || !['week', 'month', 'year', 'custom'].includes(v.kind)) return null;
    const out = { kind: v.kind, customFrom: v.customFrom, customTo: v.customTo };
    if (v.kind === 'custom' && v.customFrom && v.customTo) out.period = { kind: 'custom', from: v.customFrom, to: v.customTo };
    else if (v.kind !== 'custom') out.period = P.periodFor(v.kind, P.today());
    else return { kind: 'month' };
    return out;
  } catch (_) {
    return null;
  }
}

// ======================================================================
// Отрисовка
// ======================================================================
async function render({ flash = false } = {}) {
  const p = state.period;
  const prev = P.previousPeriod(p);
  const [trips, expenses, prevTrips, prevExp] = await Promise.all([
    db.tripsBetween(p.from, p.to),
    db.expensesBetween(p.from, p.to),
    db.tripsBetween(prev.from, prev.to),
    db.expensesBetween(prev.from, prev.to),
  ]);
  const stats = compute(trips, expenses, p, state.vatRate);
  const prevStats = { tripsTotal: prevTrips.reduce((s, t) => s + t.cost, 0), expensesTotal: prevExp.reduce((s, e) => s + e.amount, 0) };
  state.data = { trips, expenses, stats, prevStats };

  renderPeriodBar();
  renderSign(stats, flash);
  renderFacts(stats, prevStats);
  renderChart(stats);
  renderDrivers(stats);
  renderPayments(stats);
  renderCategories(stats);
  renderJournal();
  $('#btnReport').textContent = `Скачать отчёт за ${P.periodPhrase(p)} в Excel`;
}

function renderPeriodBar() {
  $$('.periodbar [data-kind]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.kind === state.kind)));
  $('#periodTitle').textContent = P.periodTitle(state.period);
  $('#periodSub').textContent = P.periodSubtitle(state.period);
  $('#customRange').hidden = state.kind !== 'custom';
  if (state.kind === 'custom') {
    $('#customFrom').value = state.period.from;
    $('#customTo').value = state.period.to;
  }
  $('#toToday').hidden = P.containsDate(state.period, P.today());
}

function renderSign(s, flash) {
  $('#sTrips').textContent = money(s.tripsTotal);
  $('#sExpenses').textContent = money(s.expensesTotal);
  const bal = $('#sBalance');
  bal.textContent = (s.balance < 0 ? '−' : '') + money(Math.abs(s.balance));
  bal.classList.toggle('is-negative', s.balance < 0);
  if (flash) {
    const sign = $('#sign');
    sign.classList.remove('is-updated');
    void sign.offsetWidth;
    sign.classList.add('is-updated');
  }
}

function renderFacts(s, prev) {
  const items = [];
  const own = s.byPayment.own.count;
  items.push(`Рейсов: <b>${s.tripsCount}</b>${own ? `, из них со своим грузом <b>${own}</b>` : ''}`);
  if (s.paidTrips) items.push(`Средний рейс: <b>${money(s.avgTrip)}</b>`);
  if (s.vatRate > 0 && s.byPayment.bn.sum) items.push(`В т.ч. НДС ${String(s.vatRate).replace('.', ',')}% с безнала: <b>${money(s.vat)}</b>`);
  if (prev.tripsTotal > 0) {
    const d = ((s.tripsTotal - prev.tripsTotal) / prev.tripsTotal) * 100;
    const cls = d >= 0 ? 'delta--up' : 'delta--down';
    const sign = d > 0 ? '+' : d < 0 ? '−' : '';
    items.push(`Рейсы к ${P.previousLabel(state.period)}: <b class="${cls}">${sign}${Math.abs(d).toFixed(Math.abs(d) < 10 ? 1 : 0).replace('.', ',')}%</b>`);
  } else if (s.tripsTotal > 0) {
    items.push(`В предыдущем периоде рейсов не было`);
  }
  $('#facts').innerHTML = items.map((i) => `<li>${i}</li>`).join('');
}

// ---------- график ----------
function shortMoney(n) {
  if (n >= 1e6) return `${(n / 1e6).toLocaleString('ru-RU', { maximumFractionDigits: 2 })} млн`;
  if (n >= 1e3) return `${Math.round(n / 1e3).toLocaleString('ru-RU')} тыс`;
  return String(Math.round(n));
}
function niceStep(max, ticks = 4) {
  const raw = max / ticks;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const f = raw / mag;
  const step = f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10;
  return step * mag;
}

function renderChart(s) {
  const unitTitle = { day: 'По дням', week: 'По неделям', month: 'По месяцам' }[s.unit];
  $('#chartTitle').textContent = unitTitle;
  $('#legend').innerHTML = `
    <li><i style="background:var(--series-trips)"></i>Рейсы <b>${money(s.tripsTotal)}</b></li>
    <li><i style="background:var(--series-exp)"></i>Расходы <b>${money(s.expensesTotal)}</b></li>`;

  const host = $('#chart');
  const series = s.series;
  const max = Math.max(0, ...series.map((b) => Math.max(b.trips, b.expenses)));
  if (max === 0) {
    host.innerHTML = `<p class="chart__empty">За этот период записей нет. Добавьте рейс или расход в форме «Новая запись».</p>`;
    return;
  }

  const W = Math.max(320, Math.round(host.clientWidth || 700));
  const H = W < 520 ? 200 : 240;
  const m = { t: 10, r: 4, b: 26, l: 58 };
  const iw = W - m.l - m.r;
  const ih = H - m.t - m.b;
  const step = niceStep(max);
  const top = Math.ceil(max / step) * step;
  const y = (v) => m.t + ih - (v / top) * ih;
  const n = series.length;
  const slot = iw / n;
  const gap = 2;
  const barW = Math.max(2, Math.min(28, (slot * 0.72 - gap) / 2));
  const r = Math.min(4, barW / 2);

  // показываем не все подписи, если столбиков много
  const every = n <= 12 ? 1 : n <= 16 ? 2 : W < 520 ? 7 : 3;

  let g = '<g class="grid">';
  for (let v = 0; v <= top + 1e-6; v += step) {
    const yy = Math.round(y(v)) + 0.5;
    g += `<line x1="${m.l}" x2="${W - m.r}" y1="${yy}" y2="${yy}" class="${v === 0 ? 'base' : ''}"/>`;
    g += `<text class="tick" x="${m.l - 8}" y="${yy + 4}" text-anchor="end">${v === 0 ? '0' : shortMoney(v)}</text>`;
  }
  g += '</g>';

  let bars = '';
  let labels = '';
  let hits = '';
  series.forEach((b, i) => {
    const cx = m.l + slot * i + slot / 2;
    const x1 = cx - gap / 2 - barW;
    const x2 = cx + gap / 2;
    if (b.trips > 0) bars += barPath('bar-trips', x1, y(b.trips), barW, y(0) - y(b.trips), r);
    if (b.expenses > 0) bars += barPath('bar-exp', x2, y(b.expenses), barW, y(0) - y(b.expenses), r);
    const showLabel = i % every === 0 || (s.unit === 'day' && n > 16 && i === n - 1 && (n - 1) % every > every / 2);
    if (showLabel) labels += `<text class="xlabel${b.weekend ? ' weekend' : ''}" x="${cx}" y="${H - 8}">${escapeHTML(b.label)}</text>`;
    hits += `<rect class="hit" data-i="${i}" x="${m.l + slot * i}" y="${m.t}" width="${slot}" height="${ih}" tabindex="-1"/>`;
  });

  host.innerHTML = `
    <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${unitTitle}: рейсы и расходы за период">
      ${g}<g>${bars}</g><g>${labels}</g><g>${hits}</g>
    </svg>
    <table class="sr-only"><caption>${unitTitle}</caption>
      <tr><th>Период</th><th>Рейсы</th><th>Расходы</th></tr>
      ${series.map((b) => `<tr><td>${escapeHTML(b.title)}</td><td>${money(b.trips)}</td><td>${money(b.expenses)}</td></tr>`).join('')}
    </table>`;

  const tip = $('#tooltip');
  $$('.hit', host).forEach((h) => {
    const show = (ev) => {
      const b = series[Number(h.dataset.i)];
      tip.innerHTML = `<b>${escapeHTML(b.title)}</b>
        <div class="tt-row"><i style="background:var(--series-trips)"></i>Рейсы: ${money(b.trips)}${b.tripsCount ? ` (${b.tripsCount})` : ''}</div>
        <div class="tt-row"><i style="background:var(--series-exp)"></i>Расходы: ${money(b.expenses)}</div>`;
      tip.hidden = false;
      const rect = h.getBoundingClientRect();
      const tw = tip.offsetWidth;
      let left = rect.left + rect.width / 2 - tw / 2;
      left = Math.max(8, Math.min(window.innerWidth - tw - 8, left));
      let topPx = (ev && ev.clientY ? ev.clientY : rect.top) - tip.offsetHeight - 14;
      if (topPx < 8) topPx = rect.bottom + 8;
      tip.style.left = `${left}px`;
      tip.style.top = `${topPx}px`;
      h.classList.add('is-active');
    };
    const hide = () => { tip.hidden = true; h.classList.remove('is-active'); };
    h.addEventListener('mousemove', show);
    h.addEventListener('mouseleave', hide);
    h.addEventListener('touchstart', (e) => { show(e.touches[0]); setTimeout(hide, 2500); }, { passive: true });
  });
}

function barPath(cls, x, yTop, w, h, r) {
  if (h < 0.5) return '';
  const rr = Math.min(r, h);
  return `<path class="${cls}" d="M${x},${yTop + h} V${yTop + rr} Q${x},${yTop} ${x + rr},${yTop} H${x + w - rr} Q${x + w},${yTop} ${x + w},${yTop + rr} V${yTop + h} Z"/>`;
}

let resizeTimer;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => state.data.stats && renderChart(state.data.stats), 150);
});

// ---------- сводки ----------
function renderDrivers(s) {
  const host = $('#drivers');
  if (!s.byDriver.length) {
    host.innerHTML = '<p class="empty">Рейсов за период нет.</p>';
    return;
  }
  const max = Math.max(...s.byDriver.map((d) => d.sum), 1);
  host.innerHTML = `<table class="mini">
    <thead><tr><th>Водитель</th><th class="num">Рейсов</th><th class="share"><span class="sr-only">Доля</span></th><th class="num">Сумма</th></tr></thead>
    <tbody>${s.byDriver.map((d) => `<tr>
      <td>${escapeHTML(d.name)}${d.variants.length > 1 ? `<span class="variants">также: ${escapeHTML(d.variants.filter((v) => v !== d.name).join(', '))}</span>` : ''}</td>
      <td class="num">${d.count}${d.own ? ` <span class="variants" style="display:inline">(${d.own} свой)</span>` : ''}</td>
      <td class="share"><div class="sharebar"><span style="width:${(d.sum / max) * 100}%"></span></div></td>
      <td class="num">${money(d.sum)}</td></tr>`).join('')}</tbody>
  </table>`;
}

function renderPayments(s) {
  const rows = Object.entries(s.byPayment).filter(([, v]) => v.count);
  const host = $('#payments');
  if (!rows.length) {
    host.innerHTML = '<p class="empty">Рейсов за период нет.</p>';
    return;
  }
  host.innerHTML = `<table class="mini"><tbody>${rows
    .map(([k, v]) => `<tr><td>${PAYMENT[k]}</td><td class="num">${v.count} ${P.plural(v.count, 'рейс', 'рейса', 'рейсов')}</td><td class="num">${money(v.sum)}</td></tr>`)
    .join('')}</tbody></table>`;
}

function renderCategories(s) {
  const host = $('#categories');
  if (!s.byCategory.length) {
    host.innerHTML = '<p class="empty">Расходов за период нет.</p>';
    return;
  }
  const max = Math.max(...s.byCategory.map((c) => c.sum), 1);
  host.innerHTML = `<table class="mini"><tbody>${s.byCategory
    .map((c) => `<tr><td>${escapeHTML(c.name)}</td><td class="share"><div class="sharebar sharebar--exp"><span style="width:${(c.sum / max) * 100}%"></span></div></td><td class="num">${money(c.sum)}</td></tr>`)
    .join('')}</tbody></table>`;
}

// ---------- журнал ----------
function matches(item, q) {
  if (!q) return true;
  const hay = nameKey([item.route, item.driver, item.comment, item.category, PAYMENT[item.payment]].filter(Boolean).join(' '));
  return q.split(/\s+/).every((w) => hay.includes(w));
}

function renderJournal() {
  const { trips, expenses } = state.data;
  $('#cntTrips').textContent = trips.length;
  $('#cntExpenses').textContent = expenses.length;
  $$('.panel [data-tab]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === state.tab)));
  const q = nameKey(state.search);
  const host = $('#journal');
  const parts = [];
  if (state.printing || state.tab === 'trips') parts.push(tripsTable(trips.filter((t) => matches(t, q)), q));
  if (state.printing || state.tab === 'expenses') parts.push(expensesTable(expenses.filter((e) => matches(e, q)), q));
  host.innerHTML = parts.join('');

  if (state.highlight) {
    const row = host.querySelector(`tr[data-type="${state.highlight.type}"][data-id="${state.highlight.id}"]`);
    if (row) row.classList.add('is-new');
    state.highlight = null;
  }
}

const byDateDesc = (a, b) => b.date.localeCompare(a.date) || (b.createdAt || 0) - (a.createdAt || 0) || b.id - a.id;

function tripsTable(list, q) {
  if (!list.length) {
    return `<p class="journal__empty">${q ? 'Ничего не найдено.' : 'Рейсов за этот период нет.'}</p>`;
  }
  const sorted = [...list].sort(byDateDesc);
  const total = sorted.reduce((s, t) => s + t.cost, 0);
  return `<table class="journal">
    ${state.printing ? '<caption class="h2" style="text-align:left;padding:6px 0">Рейсы</caption>' : ''}
    <thead><tr><th>Дата</th><th>Рейс</th><th>Водитель</th><th class="num">Стоимость</th><th>Оплата</th><th>Комментарий</th><th class="acts"><span class="sr-only">Действия</span></th></tr></thead>
    <tbody>${sorted.map((t) => `<tr data-type="trip" data-id="${t.id}" class="${isEditing('trip', t.id) ? 'is-editing' : ''}">
      <td class="date">${P.formatDate(t.date)}</td>
      <td class="route">${escapeHTML(t.route)}</td>
      <td class="driver">${escapeHTML(t.driver)}</td>
      <td class="num">${money(t.cost)}</td>
      <td class="pay">${t.payment ? `<span class="tag tag--${t.payment}">${PAYMENT[t.payment]}</span>` : '<span class="tag">не указано</span>'}</td>
      <td class="comment">${escapeHTML(t.comment)}</td>
      <td class="acts"><button type="button" class="linkbtn" data-edit>Изменить</button><button type="button" class="linkbtn linkbtn--danger" data-del>Удалить</button></td>
    </tr>`).join('')}</tbody>
    <tfoot><tr><td></td><td>Итого${q ? ' по найденным' : ''}: ${sorted.length} ${P.plural(sorted.length, 'рейс', 'рейса', 'рейсов')}</td><td></td><td class="num">${money(total)}</td><td colspan="3"></td></tr></tfoot>
  </table>`;
}

function expensesTable(list, q) {
  if (!list.length) {
    return `<p class="journal__empty">${q ? 'Ничего не найдено.' : 'Расходов за этот период нет.'}</p>`;
  }
  const sorted = [...list].sort(byDateDesc);
  const total = sorted.reduce((s, e) => s + e.amount, 0);
  return `<table class="journal">
    ${state.printing ? '<caption class="h2" style="text-align:left;padding:6px 0">Расходы</caption>' : ''}
    <thead><tr><th>Дата</th><th>Статья</th><th class="num">Сумма</th><th>За что</th><th class="acts"><span class="sr-only">Действия</span></th></tr></thead>
    <tbody>${sorted.map((e) => `<tr data-type="expense" data-id="${e.id}" class="${isEditing('expense', e.id) ? 'is-editing' : ''}">
      <td class="date">${P.formatDate(e.date)}</td>
      <td class="route">${escapeHTML(e.category)}</td>
      <td class="num">${money(e.amount)}</td>
      <td class="comment">${escapeHTML(e.comment)}</td>
      <td class="acts"><button type="button" class="linkbtn" data-edit>Изменить</button><button type="button" class="linkbtn linkbtn--danger" data-del>Удалить</button></td>
    </tr>`).join('')}</tbody>
    <tfoot><tr><td></td><td>Итого${q ? ' по найденным' : ''}</td><td class="num">${money(total)}</td><td colspan="2"></td></tr></tfoot>
  </table>`;
}

const isEditing = (type, id) => state.editing && state.editing.type === type && state.editing.id === id;

// ======================================================================
// Формы
// ======================================================================
function fillCategories() {
  $('#eCategory').innerHTML = CATEGORIES.map((c) => `<option>${c}</option>`).join('');
}

function parseMoney(s) {
  const t = String(s ?? '').replace(/\s/g, '').replace(/₽|р\.?|руб\.?/gi, '').replace(',', '.');
  if (t === '') return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : NaN;
}

function defaultEntryDate() {
  // если смотрим прошлый период — по умолчанию ставим дату внутри него
  const t = P.today();
  return P.containsDate(state.period, t) ? t : state.period.to < t ? state.period.to : state.period.from;
}

function resetTripForm(keep = {}) {
  const f = $('#tripForm');
  f.reset();
  f.date.value = keep.date || defaultEntryDate();
  if (keep.driver) f.driver.value = keep.driver;
  if (keep.payment) f.querySelector(`input[name="payment"][value="${keep.payment}"]`).checked = true;
  setError(f, '');
  updateCostHint();
  updateDriverChips();
}
function resetExpenseForm(keep = {}) {
  const f = $('#expenseForm');
  f.reset();
  f.date.value = keep.date || defaultEntryDate();
  if (keep.category) f.category.value = keep.category;
  setError(f, '');
}

function setError(form, msg, field) {
  const el = form.querySelector('.form__error');
  el.textContent = msg;
  el.hidden = !msg;
  form.querySelectorAll('[aria-invalid]').forEach((i) => i.removeAttribute('aria-invalid'));
  if (field) {
    field.setAttribute('aria-invalid', 'true');
    field.focus();
  }
}

function switchForm(type) {
  state.formType = type;
  $$('.entry [data-form]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.form === type)));
  $('#tripForm').hidden = type !== 'trip';
  $('#expenseForm').hidden = type !== 'expense';
}

function updateCostHint() {
  const f = $('#tripForm');
  const cost = parseMoney(f.cost.value);
  const pay = f.querySelector('input[name="payment"]:checked')?.value;
  const hint = $('#tCostHint');
  if (pay === 'own' && !f.cost.value) {
    hint.textContent = 'Свой груз: стоимость можно оставить пустой — запишется 0';
  } else if (pay === 'bn' && cost > 0 && state.vatRate > 0) {
    hint.textContent = `в т.ч. НДС ${String(state.vatRate).replace('.', ',')}%: ${money((cost * state.vatRate) / (100 + state.vatRate))}`;
  } else {
    hint.textContent = '';
  }
}

let suggest = { drivers: [], routes: [] };
async function refreshSuggestions() {
  suggest = await db.suggestions();
  $('#driversList').innerHTML = suggest.drivers.map((d) => `<option value="${escapeHTML(d)}">`).join('');
  $('#routesList').innerHTML = suggest.routes.map((r) => `<option value="${escapeHTML(r)}">`).join('');
  updateDriverChips();
}
function updateDriverChips() {
  const f = $('#tripForm');
  const cur = nameKey(f.driver.value);
  const seen = new Set();
  const top = [];
  for (const d of suggest.drivers) {
    const k = nameKey(d);
    if (seen.has(k)) continue;
    seen.add(k);
    top.push(d);
    if (top.length >= 6) break;
  }
  $('#driverChips').innerHTML = top
    .map((d) => `<button type="button" class="chip" data-driver="${escapeHTML(d)}" aria-pressed="${nameKey(d) === cur}">${escapeHTML(d)}</button>`)
    .join('');
}

function bindForms() {
  $$('.entry [data-form]').forEach((b) => b.addEventListener('click', () => {
    if (state.editing) cancelEdit();
    switchForm(b.dataset.form);
  }));

  const tf = $('#tripForm');
  tf.cost.addEventListener('input', updateCostHint);
  tf.driver.addEventListener('input', updateDriverChips);
  $$('input[name="payment"]', tf).forEach((r) => r.addEventListener('change', updateCostHint));
  $('#driverChips').addEventListener('click', (e) => {
    const chip = e.target.closest('[data-driver]');
    if (!chip) return;
    tf.driver.value = chip.dataset.driver;
    updateDriverChips();
    tf.cost.focus();
  });

  tf.addEventListener('submit', async (e) => {
    e.preventDefault();
    const date = tf.date.value;
    const route = tf.route.value.trim();
    const payment = tf.querySelector('input[name="payment"]:checked').value;
    let cost = parseMoney(tf.cost.value);
    if (!date) return setError(tf, 'Укажите дату рейса.', tf.date);
    if (!route) return setError(tf, 'Укажите маршрут рейса.', tf.route);
    if (Number.isNaN(cost) || (cost != null && cost < 0)) return setError(tf, 'Стоимость — это число, например 45000.', tf.cost);
    if (cost == null) {
      if (payment !== 'own') return setError(tf, 'Укажите стоимость рейса. Если это свой груз — выберите «Наш груз».', tf.cost);
      cost = 0;
    }
    const rec = { date, route, driver: tf.driver.value, cost, payment, comment: tf.comment.value };
    let id;
    let verb;
    if (state.editing && state.editing.type === 'trip') {
      id = state.editing.id;
      await db.updateTrip(id, rec);
      verb = 'Рейс сохранён';
      endEdit();
    } else {
      id = await db.addTrip(rec);
      verb = 'Рейс добавлен';
    }
    resetTripForm({ date, driver: '', payment });
    await afterSave('trip', id, date, verb);
    if (!isSheetMode()) tf.route.focus();
  });

  const ef = $('#expenseForm');
  ef.addEventListener('submit', async (e) => {
    e.preventDefault();
    const date = ef.date.value;
    const amount = parseMoney(ef.amount.value);
    if (!date) return setError(ef, 'Укажите дату расхода.', ef.date);
    if (amount == null || Number.isNaN(amount) || amount <= 0) return setError(ef, 'Укажите сумму расхода числом, например 7500.', ef.amount);
    const rec = { date, category: ef.category.value, amount, comment: ef.comment.value };
    let id;
    let verb;
    if (state.editing && state.editing.type === 'expense') {
      id = state.editing.id;
      await db.updateExpense(id, rec);
      verb = 'Расход сохранён';
      endEdit();
    } else {
      id = await db.addExpense(rec);
      verb = 'Расход добавлен';
    }
    resetExpenseForm({ date, category: rec.category });
    await afterSave('expense', id, date, verb);
    if (!isSheetMode()) ef.amount.focus();
  });

  $$('[data-cancel]').forEach((b) => b.addEventListener('click', () => {
    cancelEdit();
    if (isSheetMode()) closeSheet();
  }));
}

async function afterSave(type, id, date, verb) {
  // если запись вне текущего периода — переходим к её периоду, чтобы результат был виден сразу
  if (!P.containsDate(state.period, date)) {
    state.period = state.kind === 'custom' ? P.periodFor('month', date) : P.periodFor(state.kind, date);
    if (state.kind === 'custom') state.kind = 'month';
    saveView();
  }
  state.tab = type === 'trip' ? 'trips' : 'expenses';
  state.highlight = { type, id };
  await refreshSuggestions();
  if (isSheetMode()) {
    // на телефоне закрываем панель и показываем табличку с итогом
    closeSheet();
    $('.periodbar').scrollIntoView({ behavior: prefersReducedMotion() ? 'auto' : 'smooth', block: 'start' });
  }
  await render({ flash: true });
  const s = state.data.stats;
  toast(`${verb}. За ${P.periodPhrase(state.period)}: рейсы ${money(s.tripsTotal)}, остаток ${s.balance < 0 ? '−' : ''}${money(Math.abs(s.balance))}`);
}

async function startEdit(type, id) {
  const rec = type === 'trip' ? state.data.trips.find((t) => t.id === id) : state.data.expenses.find((e) => e.id === id);
  if (!rec) return;
  state.editing = { type, id };
  switchForm(type);
  $('.entry').classList.add('entry--editing');
  $('#entryTitle').textContent = type === 'trip' ? `Изменение рейса от ${P.formatDate(rec.date)}` : `Изменение расхода от ${P.formatDate(rec.date)}`;
  if (type === 'trip') {
    const f = $('#tripForm');
    f.date.value = rec.date;
    f.route.value = rec.route;
    f.driver.value = rec.driver;
    f.cost.value = rec.cost ? String(rec.cost).replace('.', ',') : rec.payment === 'own' ? '' : '0';
    const radio = f.querySelector(`input[name="payment"][value="${rec.payment || 'bn'}"]`);
    if (radio) radio.checked = true;
    f.comment.value = rec.comment;
    f.querySelector('[type="submit"]').textContent = 'Сохранить изменения';
    f.querySelector('[data-cancel]').hidden = false;
    updateCostHint();
    updateDriverChips();
    f.route.focus();
  } else {
    const f = $('#expenseForm');
    f.date.value = rec.date;
    f.category.value = rec.category;
    f.amount.value = String(rec.amount).replace('.', ',');
    f.comment.value = rec.comment;
    f.querySelector('[type="submit"]').textContent = 'Сохранить изменения';
    f.querySelector('[data-cancel]').hidden = false;
    f.amount.focus();
  }
  renderJournal();
  if (isSheetMode()) openSheet();
  else if (window.matchMedia('(max-width: 960px)').matches) $('.entry').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function endEdit() {
  state.editing = null;
  $('.entry').classList.remove('entry--editing');
  $('#entryTitle').textContent = 'Новая запись';
  $('#tripForm [type="submit"]').textContent = 'Добавить рейс';
  $('#expenseForm [type="submit"]').textContent = 'Добавить расход';
  $$('[data-cancel]').forEach((b) => (b.hidden = true));
}
function cancelEdit() {
  endEdit();
  resetTripForm();
  resetExpenseForm();
  renderJournal();
}

// ======================================================================
// Нижняя панель с формой (телефон и телефон в альбомной ориентации)
// ======================================================================
const SHEET_MQ = window.matchMedia('(max-width: 640px), (max-width: 960px) and (max-height: 540px)');
const isSheetMode = () => SHEET_MQ.matches;
const prefersReducedMotion = () => window.matchMedia('(prefers-reduced-motion: reduce)').matches;
const sheetIsOpen = () => $('.entry').classList.contains('is-open');

function openSheet(type) {
  if (type) switchForm(type);
  $('.entry').classList.add('is-open');
  document.documentElement.classList.add('sheet-open');
  $('#sheetBackdrop').hidden = false;
  // кнопка «Назад» на Android закрывает панель, а не уходит со страницы
  if (!(history.state && history.state.sheet)) history.pushState({ sheet: true }, '');
  setTimeout(() => {
    const f = state.formType === 'trip' ? $('#tripForm') : $('#expenseForm');
    const target = state.formType === 'trip' ? (f.route.value ? f.cost : f.route) : f.amount;
    target.focus({ preventScroll: true });
  }, 60);
}

function closeSheet({ fromHistory = false } = {}) {
  if (!sheetIsOpen()) return;
  $('.entry').classList.remove('is-open');
  document.documentElement.classList.remove('sheet-open');
  $('#sheetBackdrop').hidden = true;
  if (document.activeElement && document.activeElement.closest('.entry')) document.activeElement.blur();
  if (!fromHistory && history.state && history.state.sheet) history.back();
}

function bindSheet() {
  $$('[data-open]').forEach((b) => b.addEventListener('click', () => {
    if (state.editing) cancelEdit();
    openSheet(b.dataset.open);
  }));
  const dismiss = () => {
    if (state.editing) cancelEdit();
    closeSheet();
  };
  $('#sheetClose').addEventListener('click', dismiss);
  $('#sheetBackdrop').addEventListener('click', dismiss);
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && sheetIsOpen()) dismiss();
  });
  window.addEventListener('popstate', () => closeSheet({ fromHistory: true }));
  SHEET_MQ.addEventListener('change', () => {
    if (!isSheetMode()) closeSheet();
  });
}

// ======================================================================
// Период
// ======================================================================
function bindPeriod() {
  $$('.periodbar [data-kind]').forEach((b) => b.addEventListener('click', () => setKind(b.dataset.kind)));
  $('#prevPeriod').addEventListener('click', () => shift(-1));
  $('#nextPeriod').addEventListener('click', () => shift(1));
  $('#toToday').addEventListener('click', () => {
    state.period = state.kind === 'custom'
      ? { kind: 'custom', from: P.addDays(P.today(), -(P.daysBetween(state.period.from, state.period.to))), to: P.today() }
      : P.periodFor(state.kind, P.today());
    afterPeriodChange();
  });
  const onCustom = () => {
    let from = $('#customFrom').value;
    let to = $('#customTo').value;
    if (!from || !to) return;
    if (from > to) [from, to] = [to, from];
    state.period = { kind: 'custom', from, to };
    state.customFrom = from;
    state.customTo = to;
    afterPeriodChange();
  };
  $('#customFrom').addEventListener('change', onCustom);
  $('#customTo').addEventListener('change', onCustom);

  document.addEventListener('keydown', (e) => {
    if (e.target.closest('input, select, textarea, dialog')) return;
    if (e.key === 'ArrowLeft') shift(-1);
    if (e.key === 'ArrowRight') shift(1);
  });
}

function setKind(kind) {
  const anchor = P.containsDate(state.period, P.today()) ? P.today() : state.period.from;
  state.kind = kind;
  if (kind === 'custom') {
    const from = state.customFrom || state.period.from;
    const to = state.customTo || state.period.to;
    state.period = { kind: 'custom', from, to };
  } else {
    state.period = P.periodFor(kind, anchor);
  }
  afterPeriodChange();
}

function shift(dir) {
  state.period = P.shiftPeriod(state.period, dir);
  if (state.kind === 'custom') {
    state.customFrom = state.period.from;
    state.customTo = state.period.to;
  }
  afterPeriodChange();
}

function afterPeriodChange() {
  saveView();
  if (!state.editing) {
    // дата в форме следует за просматриваемым периодом
    $('#tripForm').date.value = defaultEntryDate();
    $('#expenseForm').date.value = defaultEntryDate();
  }
  render();
}

// ======================================================================
// Журнал: вкладки, поиск, правка, удаление
// ======================================================================
function bindJournal() {
  $$('.panel [data-tab]').forEach((b) => b.addEventListener('click', () => {
    state.tab = b.dataset.tab;
    renderJournal();
  }));
  let t;
  $('#search').addEventListener('input', (e) => {
    clearTimeout(t);
    t = setTimeout(() => {
      state.search = e.target.value;
      renderJournal();
    }, 120);
  });
  $('#journal').addEventListener('click', async (e) => {
    const row = e.target.closest('tr[data-id]');
    if (!row) return;
    const type = row.dataset.type;
    const id = Number(row.dataset.id);
    if (e.target.closest('[data-edit]')) startEdit(type, id);
    if (e.target.closest('[data-del]')) removeEntry(type, id);
  });
}

async function removeEntry(type, id) {
  const rec = type === 'trip' ? state.data.trips.find((t) => t.id === id) : state.data.expenses.find((x) => x.id === id);
  if (!rec) return;
  if (type === 'trip') await db.deleteTrip(id);
  else await db.deleteExpense(id);
  if (isEditing(type, id)) cancelEdit();
  await render({ flash: true });
  toast(type === 'trip' ? `Рейс «${rec.route}» удалён` : `Расход «${rec.category}» удалён`, 'Вернуть', async () => {
    if (type === 'trip') await db.putTrip(rec);
    else await db.putExpense(rec);
    state.highlight = { type, id };
    await render({ flash: true });
    toast('Запись возвращена');
  });
}

// ======================================================================
// Импорт, отчёт, резервная копия, настройки
// ======================================================================
function bindDialogs() {
  // ----- импорт -----
  const dlg = $('#importDialog');
  let parsed = null;
  $('#btnImport').addEventListener('click', () => {
    parsed = null;
    $('#importFile').value = '';
    $('#importPreview').innerHTML = '';
    $('#fileDropText').textContent = 'Выберите файл .xlsx или перетащите его сюда';
    $('#importConfirm').disabled = true;
    $('#importConfirm').textContent = 'Загрузить';
    dlg.showModal();
  });
  const drop = $('#fileDrop');
  ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, () => drop.classList.add('is-over')));
  ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, () => drop.classList.remove('is-over')));

  $('#importFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    $('#fileDropText').textContent = file.name;
    $('#importPreview').innerHTML = '<p class="muted">Читаю файл…</p>';
    $('#importConfirm').disabled = true;
    try {
      const { parseWorkbook } = await import('./importer.js');
      const res = await parseWorkbook(file);
      const existing = await db.existingImportKeys();
      const trips = res.trips.filter((t) => !existing.has(t.importKey));
      const expenses = res.expenses.filter((x) => !existing.has(x.importKey));
      const skipped = res.trips.length + res.expenses.length - trips.length - expenses.length;
      parsed = { trips, expenses, maxDate: [...trips, ...expenses].map((r) => r.date).sort().pop() };
      renderImportPreview(res, trips, expenses, skipped);
      $('#importConfirm').disabled = !(trips.length || expenses.length);
      $('#importConfirm').textContent = trips.length || expenses.length
        ? `Загрузить ${trips.length} ${P.plural(trips.length, 'рейс', 'рейса', 'рейсов')}${expenses.length ? ` и ${expenses.length} ${P.plural(expenses.length, 'расход', 'расхода', 'расходов')}` : ''}`
        : 'Нечего загружать';
    } catch (err) {
      console.error(err);
      $('#importPreview').innerHTML = `<p class="form__error">Не удалось прочитать файл: ${escapeHTML(err.message || err)}. Нужен файл Excel в формате .xlsx.</p>`;
    }
  });

  $('#importConfirm').addEventListener('click', async () => {
    if (!parsed) return;
    const { trips, expenses, maxDate } = parsed;
    $('#importConfirm').disabled = true;
    await db.bulkImport(trips, expenses);
    dlg.close();
    if (maxDate) {
      state.kind = state.kind === 'custom' ? 'month' : state.kind;
      state.period = P.periodFor(state.kind, maxDate);
      saveView();
    }
    await refreshSuggestions();
    await render({ flash: true });
    $('#reportNote').textContent = '';
    toast(`Загружено: ${trips.length} ${P.plural(trips.length, 'рейс', 'рейса', 'рейсов')}, ${expenses.length} ${P.plural(expenses.length, 'расход', 'расхода', 'расходов')}`);
  });

  // ----- отчёт и печать -----
  $('#btnReport').addEventListener('click', async () => {
    const btn = $('#btnReport');
    const { trips, expenses, stats } = state.data;
    if (!trips.length && !expenses.length) {
      toast('За этот период нет записей — отчёт был бы пустым');
      return;
    }
    const label = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Готовлю отчёт…';
    try {
      const { buildReport } = await import('./report.js');
      await buildReport({ period: state.period, trips, expenses, stats });
      toast('Отчёт скачан');
    } catch (err) {
      console.error(err);
      toast(`Не удалось сформировать отчёт: ${err.message || err}`);
    } finally {
      btn.disabled = false;
      btn.textContent = label;
    }
  });
  $('#btnPrint').addEventListener('click', () => window.print());

  // ----- настройки -----
  const sd = $('#settingsDialog');
  $('#btnSettings').addEventListener('click', async () => {
    const c = await db.counts();
    $('#dbInfo').textContent = `В базе ${c.trips} ${P.plural(c.trips, 'рейс', 'рейса', 'рейсов')} и ${c.expenses} ${P.plural(c.expenses, 'расход', 'расхода', 'расходов')}.`;
    $('#vatRate').value = String(state.vatRate).replace('.', ',');
    sd.showModal();
  });
  $('#vatRate').addEventListener('change', async (e) => {
    const v = parseMoney(e.target.value);
    if (v == null || Number.isNaN(v) || v < 0 || v > 50) {
      e.target.value = String(state.vatRate).replace('.', ',');
      return;
    }
    state.vatRate = v;
    await db.setSetting('vatRate', v);
    updateCostHint();
    render();
  });
  $('#btnBackup').addEventListener('click', async () => {
    const data = await db.exportAll();
    const blob = new Blob([JSON.stringify(data, null, 1)], { type: 'application/json' });
    downloadBlob(blob, `Копия базы рейсов ${P.today()}.json`);
    toast('Копия базы сохранена');
  });
  $('#restoreFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      const n = (data.trips || []).length;
      if (!confirm(`Заменить текущую базу данными из копии (${n} ${P.plural(n, 'рейс', 'рейса', 'рейсов')})? Текущие записи будут удалены.`)) return;
      await db.restoreAll(data);
      state.vatRate = Number(await db.getSetting('vatRate', 22));
      sd.close();
      await refreshSuggestions();
      await render({ flash: true });
      toast('База восстановлена из копии');
    } catch (err) {
      toast(`Не удалось восстановить: ${err.message || err}`);
    } finally {
      e.target.value = '';
    }
  });
  $('#btnWipe').addEventListener('click', async () => {
    const c = await db.counts();
    if (!c.trips && !c.expenses) {
      toast('База и так пустая');
      return;
    }
    if (!confirm(`Удалить все записи (${c.trips} рейсов и ${c.expenses} расходов)? Это нельзя отменить. Сначала сохраните копию базы, если данные ещё нужны.`)) return;
    await db.wipeAll();
    sd.close();
    cancelEdit();
    await refreshSuggestions();
    await render({ flash: true });
    toast('Все записи удалены');
  });
}

function renderImportPreview(res, trips, expenses, skipped) {
  const sheets = res.sheets.map((s) => `<li>${escapeHTML(s.name)}: ${s.trips} ${P.plural(s.trips, 'рейс', 'рейса', 'рейсов')}${s.expenses ? `, ${s.expenses} ${P.plural(s.expenses, 'расход', 'расхода', 'расходов')}` : ''}</li>`).join('');
  const total = trips.reduce((s, t) => s + t.cost, 0);
  const expTotal = expenses.reduce((s, x) => s + x.amount, 0);
  const dates = [...trips, ...expenses].map((r) => r.date).sort();
  const notes = res.notes.map((n) => `<li>${escapeHTML(n.sheet)}${n.row ? `, строка ${n.row}` : ''}: ${escapeHTML(n.text)}</li>`).join('');
  $('#importPreview').innerHTML = `<div class="preview">
    ${res.sheets.length ? '' : '<p class="form__error">В файле не нашлось листов с колонками «Дата» и «Рейс».</p>'}
    ${trips.length || expenses.length ? `<p class="preview__sum">Новых записей: ${trips.length} ${P.plural(trips.length, 'рейс', 'рейса', 'рейсов')} на ${money(total)}${expenses.length ? `, ${expenses.length} ${P.plural(expenses.length, 'расход', 'расхода', 'расходов')} на ${money(expTotal)}` : ''}${dates.length ? `, даты с ${P.formatDate(dates[0])} по ${P.formatDate(dates[dates.length - 1])}` : ''}.</p>` : ''}
    ${skipped ? `<p class="muted">${skipped} ${P.plural(skipped, 'запись уже есть', 'записи уже есть', 'записей уже есть')} в базе — они не будут добавлены повторно.</p>` : ''}
    ${sheets ? `<ul>${sheets}</ul>` : ''}
    ${notes ? `<h3 class="h3">Что стоит проверить (${res.notes.length})</h3><ul class="notes">${notes}</ul>` : ''}
  </div>`;
}

// ======================================================================
// Уведомление
// ======================================================================
let toastTimer;
function toast(text, actionLabel, action) {
  const el = $('#toast');
  $('#toastText').textContent = text;
  const btn = $('#toastAction');
  if (actionLabel) {
    btn.hidden = false;
    btn.textContent = actionLabel;
    btn.onclick = () => {
      el.hidden = true;
      action();
    };
  } else {
    btn.hidden = true;
    btn.onclick = null;
  }
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => (el.hidden = true), actionLabel ? 7000 : 4500);
}

start();

// Для отладки из консоли
window.__tc = { state, db };
