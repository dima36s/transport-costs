import * as db from './db.js';
import { PAYMENT, CATEGORIES } from './db.js';
import * as P from './periods.js';
import { compute } from './stats.js';
import { TYPES, INCOME_TYPES, TYPE, SCOPES, DIRECTION, guessDirection, amountOf, importIdentity } from './kinds.js';
import { money, escapeHTML, downloadBlob, nameKey } from './lib.js';

const $ = (s, root = document) => root.querySelector(s);
const $$ = (s, root = document) => [...root.querySelectorAll(s)];

const FORM_ID = { trip: '#tripForm', sale: '#saleForm', storage: '#storageForm', expense: '#expenseForm' };
const form = (type) => $(FORM_ID[type]);

const state = {
  kind: 'month',
  period: P.periodFor('month', P.today()),
  customFrom: null,
  customTo: null,
  scope: 'all', // 'all' | 'trip' | 'sale' | 'storage'
  tab: 'trip',
  search: '',
  formType: 'trip',
  editing: null, // { type, id }
  vatRate: 22,
  data: { trip: [], sale: [], storage: [], expense: [] },
  stats: null,
  prevStats: null,
  highlight: null, // { type, id }
  printing: false,
  directionTouched: false,
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

  fillSelects();
  TYPES.forEach((t) => resetForm(t));
  switchForm(state.scope === 'all' ? 'trip' : state.scope);
  bindForms();
  bindScope();
  bindPeriod();
  bindJournal();
  bindDialogs();
  bindSheet();
  window.addEventListener('beforeprint', () => { state.printing = true; renderJournal(); });
  window.addEventListener('afterprint', () => { state.printing = false; renderJournal(); });

  await refreshSuggestions();
  await render();

  const c = await db.counts();
  if (!Object.values(c).some(Boolean)) {
    $('#reportNote').textContent = 'База пока пустая. Добавьте первую запись в форме «Новая запись» или загрузите журнал рейсов из Excel.';
  }
}

function showFatal(msg) {
  document.querySelector('main').innerHTML = `<div class="panel" style="grid-column:1/-1"><h2 class="h2">База данных недоступна</h2><p>${escapeHTML(msg)}</p></div>`;
}

function saveView() {
  try {
    localStorage.setItem('tc-view', JSON.stringify({ kind: state.kind, customFrom: state.customFrom, customTo: state.customTo, scope: state.scope }));
  } catch (_) { /* приватный режим */ }
}
function loadView() {
  try {
    const v = JSON.parse(localStorage.getItem('tc-view') || 'null');
    if (!v) return null;
    const out = { scope: SCOPES.includes(v.scope) ? v.scope : 'all' };
    if (!['week', 'month', 'year', 'custom'].includes(v.kind)) return out;
    out.kind = v.kind;
    out.customFrom = v.customFrom;
    out.customTo = v.customTo;
    if (v.kind === 'custom' && v.customFrom && v.customTo) out.period = { kind: 'custom', from: v.customFrom, to: v.customTo };
    else if (v.kind !== 'custom') out.period = P.periodFor(v.kind, P.today());
    else out.kind = 'month';
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
  const [cur, old] = await Promise.all([db.allBetween(p.from, p.to), db.allBetween(prev.from, prev.to)]);
  state.data = cur;
  state.stats = compute(cur, p, state.vatRate, state.scope);
  state.prevStats = compute(old, prev, state.vatRate, state.scope);

  renderScope();
  renderPeriodBar();
  renderSign(flash);
  renderFacts();
  renderChart();
  renderBreakdown();
  renderPayments();
  renderCategories();
  renderJournal();
  renderActionbar();
  $('#btnReport').textContent = reportLabel();
}

const scopeIncomeLabel = () => (state.scope === 'all' ? 'Доходы' : TYPE[state.scope].many);

function reportLabel() {
  const phrase = P.periodPhrase(state.period);
  const what = { all: '', trip: ' по рейсам', sale: ' по продажам', storage: ' по хранению' }[state.scope];
  return `Скачать отчёт${what} за ${phrase} в Excel`;
}

function renderScope() {
  $$('#scopebar [data-scope]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.scope === state.scope)));
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

function signedMoney(n) {
  return (n < 0 ? '−' : '') + money(Math.abs(n));
}

function renderSign(flash) {
  const s = state.stats;
  const rows = state.scope === 'all'
    ? INCOME_TYPES.map((t) => [t, TYPE[t].many, s.incomeBy[t]])
    : [[state.scope, TYPE[state.scope].many, s.income]];
  rows.push(['expense', 'Расходы', s.expensesTotal]);
  $('#signRows').innerHTML = rows
    .map(([k, label, v]) => `<div class="sign__row"><dt>${label}</dt><dd data-k="${k}">${money(v)}</dd></div>`)
    .join('') + `<div class="sign__row sign__row--total"><dt>Остаток</dt><dd data-k="balance" class="${s.balance < 0 ? 'is-negative' : ''}">${signedMoney(s.balance)}</dd></div>`;
  const sign = $('#sign');
  sign.classList.toggle('sign--multi', rows.length > 2);
  if (flash) {
    sign.classList.remove('is-updated');
    void sign.offsetWidth;
    sign.classList.add('is-updated');
  }
}

const pct = (d) => `${d > 0 ? '+' : d < 0 ? '−' : ''}${Math.abs(d).toFixed(Math.abs(d) < 10 ? 1 : 0).replace('.', ',')}%`;
const rateText = () => String(state.vatRate).replace('.', ',');

function renderFacts() {
  const s = state.stats;
  const prev = state.prevStats;
  const sc = state.scope;
  const items = [];
  if (sc === 'all') {
    items.push(`Доходы: <b>${money(s.income)}</b>`);
    items.push(`Записей: рейсов <b>${s.countBy.trip}</b>, продаж <b>${s.countBy.sale}</b>, хранения <b>${s.countBy.storage}</b>`);
  } else if (sc === 'trip') {
    items.push(`Рейсов: <b>${s.countBy.trip}</b>${s.ownTrips ? `, из них со своим грузом <b>${s.ownTrips}</b>` : ''}`);
    if (s.paidTrips) items.push(`Средний рейс: <b>${money(s.avg.trip)}</b>`);
  } else if (sc === 'sale') {
    items.push(`Продаж: <b>${s.countBy.sale}</b>`);
    if (s.countBy.sale) items.push(`Средняя продажа: <b>${money(s.avg.sale)}</b>`);
  } else {
    const clients = s.parties.storage.length;
    items.push(`Записей: <b>${s.countBy.storage}</b>${clients ? `, клиентов <b>${clients}</b>` : ''}`);
    if (s.countBy.storage) items.push(`Средняя сумма: <b>${money(s.avg.storage)}</b>`);
  }
  if (s.vatRate > 0 && s.byPayment.bn.sum) items.push(`В т.ч. НДС ${rateText()}% с безнала: <b>${money(s.vat)}</b>`);
  if (prev.income > 0) {
    const d = ((s.income - prev.income) / prev.income) * 100;
    items.push(`${scopeIncomeLabel()} к ${P.previousLabel(state.period)}: <b class="${d >= 0 ? 'delta--up' : 'delta--down'}">${pct(d)}</b>`);
  } else if (s.income > 0) {
    items.push(`В предыдущем периоде ${sc === 'all' ? 'доходов' : TYPE[sc].none} не было`);
  }
  if (sc !== 'all' && s.generalTotal > 0) {
    items.push(`Общие расходы <b>${money(s.generalTotal)}</b> входят только в итог «Всё»`);
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

function renderChart() {
  const s = state.stats;
  const unitTitle = { day: 'По дням', week: 'По неделям', month: 'По месяцам' }[s.unit];
  const incLabel = scopeIncomeLabel();
  $('#chartTitle').textContent = unitTitle;
  $('#legend').innerHTML = `
    <li><i style="background:var(--series-income)"></i>${incLabel} <b>${money(s.income)}</b></li>
    <li><i style="background:var(--series-exp)"></i>Расходы <b>${money(s.expensesTotal)}</b></li>`;

  const host = $('#chart');
  const series = s.series;
  const max = Math.max(0, ...series.map((b) => Math.max(b.income, b.expenses)));
  if (max === 0) {
    host.innerHTML = `<p class="chart__empty">За этот период записей нет. Добавьте их в форме «Новая запись».</p>`;
    return;
  }

  const W = Math.max(300, Math.round(host.clientWidth || 700));
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
    if (b.income > 0) bars += barPath('bar-income', cx - gap / 2 - barW, y(b.income), barW, y(0) - y(b.income), r);
    if (b.expenses > 0) bars += barPath('bar-exp', cx + gap / 2, y(b.expenses), barW, y(0) - y(b.expenses), r);
    const showLabel = i % every === 0 || (s.unit === 'day' && n > 16 && i === n - 1 && (n - 1) % every > every / 2);
    if (showLabel) labels += `<text class="xlabel${b.weekend ? ' weekend' : ''}" x="${cx}" y="${H - 8}">${escapeHTML(b.label)}</text>`;
    hits += `<rect class="hit" data-i="${i}" x="${m.l + slot * i}" y="${m.t}" width="${slot}" height="${ih}"/>`;
  });

  host.innerHTML = `
    <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${unitTitle}: ${incLabel.toLowerCase()} и расходы за период">
      ${g}<g>${bars}</g><g>${labels}</g><g>${hits}</g>
    </svg>
    <table class="sr-only"><caption>${unitTitle}</caption>
      <tr><th>Период</th><th>${incLabel}</th><th>Расходы</th></tr>
      ${series.map((b) => `<tr><td>${escapeHTML(b.title)}</td><td>${money(b.income)}</td><td>${money(b.expenses)}</td></tr>`).join('')}
    </table>`;

  const tip = $('#tooltip');
  $$('.hit', host).forEach((h) => {
    const show = (ev) => {
      const b = series[Number(h.dataset.i)];
      let parts = '';
      if (state.scope === 'all') {
        const sub = INCOME_TYPES.filter((t) => b[t] > 0).map((t) => `${TYPE[t].many.toLowerCase()} ${money(b[t])}`);
        if (sub.length) parts = `<div class="tt-sub">${sub.join(', ')}</div>`;
      }
      tip.innerHTML = `<b>${escapeHTML(b.title)}</b>
        <div class="tt-row"><i style="background:var(--series-income)"></i>${incLabel}: ${money(b.income)}${state.scope !== 'all' && b.incomeCount ? ` (${b.incomeCount})` : ''}</div>${parts}
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
  resizeTimer = setTimeout(() => state.stats && renderChart(), 150);
});

// ---------- сводки ----------
function renderBreakdown() {
  const s = state.stats;
  const host = $('#breakdown');
  if (state.scope === 'all') {
    $('#breakTitle').textContent = 'По направлениям';
    const rows = s.byDirection.filter((d) => d.key !== 'general' || d.expenses);
    const tot = rows.reduce((a, d) => ({ income: a.income + d.income, expenses: a.expenses + d.expenses }), { income: 0, expenses: 0 });
    const line = (d, cls = '') => `<tr class="${cls}">
      <td>${escapeHTML(d.label)}<span class="sub">${d.key === 'general' ? 'не относятся к одному направлению' : `доходы ${money(d.income)}, расходы ${money(d.expenses)}`}</span></td>
      <td class="num c-inc">${d.key === 'general' ? '—' : money(d.income)}</td>
      <td class="num c-exp">${money(d.expenses)}</td>
      <td class="num ${d.balance < 0 ? 'neg' : ''}">${signedMoney(d.balance)}</td></tr>`;
    host.innerHTML = `<table class="mini dirs">
      <thead><tr><th>Направление</th><th class="num c-inc">Доходы</th><th class="num c-exp">Расходы</th><th class="num">Остаток</th></tr></thead>
      <tbody>${rows.map((d) => line(d, d.key === 'general' ? 'is-general' : '')).join('')}</tbody>
      <tfoot>${line({ key: 'total', label: 'Итого', income: tot.income, expenses: tot.expenses, balance: tot.income - tot.expenses })}</tfoot>
    </table>`;
    return;
  }
  const t = state.scope;
  $('#breakTitle').textContent = TYPE[t].partyTitle;
  const list = s.parties[t];
  if (!list.length) {
    host.innerHTML = `<p class="empty">За период ${TYPE[t].none} нет.</p>`;
    return;
  }
  const max = Math.max(...list.map((d) => d.sum), 1);
  host.innerHTML = `<table class="mini">
    <thead><tr><th>${TYPE[t].partyCol}</th><th class="num">${TYPE[t].countCol}</th><th class="share"><span class="sr-only">Доля</span></th><th class="num">Сумма</th></tr></thead>
    <tbody>${list.map((d) => `<tr>
      <td>${escapeHTML(d.name)}${d.variants.length > 1 ? `<span class="variants">также: ${escapeHTML(d.variants.filter((v) => v !== d.name).join(', '))}</span>` : ''}</td>
      <td class="num">${d.count}${d.own ? ` <span class="variants" style="display:inline">(${d.own} свой)</span>` : ''}</td>
      <td class="share"><div class="sharebar"><span style="width:${(d.sum / max) * 100}%"></span></div></td>
      <td class="num">${money(d.sum)}</td></tr>`).join('')}</tbody>
  </table>`;
}

function renderPayments() {
  const s = state.stats;
  const rows = Object.entries(s.byPayment).filter(([, v]) => v.count);
  const host = $('#payments');
  if (!rows.length) {
    host.innerHTML = `<p class="empty">За период ${state.scope === 'all' ? 'поступлений' : TYPE[state.scope].none} нет.</p>`;
    return;
  }
  const forms = state.scope === 'all' ? ['запись', 'записи', 'записей'] : TYPE[state.scope].plural;
  host.innerHTML = `<table class="mini"><tbody>${rows
    .map(([k, v]) => `<tr><td>${PAYMENT[k]}</td><td class="num">${v.count} ${P.plural(v.count, ...forms)}</td><td class="num">${money(v.sum)}</td></tr>`)
    .join('')}</tbody></table>`;
}

function renderCategories() {
  const s = state.stats;
  const host = $('#categories');
  $('#catTitle').textContent = state.scope === 'all' ? 'Расходы по статьям' : `Расходы направления «${TYPE[state.scope].many}»`;
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
const visibleTabs = () => (state.scope === 'all' ? TYPES : [state.scope, 'expense']);

function journalList(type) {
  const list = state.data[type];
  if (type === 'expense' && state.scope !== 'all') return list.filter((e) => e.direction === state.scope);
  return list;
}

function matches(r, q) {
  if (!q) return true;
  const hay = nameKey([r.route, r.driver, r.buyer, r.client, r.cargo, r.comment, r.category, DIRECTION[r.direction], PAYMENT[r.payment]].filter(Boolean).join(' '));
  return q.split(/\s+/).every((w) => hay.includes(w));
}

const payTag = (p) => (p ? `<span class="tag tag--${p}">${PAYMENT[p]}</span>` : '<span class="tag">не указано</span>');
const shortDate = (iso) => P.formatDate(iso).slice(0, 5);

function termText(r) {
  if (r.termFrom && r.termTo) {
    const n = P.daysBetween(r.termFrom, r.termTo) + 1;
    const a = r.termFrom.slice(0, 4) === r.termTo.slice(0, 4) ? shortDate(r.termFrom) : P.formatDate(r.termFrom);
    return `${a} – ${P.formatDate(r.termTo)}, ${n} ${P.plural(n, 'день', 'дня', 'дней')}`;
  }
  if (r.termFrom) return `с ${P.formatDate(r.termFrom)}`;
  if (r.termTo) return `по ${P.formatDate(r.termTo)}`;
  return '';
}

// Колонки журнала. Классы ячеек задают раскладку карточек на узком экране:
// route — заголовок, driver — вторая строка, num — сумма, pay — оплата, term — срок.
const COLS = {
  trip: [
    { th: 'Рейс', cls: 'route', v: (r) => escapeHTML(r.route) },
    { th: 'Водитель', cls: 'driver', v: (r) => escapeHTML(r.driver) },
    { th: 'Стоимость', cls: 'num', v: (r) => money(r.cost) },
    { th: 'Оплата', cls: 'pay', v: (r) => payTag(r.payment) },
    { th: 'Комментарий', cls: 'comment', v: (r) => escapeHTML(r.comment) },
  ],
  sale: [
    { th: 'Покупатель / за что', cls: 'route', v: (r) => escapeHTML(r.buyer) },
    { th: 'Сумма', cls: 'num', v: (r) => money(r.amount) },
    { th: 'Оплата', cls: 'pay', v: (r) => payTag(r.payment) },
    { th: 'Комментарий', cls: 'comment', v: (r) => escapeHTML(r.comment) },
  ],
  storage: [
    { th: 'Клиент', cls: 'route', v: (r) => escapeHTML(r.client) },
    { th: 'Груз', cls: 'driver', v: (r) => escapeHTML(r.cargo) },
    { th: 'Срок', cls: 'term', v: (r) => termText(r) },
    { th: 'Сумма', cls: 'num', v: (r) => money(r.amount) },
    { th: 'Оплата', cls: 'pay', v: (r) => payTag(r.payment) },
    { th: 'Комментарий', cls: 'comment', v: (r) => escapeHTML(r.comment) },
  ],
  expense: [
    { th: 'Статья', cls: 'route', v: (r) => escapeHTML(r.category) },
    { th: 'Направление', cls: 'driver', v: (r) => DIRECTION[r.direction] || '' },
    { th: 'Сумма', cls: 'num', v: (r) => money(r.amount) },
    { th: 'За что', cls: 'comment', v: (r) => escapeHTML(r.comment) },
  ],
};

function renderJournal() {
  const tabs = visibleTabs();
  if (!tabs.includes(state.tab)) state.tab = tabs[0];
  $$('#journalTabs [data-tab]').forEach((b) => {
    const t = b.dataset.tab;
    b.hidden = !tabs.includes(t);
    b.setAttribute('aria-selected', String(t === state.tab));
    b.querySelector('.seg__count').textContent = journalList(t).length;
  });
  const q = nameKey(state.search);
  const host = $('#journal');
  host.innerHTML = tabs
    .filter((t) => state.printing || t === state.tab)
    .map((t) => journalTable(t, journalList(t).filter((r) => matches(r, q)), q))
    .join('');

  if (state.highlight) {
    const row = host.querySelector(`tr[data-type="${state.highlight.type}"][data-id="${state.highlight.id}"]`);
    if (row) row.classList.add('is-new');
    state.highlight = null;
  }
}

const byDateDesc = (a, b) => b.date.localeCompare(a.date) || (b.createdAt || 0) - (a.createdAt || 0) || b.id - a.id;

function journalTable(type, list, q) {
  const cols = COLS[type];
  if (!list.length) {
    return `<p class="journal__empty">${q ? 'Ничего не найдено.' : `За этот период ${TYPE[type].none} нет.`}</p>`;
  }
  const sorted = [...list].sort(byDateDesc);
  const total = sorted.reduce((s, r) => s + amountOf(type, r), 0);
  const label = `Итого${q ? ' по найденным' : ''}: ${sorted.length} ${P.plural(sorted.length, ...TYPE[type].plural)}`;
  return `<table class="journal">
    ${state.printing ? `<caption class="h2" style="text-align:left;padding:6px 0">${TYPE[type].many}</caption>` : ''}
    <thead><tr><th>Дата</th>${cols.map((c) => `<th class="${c.cls === 'num' ? 'num' : ''}">${c.th}</th>`).join('')}<th class="acts"><span class="sr-only">Действия</span></th></tr></thead>
    <tbody>${sorted.map((r) => `<tr data-type="${type}" data-id="${r.id}" class="${isEditing(type, r.id) ? 'is-editing' : ''}">
      <td class="date">${P.formatDate(r.date)}</td>
      ${cols.map((c) => `<td class="${c.cls}">${c.v(r)}</td>`).join('')}
      <td class="acts"><button type="button" class="linkbtn" data-edit>Изменить</button><button type="button" class="linkbtn linkbtn--danger" data-del>Удалить</button></td>
    </tr>`).join('')}</tbody>
    <tfoot><tr><td></td>${cols.map((c) => `<td class="${c.cls === 'num' ? 'num' : ''}">${c.cls === 'route' ? label : c.cls === 'num' ? money(total) : ''}</td>`).join('')}<td></td></tr></tfoot>
  </table>`;
}

const isEditing = (type, id) => state.editing && state.editing.type === type && state.editing.id === id;

// ======================================================================
// Формы
// ======================================================================
function fillSelects() {
  $('#eCategory').innerHTML = CATEGORIES.map((c) => `<option>${c}</option>`).join('');
  $('#eDirection').innerHTML = Object.entries(DIRECTION)
    .map(([k, v]) => `<option value="${k}">${k === 'general' ? 'Общие — на всё сразу' : v}</option>`)
    .join('');
}

function parseMoney(s) {
  const t = String(s ?? '').replace(/\s/g, '').replace(/₽|р\.?|руб\.?/gi, '').replace(',', '.');
  if (t === '') return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : NaN;
}
const badAmount = (n) => n == null || Number.isNaN(n) || n <= 0;
const numText = (n) => String(n).replace('.', ',');

function defaultEntryDate() {
  // если смотрим прошлый период — по умолчанию ставим дату внутри него
  const t = P.today();
  return P.containsDate(state.period, t) ? t : state.period.to < t ? state.period.to : state.period.from;
}

function defaultDirection(category) {
  return state.scope !== 'all' ? state.scope : guessDirection(category);
}

function setPayment(f, value) {
  const r = f.querySelector(`input[name="payment"][value="${value}"]`);
  if (r) r.checked = true;
}

function resetForm(type, keep = {}) {
  const f = form(type);
  f.reset();
  f.date.value = keep.date || defaultEntryDate();
  if (keep.payment) setPayment(f, keep.payment);
  if (type === 'expense') {
    if (keep.category) f.category.value = keep.category;
    state.directionTouched = Boolean(keep.direction);
    f.direction.value = keep.direction || defaultDirection(f.category.value);
  }
  setError(f, '');
  updateHints(type);
  if (type === 'trip') updateDriverChips();
}

function setError(f, msg, field) {
  const el = f.querySelector('.form__error');
  el.textContent = msg;
  el.hidden = !msg;
  f.querySelectorAll('[aria-invalid]').forEach((i) => i.removeAttribute('aria-invalid'));
  if (field) {
    field.setAttribute('aria-invalid', 'true');
    field.focus();
  }
}

function switchForm(type) {
  state.formType = type;
  $$('.entry [data-form]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.form === type)));
  TYPES.forEach((t) => (form(t).hidden = t !== type));
}

function updateHints(type) {
  if (type === 'expense') return;
  const f = form(type);
  const field = type === 'trip' ? f.cost : f.amount;
  const amount = parseMoney(field.value);
  const pay = f.querySelector('input[name="payment"]:checked')?.value;
  let text = '';
  if (type === 'trip' && pay === 'own' && !field.value) {
    text = 'Свой груз: стоимость можно оставить пустой — запишется 0';
  } else if (pay === 'bn' && amount > 0 && state.vatRate > 0) {
    text = `в т.ч. НДС ${rateText()}%: ${money((amount * state.vatRate) / (100 + state.vatRate))}`;
  }
  f.querySelector('[data-hint="amount"]').textContent = text;
  if (type === 'storage') {
    let a = f.termFrom.value;
    let b = f.termTo.value;
    let t = '';
    if (a && b) {
      if (a > b) [a, b] = [b, a];
      const n = P.daysBetween(a, b) + 1;
      t = `${n} ${P.plural(n, 'день', 'дня', 'дней')}`;
    }
    f.querySelector('[data-hint="term"]').textContent = t;
  }
}

/** Проверка и сбор записи из формы: { rec } или { err: [текст, поле] } */
function readForm(type) {
  const f = form(type);
  const date = f.date.value;
  if (!date) return { err: ['Укажите дату.', f.date] };
  const payment = f.querySelector('input[name="payment"]:checked')?.value;
  const comment = f.comment.value;

  if (type === 'trip') {
    const route = f.route.value.trim();
    let cost = parseMoney(f.cost.value);
    if (!route) return { err: ['Укажите маршрут рейса.', f.route] };
    if (Number.isNaN(cost) || (cost != null && cost < 0)) return { err: ['Стоимость — это число, например 45000.', f.cost] };
    if (cost == null) {
      if (payment !== 'own') return { err: ['Укажите стоимость рейса. Если это свой груз — выберите «Наш груз».', f.cost] };
      cost = 0;
    }
    return { rec: { date, route, driver: f.driver.value, cost, payment, comment } };
  }
  if (type === 'sale') {
    const buyer = f.buyer.value.trim();
    const amount = parseMoney(f.amount.value);
    if (!buyer) return { err: ['Укажите покупателя или за что продажа.', f.buyer] };
    if (badAmount(amount)) return { err: ['Укажите сумму продажи числом, например 45000.', f.amount] };
    return { rec: { date, buyer, amount, payment, comment } };
  }
  if (type === 'storage') {
    const client = f.client.value.trim();
    const amount = parseMoney(f.amount.value);
    if (!client) return { err: ['Укажите клиента.', f.client] };
    if (badAmount(amount)) return { err: ['Укажите сумму за хранение числом, например 15000.', f.amount] };
    return { rec: { date, client, cargo: f.cargo.value, termFrom: f.termFrom.value, termTo: f.termTo.value, amount, payment, comment } };
  }
  const amount = parseMoney(f.amount.value);
  if (badAmount(amount)) return { err: ['Укажите сумму расхода числом, например 7500.', f.amount] };
  return { rec: { date, category: f.category.value, direction: f.direction.value, amount, comment } };
}

function fillForm(type, r) {
  const f = form(type);
  f.date.value = r.date;
  f.comment.value = r.comment || '';
  if (type === 'trip') {
    f.route.value = r.route;
    f.driver.value = r.driver;
    f.cost.value = r.cost ? numText(r.cost) : r.payment === 'own' ? '' : '0';
    setPayment(f, r.payment || 'bn');
    updateDriverChips();
  } else if (type === 'sale') {
    f.buyer.value = r.buyer;
    f.amount.value = numText(r.amount);
    setPayment(f, r.payment || 'bn');
  } else if (type === 'storage') {
    f.client.value = r.client;
    f.cargo.value = r.cargo || '';
    f.termFrom.value = r.termFrom || '';
    f.termTo.value = r.termTo || '';
    f.amount.value = numText(r.amount);
    setPayment(f, r.payment || 'bn');
  } else {
    f.category.value = r.category;
    f.direction.value = r.direction || guessDirection(r.category);
    state.directionTouched = true;
    f.amount.value = numText(r.amount);
  }
  updateHints(type);
}

/** Что оставить в форме после сохранения — удобно вносить подряд */
function keepAfterSave(type, rec) {
  if (type === 'expense') return { date: rec.date, category: rec.category, direction: rec.direction };
  return { date: rec.date, payment: rec.payment };
}

function focusFirst(type, opts) {
  const f = form(type);
  let el = f.amount;
  if (type === 'trip') el = f.route.value ? f.cost : f.route;
  else if (type === 'sale' && !f.buyer.value) el = f.buyer;
  else if (type === 'storage' && !f.client.value) el = f.client;
  el.focus(opts);
}

let suggest = { drivers: [], routes: [], buyers: [], clients: [], cargos: [] };
async function refreshSuggestions() {
  suggest = await db.suggestions();
  const opts = (list) => list.map((d) => `<option value="${escapeHTML(d)}">`).join('');
  $('#driversList').innerHTML = opts(suggest.drivers);
  $('#routesList').innerHTML = opts(suggest.routes);
  $('#buyersList').innerHTML = opts(suggest.buyers);
  $('#clientsList').innerHTML = opts(suggest.clients);
  $('#cargosList').innerHTML = opts(suggest.cargos);
  updateDriverChips();
}
function updateDriverChips() {
  const cur = nameKey(form('trip').driver.value);
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

  // подсказки: НДС, свой груз, длительность хранения
  for (const t of ['trip', 'sale', 'storage']) {
    const f = form(t);
    f.addEventListener('input', () => updateHints(t));
    f.addEventListener('change', () => updateHints(t));
  }
  const tf = form('trip');
  tf.driver.addEventListener('input', updateDriverChips);
  $('#driverChips').addEventListener('click', (e) => {
    const chip = e.target.closest('[data-driver]');
    if (!chip) return;
    tf.driver.value = chip.dataset.driver;
    updateDriverChips();
    tf.cost.focus();
  });

  const ef = form('expense');
  ef.category.addEventListener('change', () => {
    if (!state.directionTouched) ef.direction.value = defaultDirection(ef.category.value);
  });
  ef.direction.addEventListener('change', () => (state.directionTouched = true));

  for (const type of TYPES) {
    const f = form(type);
    f.addEventListener('submit', async (e) => {
      e.preventDefault();
      const { rec, err } = readForm(type);
      if (err) return setError(f, err[0], err[1]);
      let id;
      let verb;
      if (state.editing && state.editing.type === type) {
        id = state.editing.id;
        await db.update(type, id, rec);
        verb = TYPE[type].saved;
        endEdit();
      } else {
        id = await db.add(type, rec);
        verb = TYPE[type].added;
      }
      resetForm(type, keepAfterSave(type, rec));
      await afterSave(type, id, rec, verb);
      if (!isSheetMode()) focusFirst(type);
    });
  }

  $$('[data-cancel]').forEach((b) => b.addEventListener('click', () => {
    cancelEdit();
    if (isSheetMode()) closeSheet();
  }));
}

function visibleInScope(type, rec) {
  if (state.scope === 'all' || type === state.scope) return true;
  return type === 'expense' && rec.direction === state.scope;
}

async function afterSave(type, id, rec, verb) {
  // если запись вне текущего периода или направления — переходим туда, где её видно
  if (!P.containsDate(state.period, rec.date)) {
    state.period = state.kind === 'custom' ? P.periodFor('month', rec.date) : P.periodFor(state.kind, rec.date);
    if (state.kind === 'custom') state.kind = 'month';
  }
  if (!visibleInScope(type, rec)) state.scope = 'all';
  saveView();
  state.tab = type;
  state.highlight = { type, id };
  await refreshSuggestions();
  if (isSheetMode()) {
    // на телефоне закрываем панель и показываем табличку с итогом
    closeSheet();
    $('#scopebar').scrollIntoView({ behavior: prefersReducedMotion() ? 'auto' : 'smooth', block: 'start' });
  }
  await render({ flash: true });
  const s = state.stats;
  toast(`${verb}. За ${P.periodPhrase(state.period)}: ${scopeIncomeLabel().toLowerCase()} ${money(s.income)}, остаток ${signedMoney(s.balance)}`);
}

function startEdit(type, id) {
  const rec = state.data[type].find((r) => r.id === id);
  if (!rec) return;
  state.editing = { type, id };
  switchForm(type);
  $('.entry').classList.add('entry--editing');
  $('#entryTitle').textContent = `Изменение ${TYPE[type].gen} от ${P.formatDate(rec.date)}`;
  fillForm(type, rec);
  const f = form(type);
  f.querySelector('[type="submit"]').textContent = 'Сохранить изменения';
  f.querySelector('[data-cancel]').hidden = false;
  renderJournal();
  if (isSheetMode()) openSheet();
  else {
    if (window.matchMedia('(max-width: 960px)').matches) $('.entry').scrollIntoView({ behavior: 'smooth', block: 'start' });
    focusFirst(type, { preventScroll: true });
  }
}

function endEdit() {
  state.editing = null;
  $('.entry').classList.remove('entry--editing');
  $('#entryTitle').textContent = 'Новая запись';
  TYPES.forEach((t) => (form(t).querySelector('[type="submit"]').textContent = TYPE[t].add));
  $$('[data-cancel]').forEach((b) => (b.hidden = true));
}
function cancelEdit() {
  endEdit();
  TYPES.forEach((t) => resetForm(t));
  renderJournal();
}

// ======================================================================
// Направления
// ======================================================================
function bindScope() {
  $$('#scopebar [data-scope]').forEach((b) => b.addEventListener('click', () => setScope(b.dataset.scope)));
}

function setScope(scope) {
  if (state.scope === scope) return;
  state.scope = scope;
  saveView();
  if (!state.editing) {
    if (scope !== 'all') switchForm(scope);
    const ef = form('expense');
    if (!state.directionTouched) ef.direction.value = defaultDirection(ef.category.value);
  }
  if (scope !== 'all') state.tab = scope;
  render();
}

function renderActionbar() {
  const main = state.scope === 'all' ? 'trip' : state.scope;
  $$('#actionbar [data-open]').forEach((b) => b.classList.toggle('is-primary', b.dataset.open === main));
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
  setTimeout(() => focusFirst(state.formType, { preventScroll: true }), 60);
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
    state.period = { kind: 'custom', from: state.customFrom || state.period.from, to: state.customTo || state.period.to };
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
    TYPES.forEach((t) => (form(t).date.value = defaultEntryDate()));
  }
  render();
}

// ======================================================================
// Журнал: вкладки, поиск, правка, удаление
// ======================================================================
function bindJournal() {
  $$('#journalTabs [data-tab]').forEach((b) => b.addEventListener('click', () => {
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
  $('#journal').addEventListener('click', (e) => {
    const row = e.target.closest('tr[data-id]');
    if (!row) return;
    const type = row.dataset.type;
    const id = Number(row.dataset.id);
    if (e.target.closest('[data-edit]')) startEdit(type, id);
    if (e.target.closest('[data-del]')) removeEntry(type, id);
  });
}

function deletedText(type, r) {
  return {
    trip: `Рейс «${r.route}» удалён`,
    sale: `Продажа «${r.buyer}» удалена`,
    storage: `Хранение «${r.client}» удалено`,
    expense: `Расход «${r.category}» удалён`,
  }[type];
}

async function removeEntry(type, id) {
  const rec = state.data[type].find((r) => r.id === id);
  if (!rec) return;
  await db.remove(type, id);
  if (isEditing(type, id)) cancelEdit();
  await render({ flash: true });
  toast(deletedText(type, rec), 'Вернуть', async () => {
    await db.put(type, rec);
    state.highlight = { type, id };
    await render({ flash: true });
    toast('Запись возвращена');
  });
}

// ======================================================================
// Импорт, отчёт, резервная копия, настройки
// ======================================================================
function countsText(c) {
  const part = (n, t) => `${n} ${P.plural(n, ...TYPE[t].plural)}`;
  return `${part(c.trip, 'trip')}, ${part(c.sale, 'sale')}, ${c.storage} ${P.plural(c.storage, 'запись', 'записи', 'записей')} о хранении и ${part(c.expense, 'expense')}`;
}

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
      const trips = res.trips.filter((t) => !existing.has(importIdentity(t.importKey)));
      const expenses = res.expenses.filter((x) => !existing.has(importIdentity(x.importKey)));
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
    const types = state.scope === 'all' ? TYPES : [state.scope, 'expense'];
    if (!types.some((t) => journalList(t).length)) {
      toast('За этот период нет записей — отчёт был бы пустым');
      return;
    }
    const label = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Готовлю отчёт…';
    try {
      const { buildReport } = await import('./report.js');
      await buildReport({ period: state.period, scope: state.scope, data: state.data, stats: state.stats });
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
    $('#dbInfo').textContent = `В базе ${countsText(await db.counts())}.`;
    $('#vatRate').value = rateText();
    sd.showModal();
  });
  $('#vatRate').addEventListener('change', async (e) => {
    const v = parseMoney(e.target.value);
    if (v == null || Number.isNaN(v) || v < 0 || v > 50) {
      e.target.value = rateText();
      return;
    }
    state.vatRate = v;
    await db.setSetting('vatRate', v);
    ['trip', 'sale', 'storage'].forEach(updateHints);
    render();
  });
  $('#btnBackup').addEventListener('click', async () => {
    const data = await db.exportAll();
    const blob = new Blob([JSON.stringify(data, null, 1)], { type: 'application/json' });
    downloadBlob(blob, `Копия базы учёта ${P.today()}.json`);
    toast('Копия базы сохранена');
  });
  $('#restoreFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      const c = { trip: (data.trips || []).length, sale: (data.sales || []).length, storage: (data.storage || []).length, expense: (data.expenses || []).length };
      if (!confirm(`Заменить текущую базу данными из копии (${countsText(c)})? Текущие записи будут удалены.`)) return;
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
    if (!Object.values(c).some(Boolean)) {
      toast('База и так пустая');
      return;
    }
    if (!confirm(`Удалить все записи (${countsText(c)})? Это нельзя отменить. Сначала сохраните копию базы, если данные ещё нужны.`)) return;
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
