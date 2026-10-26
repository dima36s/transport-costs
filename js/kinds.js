// Типы записей и направления учёта — общий справочник для форм, итогов и отчёта.

export const TYPES = ['trip', 'sale', 'storage', 'expense'];
export const INCOME_TYPES = ['trip', 'sale', 'storage'];

export const TYPE = {
  trip: {
    many: 'Рейсы',
    gen: 'рейса',
    plural: ['рейс', 'рейса', 'рейсов'],
    add: 'Добавить рейс',
    added: 'Рейс добавлен',
    saved: 'Рейс сохранён',
    partyKey: 'driver',
    partyTitle: 'По водителям',
    partyCol: 'Водитель',
    partyEmpty: 'Без водителя',
    countCol: 'Рейсов',
    none: 'рейсов',
  },
  sale: {
    many: 'Продажи',
    gen: 'продажи',
    plural: ['продажа', 'продажи', 'продаж'],
    add: 'Добавить продажу',
    added: 'Продажа добавлена',
    saved: 'Продажа сохранена',
    partyKey: 'buyer',
    partyTitle: 'По покупателям',
    partyCol: 'Покупатель',
    partyEmpty: 'Без покупателя',
    countCol: 'Продаж',
    none: 'продаж',
  },
  storage: {
    many: 'Хранение',
    gen: 'записи о хранении',
    plural: ['запись', 'записи', 'записей'],
    add: 'Добавить хранение',
    added: 'Хранение добавлено',
    saved: 'Хранение сохранено',
    partyKey: 'client',
    partyTitle: 'По клиентам',
    partyCol: 'Клиент',
    partyEmpty: 'Без клиента',
    countCol: 'Записей',
    none: 'записей о хранении',
  },
  expense: {
    many: 'Расходы',
    gen: 'расхода',
    plural: ['расход', 'расхода', 'расходов'],
    add: 'Добавить расход',
    added: 'Расход добавлен',
    saved: 'Расход сохранён',
    countCol: 'Записей',
    none: 'расходов',
  },
};

export const SCOPES = ['all', 'trip', 'sale', 'storage'];

/** К какому направлению относится расход. «Общие» входят только в итог «Всё». */
export const DIRECTION = {
  trip: 'Рейсы',
  sale: 'Продажи',
  storage: 'Хранение',
  general: 'Общие',
};

const TRIP_CATEGORIES = ['Топливо', 'Платные дороги', 'ТО и ремонт', 'Штрафы'];

/** Направление по статье — для старых расходов и подсказки в форме */
export function guessDirection(category) {
  if (TRIP_CATEGORIES.includes(category)) return 'trip';
  if (category === 'Закупка товара') return 'sale';
  return 'general';
}

/**
 * Ключ для поиска уже загруженных строк Excel. У расходов статья в ключе не участвует:
 * распознавание статей могло поменяться, а повторная загрузка не должна давать дублей.
 */
export function importIdentity(key) {
  if (!key) return key;
  const parts = key.split('|');
  return parts[0] === 'e' && parts.length === 6 ? [parts[0], parts[1], parts[3], parts[4], parts[5]].join('|') : key;
}

export function amountOf(type, r) {
  return Number(type === 'trip' ? r.cost : r.amount) || 0;
}
