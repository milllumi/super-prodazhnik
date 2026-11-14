/**
 * Бэкенд для сайта продавцов и руководителей (Google Apps Script).
 *
 * ДОСТУП выдаётся автоматически по вкладке «Доступ» (A — роль, B — ФИО):
 *   новичок, ПТТ            -> личный кабинет продавца (+ рейтинг)
 *   РТТ, УТТ (и ТМ)         -> вид руководителя салона + калькулятор бонуса
 * Кого нет в списке — не войдёт. Список можно менять в любой момент: права проверяются при каждом запросе.
 * PIN (4 цифры) сотрудник придумывает сам при первом входе. PIN хранится только в виде хеша в свойствах скрипта.
 * Сброс PIN: в колонке C напротив сотрудника поставьте любой знак (например «да») — при следующем входе он задаст новый PIN.
 * В колонке D скрипт сам отмечает «PIN задан».
 *
 * САЛОНЫ руководителей (РТТ/УТТ) берутся из таблицы «KPI ОСС текущий месяц» (ссылка на вкладке «Ссылки»),
 * лист «Сотрудники по салонам» (A — менеджер, B — магазин). Для продавцов оттуда же подставляется ТТ, если в «До сотрудника» её нет.
 *
 * КАЖДЫЙ МЕСЯЦ: если таблица новая — поменяйте SPREADSHEET_ID ниже и обновите ссылку «Kpi ОСС текущий месяц» на вкладке «Ссылки»,
 * затем «Развернуть → Управление развертываниями → Изменить → Новая версия».
 */
const CONFIG = {
  SPREADSHEET_ID: '1F9giGlGwB2laFTTx4t9rwk1ZJc_Tmkv4IstY715HpSY',
  SHEET_PLAN:    'До сотрудника ДЛ',          // рейтинг и личные результаты — из этого листа
  SHEET_SCHEDULE:'График',                    // плановые и оставшиеся смены
  KPI_BIND_SHEET:'закрепление',               // лист в KPI ОСС: привязка ТМ / УТТ / РТТ к салонам (регистр не важен)
  SHEET_SUMMARY: 'общий свод',                // страница «Общий свод» для РТТ / УТТ (ищется в этой книге, затем в KPI ОСС)
  SHEET_INCOME:  'Доход сотрудника',
  SHEET_ACCESS:  'Доступ',
  SHEET_SALONS:  'Детали продаж салоны',
  SHEET_DAYPLAN: 'План на день',
  SHEET_LINKS:   'Ссылки',
  SHEET_NEWS:    'Новости',                   // создаётся автоматически: ID, Дата, Автор, Роль, Заголовок, Текст, Статус, Модератор
  SHEET_GROUPS:  'Группа товаров',            // сюда ежечасно выгружается отчёт DataLens «Товарные группы»
  GROUPS_COLS:   11,                          // импорт пишет только A:K, столбцы L и правее не трогает
  GROUPS_NUM_COL: 11,                         // K («Сумма») — числовой формат
  GROUPS_NUM_FMT: '#,##0.00',
  LINK_KPI_KEY:  'kpi осс текущий месяц',     // название строки на вкладке «Ссылки» (регистр не важен)
  KPI_SHEET:     'Сотрудники по салонам',     // лист в KPI ОСС: A — менеджер, B — магазин
  SELLER_ROLES:  ['новичок', 'птт'],
  MANAGER_ROLES: ['ртт', 'утт', 'тм'],
  TOKEN_TTL: 6 * 3600,
  MAX_FAILS: 5,                               // неверных PIN подряд до блокировки на 10 минут
};
const METRICS = [
  { key: 'sim',  label: 'Сим' },
  { key: 'smart', label: 'Смарты' },
  { key: 'acc',  label: 'Аксессуары' },
  { key: 'serv', label: 'Услуги' },
  { key: 'ins',  label: 'Страховки' },
];
const SALON_DIRS = { 'телефоны': 'phones', 'аксессуары': 'acc', 'услуги': 'serv', 'сим': 'sim', 'страховки': 'ins' };
const DIR_LABELS = { phones: 'Телефоны', acc: 'Аксессуары', serv: 'Услуги', sim: 'Сим', ins: 'Страховки' };
const DIR_UNITS  = { phones: '₽', acc: '₽', serv: '₽', sim: 'шт', ins: '₽' };
const DIR_ORDER  = ['phones', 'acc', 'serv', 'sim', 'ins'];

/* ---------- скорость: книга открывается один раз, тяжёлые чтения — в кэше (gzip, кусками) ---------- */
const TTL_ = 2400;                                  // кэш данных: 40 мин; триггер warmCache() обновляет его каждые 30 мин
let SS_ = null, FRESH_ = false, KB_ = undefined;
function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)); }
function cacheDrop_(key) { try { CacheService.getScriptCache().remove('c:' + key); } catch (e) {} }
function cached_(key, ttl, fn) {
  const c = CacheService.getScriptCache();
  if (!FRESH_) {
    try {
      const n = Number(c.get('c:' + key) || 0);
      if (n > 0) {
        const ks = []; for (let i = 0; i < n; i++) ks.push('c:' + key + ':' + i);
        const parts = c.getAll(ks); let s = '';
        for (let i = 0; i < n && s !== null; i++) s = parts[ks[i]] == null ? null : s + parts[ks[i]];
        if (s !== null) return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(s), 'application/x-gzip')).getDataAsString());
      }
    } catch (e) {}
  }
  const val = fn();
  try {
    const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(val), 'text/plain', 'v.json')).getBytes());
    const SZ = 90000, n = Math.ceil(b64.length / SZ), o = {};
    for (let i = 0; i < n; i++) o['c:' + key + ':' + i] = b64.substr(i * SZ, SZ);
    o['c:' + key] = String(n);
    c.putAll(o, ttl);
  } catch (e) {}
  return val;
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    FRESH_ = !!req.fresh;
    if (req.action === 'start') out = start_(req);
    else if (req.action === 'setpin') out = setpin_(req);
    else if (req.action === 'login') out = login_(req);
    else if (req.action === 'data') out = data_(req);
    else if (req.action === 'links_get') out = linksGet_(req);
    else if (req.action === 'tm_sheets') out = tmSheets_(req);
    else if (req.action === 'tm_report') out = tmReport_(req);
    else if (req.action === 'cards_get') out = cardsGet_(req);
    else if (req.action === 'news_list') out = newsList_(req);
    else if (req.action === 'news_add') out = newsAdd_(req);
    else if (req.action === 'news_moderate') out = newsModerate_(req);
    else if (req.action === 'links_set') out = linksSet_(req);
    else if (req.action === 'groups_import') out = groupsImport_(req);
    else out = { error: 'bad_request' };
  } catch (err) { out = { error: 'server', message: String(err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return ContentService.createTextOutput('ok'); }

/* ---------- доступ и роли ---------- */
function readAccess_(ss) { return cached_('access', 300, function () { return readAccessRaw_(ss); }); }
function readAccessRaw_(ss) {
  const sh = ss.getSheetByName(CONFIG.SHEET_ACCESS);
  const rows = sh.getDataRange().getDisplayValues(), out = [];
  for (let i = 0; i < rows.length; i++) {
    const role = norm_(rows[i][0]), name = String(rows[i][1] || '').trim();
    if (!name) continue;
    const mgr = CONFIG.MANAGER_ROLES.indexOf(role) >= 0, sel = CONFIG.SELLER_ROLES.indexOf(role) >= 0;
    if (!mgr && !sel) continue;                       // «Доступ», «Всего (…)», «Итого» и пр. пропускаем
    out.push({ name: name, key: norm_(name), manager: mgr, roleTitle: String(rows[i][0]).trim().toUpperCase(),
      row: i + 1, reset: String(rows[i][2] || '').trim() !== '' });
  }
  return out;
}
function findUser_(list, login) {
  const t = norm_(login).split(' ').filter(String);
  if (!t.length) return { status: 'none' };
  const exact = list.filter(function (a) { return a.key === t.join(' '); });
  if (exact.length === 1) return { status: 'ok', user: exact[0] };
  const m = list.filter(function (a) { const p = a.key.split(' '); return t.every(function (x, i) { return p[i] === x; }); });
  if (m.length === 1) return { status: 'ok', user: m[0] };
  return { status: m.length ? 'ambiguous' : 'none' };
}
function salt_() {
  const P = PropertiesService.getScriptProperties();
  let s = P.getProperty('salt');
  if (!s) { s = Utilities.getUuid(); P.setProperty('salt', s); }
  return s;
}
function hash_(user, pin) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt_() + '|' + user.key + '|' + pin);
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function hasPin_(user) { return !user.reset && !!PropertiesService.getScriptProperties().getProperty('pin:' + user.key); }
function lockKey_(user) { return 'f:' + user.key; }
function isLocked_(user) { return Number(CacheService.getScriptCache().get(lockKey_(user)) || 0) >= CONFIG.MAX_FAILS; }
function fail_(user) {
  const c = CacheService.getScriptCache();
  c.put(lockKey_(user), String(Number(c.get(lockKey_(user)) || 0) + 1), 600);
}
function newToken_(user) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('t:' + token, JSON.stringify({ key: user.key }), CONFIG.TOKEN_TTL);
  CacheService.getScriptCache().remove(lockKey_(user));
  return { token: token, role: user.manager ? 'manager' : 'seller', roleTitle: user.roleTitle, canLinks: user.roleTitle === 'ТМ' };
}

// шаг 1: что показать после ввода фамилии — ввод PIN или создание PIN
function start_(req) {
  const ss = ss_();
  const f = findUser_(readAccess_(ss), req.login);
  if (f.status === 'ambiguous') return { error: 'ambiguous' };
  if (f.status === 'none') return { error: 'no_access' };
  if (isLocked_(f.user)) return { error: 'locked' };
  return { status: hasPin_(f.user) ? 'enter' : 'set' };
}
// первый вход: сотрудник сам задаёт PIN
function setpin_(req) {
  const ss = ss_();
  const f = findUser_(readAccess_(ss), req.login);
  if (f.status === 'ambiguous') return { error: 'ambiguous' };
  if (f.status !== 'ok') return { error: 'no_access' };
  const pin = String(req.pin || '');
  if (!/^\d{4}$/.test(pin)) return { error: 'bad_pin_format' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (hasPin_(f.user)) return { error: 'already_set' };
    PropertiesService.getScriptProperties().setProperty('pin:' + f.user.key, hash_(f.user, pin));
    const sh = ss.getSheetByName(CONFIG.SHEET_ACCESS);
    sh.getRange(f.user.row, 3).clearContent();                                   // снять отметку «сбросить PIN»
    cacheDrop_('access');
    sh.getRange(f.user.row, 4).setValue('PIN задан ' + Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm'));
  } finally { lock.releaseLock(); }
  return newToken_(f.user);
}
function login_(req) {
  const ss = ss_();
  const f = findUser_(readAccess_(ss), req.login);
  if (f.status === 'ambiguous') return { error: 'ambiguous' };
  if (f.status !== 'ok') return { error: 'no_access' };
  if (isLocked_(f.user)) return { error: 'locked' };
  if (!hasPin_(f.user)) return { error: 'need_pin' };
  const stored = PropertiesService.getScriptProperties().getProperty('pin:' + f.user.key);
  if (hash_(f.user, String(req.pin || '').trim()) !== stored) { fail_(f.user); return { error: 'bad_credentials' }; }
  return newToken_(f.user);
}

function data_(req) {
  const raw = CacheService.getScriptCache().get('t:' + req.token);
  if (!raw) return { error: 'auth' };
  const key = JSON.parse(raw).key;
  const ss = ss_();
  const access = readAccess_(ss);
  const user = access.filter(function (a) { return a.key === key; })[0];
  if (!user) return { error: 'auth' };                // убрали из списка «Доступ» — доступ закрыт сразу
  return cached_('u:' + key, TTL_, function () {
    const d = user.manager ? managerData_(ss, user, access) : sellerData_(ss, user, access);
    d.updated = new Date().toISOString(); return d;
  });
}
// Триггер «по времени» (каждые 30 минут): один раз запустите installTrigger() — дальше таблица читается в фоне, а сайт отвечает из кэша.
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'warmCache') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('warmCache').timeBased().everyMinutes(30).create();
  warmCache();
}
function warmCache() {
  FRESH_ = true;
  const ss = ss_();
  try { readAccess_(ss); readSalonSheet_(ss); readDayPlan_(ss); readMonthPlan_(ss); readPlan_(ss); readIncome_(ss); readSchedule_(ss); } catch (e) {}
  try { kpiShopsDrop_(); } catch (e) {}
  const access = readAccess_(ss);
  try { kpiShops_(ss, access); kpiGrades_(ss); } catch (e) {}
  Object.keys(TM_REPORTS).forEach(function (k) { try { tmReportRaw_(ss, k); } catch (e) {} });
  try { cardsAll_(ss); } catch (e) {}
  const t0 = Date.now();
  access.forEach(function (u) {                                   // готовые ответы для руководителей и продавцов (в пределах лимита времени)
    if (Date.now() - t0 > 240000) return;
    try { cached_('u:' + u.key, TTL_, function () { const d = u.manager ? managerData_(ss, u, access) : sellerData_(ss, u, access); d.updated = new Date().toISOString(); return d; }); } catch (e) {}
  });
  FRESH_ = false;
}
function kpiShopsDrop_() { CacheService.getScriptCache().remove('kpiShops2'); CacheService.getScriptCache().remove('kpiGrades'); }

/* ---------- новости: ТМ публикует сам, остальные предлагают → модерация ТМ ---------- */
const NEWS_HEAD = ['ID', 'Дата', 'Автор', 'Роль', 'Заголовок', 'Текст', 'Статус', 'Модератор'];
const ST_PUB = 'опубликовано', ST_MOD = 'на модерации', ST_REJ = 'отклонено';
function newsSheet_(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_NEWS);
  if (!sh) { sh = ss.insertSheet(CONFIG.SHEET_NEWS); sh.getRange(1, 1, 1, 8).setValues([NEWS_HEAD]); sh.setFrozenRows(1); }
  return sh;
}
function safeCell_(t, numeric) {
  if (numeric) {
    if (typeof t === 'number') return isFinite(t) ? t : '';
    const n = num_(t);
    return n === null ? '' : n;
  }
  if (typeof t === 'number') return isFinite(t) ? t : '';
  t = String(t === null || t === undefined ? '' : t);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}
function newsList_(req) {
  const ss = ss_();
  const u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  const tm = u.roleTitle === 'ТМ';
  const v = newsSheet_(ss).getDataRange().getDisplayValues(), items = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i]; if (!r[0]) continue;
    const st = norm_(r[6]), mine = norm_(r[2]) === u.key;
    if (st !== ST_PUB && !tm && !mine) continue;               // чужие предложения видит только ТМ
    items.push({ id: r[0], date: r[1], author: short_(r[2]), role: r[3], title: r[4], text: r[5], status: st, mine: mine });
  }
  items.reverse();
  return { canModerate: tm, items: items.slice(0, 120) };
}
function newsAdd_(req) {
  const ss = ss_();
  const u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  const title = String(req.title || '').trim(), text = String(req.text || '').trim();
  if (title.length < 3 || title.length > 100 || text.length < 1 || text.length > 1500) return { error: 'bad_value' };
  const tm = u.roleTitle === 'ТМ';
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = newsSheet_(ss);
    if (!tm) {
      const v = sh.getDataRange().getDisplayValues();
      const pend = v.filter(function (r, i) { return i > 0 && norm_(r[2]) === u.key && norm_(r[6]) === ST_MOD; }).length;
      if (pend >= 5) return { error: 'too_many' };
    }
    const row = sh.getLastRow() + 1;
    const id = Utilities.getUuid().slice(0, 8);
    const when = Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm');
    const rg = sh.getRange(row, 1, 1, 8); rg.setNumberFormat('@');
    rg.setValues([[id, when, u.name, u.roleTitle, safeCell_(title), safeCell_(text), tm ? ST_PUB : ST_MOD, tm ? u.name : '']]);
    return { ok: true, status: tm ? ST_PUB : ST_MOD };
  } finally { lock.releaseLock(); }
}
function newsModerate_(req) {
  const ss = ss_();
  const u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (u.roleTitle !== 'ТМ') return { error: 'forbidden' };
  const d = String(req.decision || '');
  if (['approve', 'reject', 'delete'].indexOf(d) < 0) return { error: 'bad_request' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = newsSheet_(ss), v = sh.getDataRange().getDisplayValues();
    for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(req.id)) {
      if (d === 'delete') sh.deleteRow(i + 1);
      else { sh.getRange(i + 1, 7).setValue(d === 'approve' ? ST_PUB : ST_REJ); sh.getRange(i + 1, 8).setValue(u.name); }
      return { ok: true };
    }
    return { error: 'not_found' };
  } finally { lock.releaseLock(); }
}

/* ---------- «Бонус ТМ»: листы KPI ОСС целиком (только для ТМ), разбор — на стороне страницы ---------- */
const TM_SHEETS = [['Экономические KPI', 400, 116], ['я ЛИСТ', 250, 12], ['Операторские РДО ММБ', 300, 106]];
function tmSheets_(req) {
  const ss = ss_();
  const u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (u.roleTitle !== 'ТМ') return { error: 'forbidden' };
  const book = kpiBook_(ss);
  if (!book) return { error: 'no_kpi' };
  const tables = [];
  TM_SHEETS.forEach(function (t) {
    const sh = sheetByName_(book, t[0]);
    if (!sh) return;
    const rows = Math.min(t[1], sh.getLastRow()), cols = Math.min(t[2], sh.getLastColumn());
    if (rows > 0 && cols > 0) tables.push(sh.getRange(1, 1, rows, cols).getDisplayValues());
  });
  return tables.length ? { me: u.name, tables: tables } : { error: 'no_kpi' };
}

/* ---------- ссылки: редактировать может только ТМ ---------- */
function sessionUser_(ss, token) {
  const raw = CacheService.getScriptCache().get('t:' + token);
  if (!raw) return null;
  const key = JSON.parse(raw).key;
  return readAccess_(ss).filter(function (a) { return a.key === key; })[0] || null;
}
function linksGet_(req) {
  const ss = ss_();
  const u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (u.roleTitle !== 'ТМ') return { error: 'forbidden' };
  const sh = ss.getSheetByName(CONFIG.SHEET_LINKS);
  const v = sh.getDataRange().getDisplayValues(), out = [];
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim()) out.push({ row: i + 1, name: String(v[i][0]).trim(), value: String(v[i][1] || '') });
  return { links: out };
}
function linksSet_(req) {
  const ss = ss_();
  const u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (u.roleTitle !== 'ТМ') return { error: 'forbidden' };
  const sh = ss.getSheetByName(CONFIG.SHEET_LINKS);
  const row = Number(req.row), val = String(req.value === undefined ? '' : req.value).trim();
  if (!(row >= 1 && row <= sh.getLastRow()) || !String(sh.getRange(row, 1).getDisplayValue()).trim()) return { error: 'bad_request' };
  if (val.length > 500 || /^[=+@-]/.test(val)) return { error: 'bad_value' };      // без формул
  sh.getRange(row, 2).setValue(val);
  CacheService.getScriptCache().removeAll(['kpiShops']);
  return { ok: true };
}

/* ---------- импорт выгрузки DataLens «Товарные группы» ---------- */
function groupsImportInit() {
  const P = PropertiesService.getScriptProperties();
  if (!P.getProperty('groupsImportKey')) P.setProperty('groupsImportKey', Utilities.getUuid().replace(/-/g, '').slice(0, 24));
  Logger.log(P.getProperty('groupsImportKey'));
  return P.getProperty('groupsImportKey');
}
function groupsImport_(req) {
  const P = PropertiesService.getScriptProperties();
  const key = P.getProperty('groupsImportKey');
  if (!key) return { error: 'import_not_configured' };
  if (String(req.key || '') !== key) return { error: 'forbidden' };
  const rows = req.rows;
  if (!Array.isArray(rows) || !rows.length) return { error: 'bad_request' };
  if (rows.length > 20000) return { error: 'too_many_rows' };
  const maxWidth = Math.max.apply(null, rows.map(function (r) { return Array.isArray(r) ? r.length : 0; }));
  if (!(maxWidth > 0)) return { error: 'bad_request' };
  const width = Math.min(maxWidth, CONFIG.GROUPS_COLS);
  const numCol = CONFIG.GROUPS_NUM_COL - 1;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ss_();
    let sh = ss.getSheetByName(CONFIG.SHEET_GROUPS);
    if (!sh) sh = ss.insertSheet(CONFIG.SHEET_GROUPS);
    // на листе может быть объект-таблица Google: её строку заголовка очищать нельзя
    const hdr = [];
    for (let c = 0; c < width; c++) hdr.push(safeCell_(rows[0][c], false));
    const cur = sh.getRange(1, 1, 1, width).getValues()[0].map(function (v) { return String(v); });
    if (cur.join('\u0001') !== hdr.map(function (v) { return String(v); }).join('\u0001'))
      sh.getRange(1, 1, 1, width).setValues([hdr]);
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, width).clearContent();
    const values = rows.slice(1).map(function (r) {
      const out = [];
      for (let c = 0; c < width; c++) out.push(safeCell_(r[c], c === numCol));
      return out;
    });
    if (values.length) sh.getRange(2, 1, values.length, width).setValues(values);
    if (numCol < width) sh.getRange(2, CONFIG.GROUPS_NUM_COL, Math.max(values.length, 1), 1).setNumberFormat(CONFIG.GROUPS_NUM_FMT);
    const updated = Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm');
    P.setProperty('groupsUpdatedAt', updated);
    return { ok: true, rows: values.length, cols: width, updated: updated };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- привязка салонов из KPI ОСС ---------- */
function kpiSpreadsheetId_(ss) {
  const sh = ss.getSheetByName(CONFIG.SHEET_LINKS);
  if (!sh) return null;
  const v = sh.getDataRange().getDisplayValues();
  for (let i = 0; i < v.length; i++) {
    if (norm_(v[i][0]) === CONFIG.LINK_KPI_KEY) {
      const m = /\/d\/([\w-]{20,})/.exec(String(v[i][1])) || /^([\w-]{25,})$/.exec(String(v[i][1]).trim());
      return m ? m[1] : null;
    }
  }
  return null;
}
// { 'название салона': 'грейд' } из листа «Сопоставления» (KPI ОСС): E — салон, J — «3 грейд»
function kpiGrades_(ss) {
  const cache = CacheService.getScriptCache(), hit = cache.get('kpiGrades');
  if (hit) return JSON.parse(hit);
  const out = {}, book = kpiBook_(ss), sh = book && sheetByName_(book, 'Сопоставления');
  if (sh) {
    const v = sh.getDataRange().getDisplayValues();
    let ci = 4, cg = 9;
    for (let c = 0; c < (v[0] || []).length; c++) { const h = norm_(v[0][c]); if (h === 'салон' && ci === 4) ci = c; if (h.indexOf('грейд') >= 0) cg = c; }
    for (let r = 1; r < v.length; r++) {
      const n = norm_(v[r][ci]), g = String(v[r][cg] || '').trim();
      if (n && g) out[n] = /флагман/i.test(g) ? (/\+/.test(g) ? 'Ф+' : 'Ф') : (/^\d/.test(g) ? g.charAt(0) : '');
    }
  }
  try { cache.put('kpiGrades', JSON.stringify(out), 600); } catch (e) {}
  return out;
}
function sheetByName_(book, name) {
  const t = norm_(name);
  return book.getSheets().filter(function (x) { return norm_(x.getName()) === t; })[0] || null;
}
function kpiBook_(ss) { if (KB_ === undefined) { const id = kpiSpreadsheetId_(ss); KB_ = id ? SpreadsheetApp.openById(id) : null; } return KB_; }

// { 'фио в нижнем регистре': ['Салон 1', 'Салон 2'] } — салоны руководителя.
// Источники (KPI ОСС), по порядку:
//  1) лист «Сопоставления»: E — салон, «Руководитель» — ФИО РТТ/УТТ (привязка по ФИО), «РДО / ТМ / ВрИО ТМ / Партнер» — ФАМИЛИЯ ТМ
//     (ТМ привязан к салонам только по фамилии: в таблице имя и отчество у него не указаны);
//  2) лист «закрепление» (если когда-нибудь появится);
//  3) «Сотрудники по салонам» (A — сотрудник, B — магазин) — только для тех, кого нет в источниках выше (ТТ продавцов).
function kpiShops_(ss, access) {
  const cache = CacheService.getScriptCache();
  const hit = FRESH_ ? null : cache.get('kpiShops2');
  if (hit) return JSON.parse(hit);
  const book = kpiBook_(ss), map = {}, tmBy = {}, emp = {};
  const add = function (m, n, shop) { if (!n || !shop) return; (m[n] = m[n] || []); if (m[n].indexOf(shop) < 0) m[n].push(shop); };
  if (book) {
    const soz = sheetByName_(book, 'Сопоставления');
    if (soz) {
      const v = soz.getDataRange().getDisplayValues(), H = (v[0] || []).map(norm_);
      const iSalon = H.indexOf('салон'), iMgr = H.indexOf('руководитель');
      let iTm = -1; for (let c = 0; c < H.length; c++) if (H[c].indexOf('рдо') === 0 || /(^|\s)тм(\s|$)/.test(H[c])) { iTm = c; break; }
      for (let r = 1; r < v.length && iSalon >= 0; r++) {
        const shop = String(v[r][iSalon] || '').trim(); if (!shop) continue;
        if (iTm >= 0) { const sn = norm_(v[r][iTm]).split(' ')[0]; if (sn) add(tmBy, sn, shop); }
        if (iMgr >= 0) { const nm = norm_(v[r][iMgr]); if (nm) add(map, nm, shop); }
      }
    }
    const bind = sheetByName_(book, CONFIG.KPI_BIND_SHEET);
    if (bind && access) {
      const salonNames = Object.keys(readSalonSheet_(ss).salons);
      const mgrs = (access || []).filter(function (a) { return a.manager; });
      let cur = null;
      bind.getDataRange().getDisplayValues().forEach(function (row) {
        const who = [], shops = [];
        row.forEach(function (c) {
          const t = norm_(c); if (!t) return;
          const m = mgrs.filter(function (a) { return a.key === t; })[0] || mgrs.filter(function (a) { return a.key.indexOf(t + ' ') === 0; })[0];
          if (m) { if (who.indexOf(m.key) < 0) who.push(m.key); return; }
          const sl = salonNames.filter(function (k) { return norm_(k) === t; })[0];
          if (sl) shops.push(sl);
        });
        if (who.length) cur = who;
        if (cur && shops.length) cur.forEach(function (k) { shops.forEach(function (x) { add(map, k, x); }); });
      });
    }
    const sh = sheetByName_(book, CONFIG.KPI_SHEET);
    if (sh) {
      const v = sh.getRange(1, 1, sh.getLastRow(), 2).getDisplayValues();
      for (let i = 1; i < v.length; i++) {
        const n = norm_(v[i][0]), shop = String(v[i][1] || '').trim();
        if (!n || !shop || shop === 'nan' || n.indexOf('<object') === 0) continue;
        add(emp, n, shop);
      }
    }
  }
  Object.keys(emp).forEach(function (k) { if (!map[k]) map[k] = emp[k]; });
  // ТМ — только по фамилии: все его салоны из колонки «РДО / ТМ …» (+ те, что закреплены за ним как за РТТ/УТТ)
  (access || []).forEach(function (a) {
    if (a.roleTitle !== 'ТМ') return;
    const list = tmBy[a.key.split(' ')[0]];
    if (list && list.length) { const u = list.slice(); (map[a.key] || []).forEach(function (x) { if (u.indexOf(x) < 0) u.push(x); }); map[a.key] = u; }
  });
  try { cache.put('kpiShops2', JSON.stringify(map), TTL_); } catch (e) {}
  return map;
}

/* ---------- смены из листа «График» ---------- */
// { 'фио': { planned, done, left } } — смены = отметки I, II, III, IV, V… в колонках дней (с J), день берётся из строки с 1,2,3…
function readSchedule_(ss) { return cached_('sched', TTL_, function () { return readScheduleRaw_(ss); }); }
function readScheduleRaw_(ss) {
  const sh = ss.getSheetByName(CONFIG.SHEET_SCHEDULE), out = {}, cal = { today: 0, dim: 0 };
  if (!sh) return { people: out, cal: cal };
  const v = sh.getDataRange().getDisplayValues();
  const tz = ss.getSpreadsheetTimeZone() || 'Europe/Moscow';
  let dayRow = -1;
  for (let r = 0; r < Math.min(v.length, 10); r++) if (String(v[r][9]).trim() === '1' && String(v[r][10]).trim() === '2') { dayRow = r; break; }
  if (dayRow < 0) return { people: out, cal: cal };
  const month = Number(v[dayRow][4]) || Number(String(v[dayRow][4]).trim()), year = Number(v[dayRow][7]);
  const now = new Date();
  const cm = Number(Utilities.formatDate(now, tz, 'M')), cy = Number(Utilities.formatDate(now, tz, 'yyyy')), cd = Number(Utilities.formatDate(now, tz, 'd'));
  const today = (month && year) ? ((cy * 12 + cm) === (year * 12 + month) ? cd : ((cy * 12 + cm) > (year * 12 + month) ? 99 : 0)) : cd;
  cal.today = today; cal.dim = (month && year) ? new Date(year, month, 0).getDate() : 0;
  const shift = /^(i{1,3}|iv|v|vi{0,3}|ix|x)$/i;
  for (let r = dayRow + 1; r < v.length; r++) {
    const name = String(v[r][4] || '').trim();
    if (!/^\d+$/.test(String(v[r][2]).trim()) || name.split(' ').length < 2) continue;
    const o = out[norm_(name)] = out[norm_(name)] || { planned: 0, done: 0, left: 0 };
    for (let c = 9; c < v[r].length; c++) {
      const d = c - 8; if (d > 31) break;
      if (!shift.test(String(v[r][c]).trim())) continue;
      o.planned++;
      if (d < today) o.done++; else o.left++;            // сегодняшняя смена ещё не попала в факт продаж — считаем её «осталось»
    }
  }
  return { people: out, cal: cal };
}

/* ---------- «Общий свод» для РТТ / УТТ: лист как есть, только строки их салонов ---------- */
// В «общем своде» вместо «План на день» ставим «План на месяц»: по салону — из листа «План», по строке руководителя — сумма его салонов;
// если плана на месяц нет (например «Кредиты») — план на день × дни месяца.
function summaryMonthPlan_(ss, v, salonNames) {
  let hr = -1;
  for (let r = 0; r < Math.min(v.length, 10); r++) if (v[r].some(function (c) { return norm_(c) === 'факт'; })) hr = r;
  if (hr < 1) return v;
  const mp = readMonthPlan_(ss), dim = readSalonSheet_(ss).daysInMonth || 30;
  const dirOf = function (c) {                                   // направление группы: ближайшая слева непустая ячейка строки над шапкой
    for (let i = c; i >= 0; i--) { const t = norm_(v[hr - 1][i]); if (t) return SALON_DIRS[t] || t; }
    return '';
  };
  const cols = [];
  v[hr].forEach(function (h, c) { if (/план на день/.test(norm_(h))) cols.push({ c: c, dir: dirOf(c) }); });
  if (!cols.length) return v;
  const out = v.map(function (r) { return r.slice(); });
  const sn = {}; (salonNames || []).forEach(function (n) { sn[norm_(n)] = true; });
  const sumOf = function (dir) { let t = 0, any = false; Object.keys(sn).forEach(function (k) { if (mp[k] && mp[k][dir] > 0) { t += mp[k][dir]; any = true; } }); return any ? t : null; };
  cols.forEach(function (x) { out[hr][x.c] = 'План на месяц'; });
  for (let r = hr + 1; r < out.length; r++) {
    const key = norm_(out[r][0]); if (!key) continue;
    cols.forEach(function (x) {
      const own = mp[key] && mp[key][x.dir] > 0 ? mp[key][x.dir] : null;
      const day = num_(out[r][x.c]);
      let val = own;
      if (val === null && !mp[key] && salonNames && !sn[key]) val = sumOf(x.dir);       // строка руководителя
      if (val === null && day !== null) val = Math.round(day * dim);
      if (val !== null) out[r][x.c] = String(Math.round(val));
    });
  }
  return out;
}
function summaryPage_(ss, salonNames, surname) {
  let sh = sheetByName_(ss, CONFIG.SHEET_SUMMARY);
  if (!sh) { const kb = kpiBook_(ss); if (kb) sh = sheetByName_(kb, CONFIG.SHEET_SUMMARY); }
  if (!sh) return { missing: true, name: CONFIG.SHEET_SUMMARY };
  const v0 = cached_('grid:summary', TTL_, function () { return sh.getDataRange().getDisplayValues(); });
  const v = summaryMonthPlan_(ss, v0, salonNames);
  const want = salonNames.map(norm_), sn = norm_(surname);
  const hit = function (row) {
    const t = norm_(row[0]);                                   // название салона / фамилия руководителя — в первой колонке
    return !!t && (t === sn || want.some(function (w) { return t === w || (t.length > 6 && (t.indexOf(w) >= 0 || w.indexOf(t) >= 0)); }));
  };
  // шапка — все строки до последней строки заголовков (где есть «Факт»); остальное — данные
  let hr = -1;
  for (let r = 0; r < Math.min(v.length, 10); r++) if (v[r].some(function (c) { return norm_(c) === 'факт'; })) hr = r;
  if (hr < 0) hr = Math.min(3, v.length - 1);
  const head = v.slice(0, hr + 1), body = v.slice(hr + 1).filter(hit);
  let w = 0; head.concat(body).forEach(function (r) { for (let i = r.length; i > 0; i--) if (String(r[i - 1]).trim()) { w = Math.max(w, i); break; } });
  const trim = function (rows) { return rows.slice(0, 400).map(function (r) { return r.slice(0, w); }); };
  return { name: sh.getName(), head: trim(head), rows: trim(body), found: body.length };
}

/* ---------- продавец ---------- */
function sellerData_(ss, sess, access) {
  const allowed = {};
  access.forEach(function (a) { if (!a.manager) allowed[a.key] = true; });
  // в рейтинг попадают только продавцы из списка «Доступ» (руководители и посторонние — нет)
  const sellers = readPlan_(ss).filter(function (s) { return allowed[norm_(s.name)]; });
  const income = readIncome_(ss);

  sellers.forEach(function (s) {
    if (s.rating !== null && s.rating !== undefined) { s.score = s.rating; return; }
    const pcts = s.metrics.filter(function (m) { return m.planMonth > 0; }).map(function (m) { return m.pct; });
    s.score = pcts.length ? pcts.reduce(function (a, b) { return a + b; }, 0) / pcts.length : 0;
  });
  const ranked = sellers.slice().sort(function (a, b) {
    const pa = a.place || 9999, pb = b.place || 9999;
    return pa !== pb ? pa - pb : b.score - a.score;
  });
  const meKey = sess.key;
  const rating = ranked.map(function (s, i) {
    return { place: i + 1, name: short_(s.name), score: Math.round(s.score), me: norm_(s.name) === meKey };
  });
  const me = sellers.filter(function (s) { return norm_(s.name) === meKey; })[0];
  if (!me) return { error: 'no_data' };

  const inc = income.byName[meKey] || { days: [] };
  const sched = readSchedule_(ss), mySched = sched.people[meKey] || null;
  const worked = inc.days.filter(function (d) { return d.v !== null; });
  const total = worked.reduce(function (a, d) { return a + d.v; }, 0);
  // «Отработано смен» везде одно и то же число — по графику (смены до сегодняшнего дня); если графика нет — по дням с доходом
  const shifts = (mySched && mySched.done > 0) ? mySched.done : (worked.filter(function (d) { return d.v > 0; }).length || worked.length);
  return {
    role: 'seller', name: me.name, month: income.month || '', updated: new Date().toISOString(),
    place: rating.filter(function (r) { return r.me; })[0].place, total: rating.length,
    tt: me.tt || (kpiShops_(ss, access)[meKey] || []).join(' · '), hours: me.hours, metrics: me.metrics, quality: me.quality, score: Math.round(me.score),
    schedule: mySched, cal: sched.cal,
    bonus: { total: total, shifts: shifts, avgDaily: shifts ? total / shifts : 0, days: worked },
    rating: rating,
  };
}

/* ---------- руководитель (РТТ / УТТ / ТМ) ---------- */
function managerData_(ss, sess, access) {
  const all = readSalonSheet_(ss);                         // { салон: { phones:{...}, acc:{...} ... } }
  const dayPlan = readDayPlan_(ss);                        // { салон: { phones: N, ... } }
  const mplan = readMonthPlan_(ss);                        // { салон: { sim, phones, acc, serv, ins } } — план на МЕСЯЦ (лист «План»)
  const names = kpiShops_(ss, access)[sess.key] || [];             // салоны руководителя — из KPI ОСС
  const salons = [], unmatched = [];
  names.forEach(function (n) {
    const keys = Object.keys(all.salons);
    const nn = norm_(n);
    const key = keys.filter(function (k) { return norm_(k) === nn; })[0] ||
                keys.filter(function (k) { return norm_(k).indexOf(nn) >= 0 || nn.indexOf(norm_(k)) >= 0; })[0];
    if (!key && mplan[nn]) {                                   // салон есть в плане, но продаж на листе ещё нет — показываем с нулями
      const dirs0 = DIR_ORDER.filter(function (d) { return mplan[nn][d] > 0; }).map(function (d) {
        const p = mplan[nn][d]; return { key: d, label: DIR_LABELS[d], unit: DIR_UNITS[d], plan: p, planSrc: 'месяц', fact: 0,
          planDay: all.daysInMonth ? p / all.daysInMonth : 0, lag: null, trend: null, forecast: null, days: [] };
      });
      if (dirs0.length) { if (salons.every(function (s) { return s.name !== n; })) salons.push({ name: n, dirs: dirs0 }); return; }
    }
    if (!key) { unmatched.push(n); return; }
    const dirs = DIR_ORDER.filter(function (d) { return all.salons[key][d]; }).map(function (d) {
      const x = all.salons[key][d], dp = dayPlan[key] && dayPlan[key][d];
      const mp = mplan[norm_(key)] && mplan[norm_(key)][d];
      let plan = mp > 0 ? mp : 0, src = 'месяц';
      if (!plan) { plan = x.plan; src = 'на дату'; }               // запасной вариант: «факт + отставание» из деталей продаж
      if ((!plan || plan <= 0) && dp > 0 && all.daysInMonth) { plan = dp * all.daysInMonth; src = 'день×дни'; }
      return { key: d, label: DIR_LABELS[d], unit: DIR_UNITS[d], plan: plan || 0, planSrc: src, fact: x.fact || 0,
        planDay: plan && all.daysInMonth ? plan / all.daysInMonth : (dp || 0),
        lag: x.plan > 0 ? x.plan - (x.fact || 0) : null,          // отставание от плана на текущую дату (как в таблице)
        trend: x.forecast,                                        // «Прогноз» из таблицы = выполнение плана на дату, %
        forecast: null, days: x.days };                           // прогноз на конец месяца считает страница: факт ÷ дни × дни месяца
    });
    if (salons.every(function (s) { return s.name !== key; })) salons.push({ name: key, dirs: dirs });
  });

  const me = readPlan_(ss).filter(function (s) { return norm_(s.name) === sess.key; })[0];
  const gr = kpiGrades_(ss); let grade = '';
  salons.concat(names.map(function (n) { return { name: n }; })).forEach(function (x) { if (!grade && gr[norm_(x.name)]) grade = gr[norm_(x.name)]; });
  return {
    role: 'manager', roleTitle: sess.roleTitle, name: sess.name, month: all.month, updated: new Date().toISOString(),
    daysInMonth: all.daysInMonth, daysPassed: all.daysPassed, salons: salons, kpiShops: names, unmatched: unmatched,
    summary: summaryPage_(ss, salons.map(function (x) { return x.name; }).concat(names), String(sess.name).split(' ')[0]),
    today: readSchedule_(ss).cal.today,
    canLinks: sess.roleTitle === 'ТМ',
    prefill: me ? { grade: grade, hours: me.hours.fact, att: me.quality.attach, vmr: me.quality.vmr, credits: me.quality.credits } : { grade: grade },
  };
}

function readSalonSheet_(ss) { return cached_('salons', TTL_, function () { return readSalonSheetRaw_(ss); }); }
function readSalonSheetRaw_(ss) {
  const v = ss.getSheetByName(CONFIG.SHEET_SALONS).getDataRange().getDisplayValues();
  const dre = /^(\d{2})\.(\d{2})\.(\d{4})$/;
  const salons = {};
  let month = '', daysInMonth = 0, maxDay = 0;
  const M = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  for (let r = 0; r < v.length; r++) {
    const dir = SALON_DIRS[norm_(v[r][2])];
    if (!dir || num_(v[r][1]) === null) continue;               // начало блока: «число | направление | значения по салонам»
    const gapRow = v[r];
    const fcRow = norm_(v[r + 1] && v[r + 1][2]) === 'прогноз' ? v[r + 1] : null;
    let hr = -1;
    for (let k = r + 1; k < Math.min(r + 8, v.length); k++) if (norm_(v[k][2]) === 'дата') { hr = k; break; }
    if (hr < 0) continue;
    const cols = [];
    for (let c = 3; c < v[hr].length; c++) {
      const t = String(v[hr][c]).trim();
      if (!t || t === 'Итого') break;
      cols.push({ c: c, name: t });
    }
    const dayRows = [];
    let totRow = null;
    for (let k = hr + 1; k < v.length; k++) {
      const t = String(v[k][2]).trim();
      if (dre.test(t)) dayRows.push(v[k]);
      else if (norm_(t) === 'итого') { totRow = v[k]; break; }
      else if (!t && !dayRows.length) continue;
    }
    cols.forEach(function (col) {
      const fact = totRow ? (num_(totRow[col.c]) || 0) : 0;
      const gap = num_(gapRow[col.c]);                          // «план − факт»
      const days = dayRows.map(function (dr) { return { d: String(dr[2]).trim(), v: num_(dr[col.c]) }; });
      if (!salons[col.name]) salons[col.name] = {};
      salons[col.name][dir] = { fact: fact, plan: gap === null ? 0 : gap + fact,
        forecast: fcRow ? num_(fcRow[col.c]) : null, days: days };
    });
    dayRows.forEach(function (dr) {
      const m = dre.exec(String(dr[2]).trim());
      if (m) {
        month = M[Number(m[2]) - 1] + ' ' + m[3];
        daysInMonth = new Date(Number(m[3]), Number(m[2]), 0).getDate();
        let any = false;
        cols.forEach(function (col) { if (num_(dr[col.c])) any = true; });
        if (any) maxDay = Math.max(maxDay, Number(m[1]));
      }
    });
  }
  return { salons: salons, month: month, daysInMonth: daysInMonth, daysPassed: maxDay };
}

// План на МЕСЯЦ по салонам и направлениям — лист «План» (A:Регион … E:Салон, F:SIM итого, I:Телефоны, J:Аксесс, M:Услуги, N:Страховки)
function readMonthPlan_(ss) { return cached_('mplan', TTL_, function () { return readMonthPlanRaw_(ss); }); }
function readMonthPlanRaw_(ss) {
  const sh = sheetByName_(ss, 'План'), out = {};
  if (!sh) return out;
  const v = sh.getDataRange().getDisplayValues();
  let hr = -1;
  for (let r = 0; r < Math.min(v.length, 8) && hr < 0; r++) if (v[r].some(function (c) { return norm_(c) === 'салон'; })) hr = r;
  if (hr < 0) return out;
  const H = v[hr].map(norm_);
  const at = function (f) { for (let i = 0; i < H.length; i++) if (f(H[i])) return i; return -1; };
  const ci = { name: at(function (h) { return h === 'салон'; }), sim: at(function (h) { return h.indexOf('sim-карты итого') === 0; }),
    phones: at(function (h) { return h === 'телефоны'; }), acc: at(function (h) { return h.indexOf('аксесс') === 0; }),
    serv: at(function (h) { return h === 'услуги'; }), ins: at(function (h) { return h === 'страховки'; }) };
  for (let r = hr + 1; r < v.length; r++) {
    const nm = norm_(v[r][ci.name]); if (!nm || out[nm]) continue;
    const o = {};
    ['sim', 'phones', 'acc', 'serv', 'ins'].forEach(function (k) {
      if (ci[k] < 0) return;
      const raw = String(v[r][ci[k]]); let n = num_(raw);
      if (n === null) return;
      if (/%/.test(raw)) n = n / 100;                          // «Аксесс» в таблице отформатирован как % (×100)
      o[k] = n;
    });
    out[nm] = o;
  }
  return out;
}

/* ---------- Страницы отчётов для ТМ: лист целиком (значения как в таблице), разбор и оформление — на странице ---------- */
const TM_REPORTS = { details: 'Детали продаж салоны', summary: 'общий свод', fact: 'Факт продаж', dl: 'До сотрудника ДЛ', sim: 'Детализация СИМ', recon: 'Сопоставления 1с и рег' };
function tmReport_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (u.roleTitle !== 'ТМ') return { error: 'forbidden' };
  if (!TM_REPORTS[req.name]) return { error: 'bad_request' };
  return tmReportRaw_(ss, req.name);
}
function tmReportRaw_(ss, name) {
  const title = TM_REPORTS[name];
  return cached_('rep:' + name, TTL_, function () {
    const sh = sheetByName_(ss, title);
    if (!sh) return { missing: true, name: title };
    const rows = Math.min(sh.getLastRow(), 800), cols = Math.min(sh.getLastColumn(), 130);
    if (rows < 1 || cols < 1) return { missing: true, name: title };
    let g = sh.getRange(1, 1, rows, cols).getDisplayValues();
    let lr = g.length; while (lr > 0 && !g[lr - 1].some(function (c) { return String(c).trim(); })) lr--;
    let lc = 0; g.forEach(function (r) { for (let i = r.length; i > lc; i--) if (String(r[i - 1]).trim()) { lc = i; break; } });
    g = g.slice(0, lr).map(function (r) { return r.slice(0, lc); });
    if (name === 'summary') g = summaryMonthPlan_(ss, g);          // «План на день» → «План на месяц»
    return { name: sh.getName(), grid: g, at: new Date().toISOString() };
  });
}

/* ---------- «Остальные карточки» из «Деталей продаж салоны»: всё, кроме телефонов/аксессуаров/услуг/страховок/сим ---------- */
// Находит на листе все таблицы «дата × салоны» (и сводные «Значения») правее и ниже основных блоков: защита экрана, ПЗУ, часы,
// наушники, БП, чистка, пакет, трафик и т. п. Результат фильтруется по салонам руководителя (итоги пересчитываются).
function cardsAll_(ss) {
  return cached_('cards:all', TTL_, function () {
    const sh = sheetByName_(ss, CONFIG.SHEET_SALONS);
    if (!sh) return { missing: true, cards: [] };
    const rows = Math.min(sh.getLastRow(), 700), cols = Math.min(sh.getLastColumn(), 130);
    return { cards: parseCards_(sh.getRange(1, 1, rows, cols).getDisplayValues()) };
  });
}
function cardsGet_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (!u.manager) return { error: 'forbidden' };
  const names = kpiShops_(ss, readAccess_(ss))[u.key] || [];
  const all = cardsAll_(ss);
  if (all.missing) return { missing: true, name: CONFIG.SHEET_SALONS };
  const want = {}; names.forEach(function (n) { want[norm_(n)] = true; });
  const out = [];
  all.cards.forEach(function (cd) {
    const idx = []; cd.salons.forEach(function (n, i) { if (want[norm_(n)]) idx.push(i); });
    if (!idx.length) return;
    const pick = function (a) { return idx.map(function (i) { return a[i]; }); };
    const sum = function (a) { let t = 0, any = false; a.forEach(function (v) { if (v !== null) { t += v; any = true; } }); return any ? t : null; };
    const c2 = { title: cd.title, dir: cd.dir, kind: cd.kind, pct: cd.pct, salons: pick(cd.salons), sums: (cd.sums || []).map(function (x) { return { label: x.label, vals: pick(x.vals) }; }) };
    if (cd.kind === 'pivot') c2.rows = cd.rows.map(function (r) { return { label: r.label, vals: pick(r.vals) }; });
    else {
      c2.days = cd.days; c2.plan = cd.plan; c2.v = cd.v.map(pick);
      c2.tot = cd.pct ? pick(cd.tot) : c2.v.length ? idx.map(function (x, j) { return sum(c2.v.map(function (r) { return r[j]; })); }) : [];
    }
    out.push(c2);
  });
  return { cards: out, salons: names, month: all.month || '' };
}
function parseCards_(g) {
  const cell = function (r, c) { return String((g[r] && g[r][c]) || '').trim(); };
  const isDate = function (s) { return /^\d{2}\.\d{2}\.\d{4}$/.test(String(s).trim()); };
  const bad = /^(дата|план на день|торговая точка\S*|значения|итого|закрытие)$/;
  const SL = /^(отставание|прогноз|трафик)$/;
  const MAIN = { 'телефоны': 1, 'аксессуары': 1, 'услуги': 1, 'страховки': 1, 'сим': 1 };
  const out = []; let curDir = '';
  for (let r = 0; r < g.length; r++) {
    if (SALON_DIRS[norm_(cell(r, 2))] && num_(cell(r, 1)) !== null && norm_(cell(r + 1, 2)) === 'прогноз') curDir = cell(r, 2);
    for (let c = 0; c < g[r].length; c++) {
      const h = norm_(cell(r, c));
      const viaDays = h === '' && isDate(cell(r + 1, c)) && !!cell(r, c + 1) && num_(cell(r, c + 1)) === null && !isDate(cell(r, c + 1)) && !isDate(cell(r, c - 1));
      if (h !== 'дата' && h !== 'значения' && !viaDays) continue;       // шапка с «Дата» либо пустая ячейка над датами
      const scs = [];
      for (let k = c + 1; k < g[r].length; k++) {
        const t = cell(r, k), nt = norm_(t);
        if (!t) { if (scs.length || k - c > 2) break; continue; }
        if (nt === 'итого' || nt === 'закрытие' || nt === 'дата' || nt === 'значения' || nt === 'план на день' || /^#/.test(t)) break;
        scs.push(k);
      }
      if (!scs.length) continue;
      // название и строки-подсказки («Отставание», «Прогноз», «Трафик») над шапкой
      const cands = [], sums = []; let gap = 0;
      for (let k = r - 1; k >= 0 && k >= r - 8; k--) {
        const reg = [cell(k, c - 1), cell(k, c), cell(k, c + 1)].join('');
        if (!reg && !scs.some(function (x) { return cell(k, x); })) { if (++gap > 1) break; continue; }
        gap = 0;
        const l = cell(k, c), nl = norm_(l);
        if (SL.test(nl) && scs.some(function (x) { return num_(cell(k, x)) !== null; })) { sums.unshift({ label: l, vals: scs.map(function (x) { return num_(cell(k, x)); }) }); continue; }
        if (l && num_(l) === null && !bad.test(nl) && cands.indexOf(l) < 0) cands.unshift(l);
        if (k < r - 5) break;
      }
      let title = cands.filter(function (x, i) { return !cands.some(function (y, j) { return j !== i && y.length > x.length && y.toLowerCase().indexOf(x.toLowerCase()) >= 0; }); }).join(' · ');
      if (c === 2 && MAIN[norm_(title)]) continue;                               // основные пять направлений — не берём
      if (!title) title = (curDir ? curDir + ' · ' : '') + (sums.map(function (x) { return x.label.toLowerCase(); }).join(', ') || 'таблица');
      const salons = scs.map(function (x) { return cell(r, x); });
      if (h === 'значения') {
        const rows = [];
        for (let k = r + 1; k < g.length && k < r + 12; k++) { const lb = cell(k, c); if (!lb) break; rows.push({ label: lb, vals: scs.map(function (x) { return num_(cell(k, x)); }) }); }
        if (rows.length) out.push({ title: 'Сводка: ' + rows.map(function (x) { return x.label; }).join(', '), dir: curDir, kind: 'pivot', pct: false, salons: salons, rows: rows, sums: [] });
        continue;
      }
      const dr = []; let lastD = r;
      for (let k = r + 1; k < g.length && k <= r + 45; k++) {
        const t = cell(k, c);
        if (isDate(t)) { dr.push(k); lastD = k; } else if (norm_(t) === 'итого') break; else if (dr.length && k > lastD + 3) break;
      }
      if (!dr.length) continue;
      let pct = false;
      const v = dr.map(function (k) { return scs.map(function (x) { const raw = cell(k, x); if (/%/.test(raw)) pct = true; return num_(raw); }); });
      let tot = [];
      for (let k = lastD + 1; k <= lastD + 4 && k < g.length; k++) if (norm_(cell(k, c)) === 'итого') { tot = scs.map(function (x) { return num_(cell(k, x)); }); break; }
      out.push({ title: title, dir: curDir, kind: 'days', pct: pct, salons: salons, days: dr.map(function (k) { return cell(k, c); }),
        plan: dr.map(function (k) { return num_(cell(k, c - 1)); }), v: v, tot: tot, sums: sums });
    }
  }
  return out;
}

function readDayPlan_(ss) { return cached_('dayplan', TTL_, function () { return readDayPlanRaw_(ss); }); }
function readDayPlanRaw_(ss) {
  const sh = ss.getSheetByName(CONFIG.SHEET_DAYPLAN);
  const out = {};
  if (!sh) return out;
  const v = sh.getDataRange().getDisplayValues();
  let hr = -1;
  for (let r = 0; r < Math.min(v.length, 10) && hr < 0; r++) if (norm_(v[r][1]) === 'сим') hr = r;
  if (hr < 0) return out;
  const map = { 'сим': 'sim', 'телефоны': 'phones', 'аксессуары': 'acc', 'услуги': 'serv', 'страховки': 'ins' };
  for (let r = hr + 1; r < v.length; r++) {
    const name = String(v[r][0]).trim();
    if (!name) break;
    out[name] = {};
    for (let c = 1; c < v[hr].length; c++) { const k = map[norm_(v[hr][c])]; if (k) out[name][k] = num_(v[r][c]) || 0; }
  }
  return out;
}

/* ---------- чтение листов продавцов ---------- */
function readPlan_(ss) { return cached_('plan', TTL_, function () { return readPlanRaw_(ss); }); }
function readPlanRaw_(ss) {
  const v = ss.getSheetByName(CONFIG.SHEET_PLAN).getDataRange().getDisplayValues();
  let hr = -1, nc = -1;
  for (let r = 0; r < Math.min(v.length, 15) && hr < 0; r++)
    for (let c = 0; c < v[r].length; c++) if (String(v[r][c]).trim() === 'ФИО') { hr = r; nc = c; break; }
  if (hr < 0) throw new Error('Не найден заголовок «ФИО» во вкладке ' + CONFIG.SHEET_PLAN);
  const out = [];
  for (let r = hr + 1; r < v.length; r++) {
    const row = v[r], name = String(row[nc] || '').trim();
    if (!name) break;
    const g = function (o) { return num_(row[nc + o]); };
    const metrics = METRICS.map(function (m, i) {
      const planMonth = g(7 + i) || 0, fact = g(12 + i) || 0;
      const fRaw = String(row[nc + 21 + i] || '');
      const fv = num_(fRaw);
      let forecast = null, forecastPct = null;
      if (fv !== null) {
        if (/%/.test(fRaw)) { forecastPct = fv; forecast = planMonth * fv / 100; }
        else { forecast = fv; forecastPct = planMonth > 0 ? fv / planMonth * 100 : null; }
      }
      return { key: m.key, label: m.label, planDay: Math.max(0, g(2 + i) || 0), planMonth: planMonth, fact: fact,
        pct: planMonth > 0 ? fact / planMonth * 100 : 0, forecast: forecast, forecastPct: forecastPct };
    });
    out.push({
      name: name, place: String(v[hr][nc - 1] || '').trim() === 'Место' ? num_(row[nc - 1]) : num_(row[0]),
      rating: num_(row[nc + 26]), tt: String(row[1] || '').indexOf('#') === 0 ? '' : String(row[1] || ''),
      hours: { plan: num_(row[2]) || 0, fact: num_(row[4]) || 0, fact1c: num_(row[3]) },
      metrics: metrics,
      quality: { vmr: g(17), credits: g(18), creditsRub: g(19), attach: g(20), simLeft: g(1) },
    });
  }
  return out;
}

function readIncome_(ss) { return cached_('income', TTL_, function () { return readIncomeRaw_(ss); }); }
function readIncomeRaw_(ss) {
  const v = ss.getSheetByName(CONFIG.SHEET_INCOME).getDataRange().getDisplayValues();
  const dre = /^\d{2}\.\d{2}\.\d{4}$/;
  let hr = -1, nc = -1;
  for (let r = 0; r < Math.min(v.length, 15) && hr < 0; r++)
    for (let c = 0; c < v[r].length - 1; c++)
      if (String(v[r][c]).trim() === 'Продавец' && dre.test(String(v[r][c + 1]).trim())) { hr = r; nc = c; break; }
  if (hr < 0) return { byName: {}, month: '' };
  const cols = [];
  for (let c = nc + 1; c < v[hr].length && dre.test(String(v[hr][c]).trim()); c++) cols.push(c);
  const byName = {};
  for (let r = hr + 1; r < v.length; r++) {
    const name = String(v[r][nc] || '').trim();
    if (!name || name === 'Итого') break;
    byName[norm_(name)] = { days: cols.map(function (c) { return { d: String(v[hr][c]).trim(), v: num_(v[r][c]) }; }) };
  }
  const M = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  const p = cols.length ? String(v[hr][cols[0]]).split('.') : null;
  return { byName: byName, month: p ? M[Number(p[1]) - 1] + ' ' + p[2] : '' };
}

/* ---------- утилиты ---------- */
function num_(s) {
  s = String(s === undefined || s === null ? '' : s).replace(/[\s ]/g, '').replace('%', '').replace(',', '.');
  if (s === '' || s.charAt(0) === '#') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function norm_(s) { return String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' ').trim(); }
function short_(full) { return String(full).trim().split(/\s+/).slice(0, 2).join(' '); }
