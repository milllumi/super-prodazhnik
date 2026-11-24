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
  SHEET_SCHEDULE:'График',                    // плановые и оставшиеся смены
  KPI_BIND_SHEET:'закрепление',               // лист в KPI ОСС: привязка ТМ / УТТ / РТТ к салонам (регистр не важен)
  SHEET_SUMMARY: 'общий свод',                // страница «Общий свод» для РТТ / УТТ (ищется в этой книге, затем в KPI ОСС)
  SHEET_INCOME:  'Доход сотрудника',
  SHEET_ACCESS:  'Доступ',
  SHEET_SALONS:  'Детали продаж салоны',
  SHEET_LINKS:   'Ссылки',
  SHEET_NEWS:    'Новости',                   // создаётся автоматически: ID, Дата, Автор, Роль, Заголовок, Текст, Статус, Модератор
  LINK_KPI_KEY:  'kpi осс текущий месяц',     // название строки на вкладке «Ссылки» (регистр не важен)
  KPI_SHEET:     'Сотрудники по салонам',     // лист в KPI ОСС: A — менеджер, B — магазин
  SELLER_ROLES:  ['новичок', 'птт', 'сп'],
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
const TTL_ = 21600;                                 // кэш данных: 6 ч (максимум CacheService); warmCache() обновляет его после выгрузки продаж (08:20 и 15:20)
const CACHE_EPOCH = 'e12';
let SS_ = null, FRESH_ = false, KB_ = undefined;
function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)); }
function cacheDrop_(key) { try { CacheService.getScriptCache().remove('c:' + CACHE_EPOCH + ':' + key); } catch (e) {} }
function cached_(key, ttl, fn) {
  key = CACHE_EPOCH + ':' + key;
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
    const req = JSON.parse((e.parameter && e.parameter.payload) || e.postData.contents);   // payload — отправка формой (автовыгрузка из DataLens)
    FRESH_ = !!req.fresh;
    if (req.action === 'start') out = start_(req);
    else if (req.action === 'setpin') out = setpin_(req);
    else if (req.action === 'login') out = login_(req);
    else if (req.action === 'data') out = data_(req);
    else if (req.action === 'links_get') out = linksGet_(req);
    else if (req.action === 'tm_sheets') out = tmSheets_(req);
    else if (req.action === 'tm_report') out = tmReport_(req);
    else if (req.action === 'cards_get') out = cardsGet_(req);
    else if (req.action === 'recon_get') out = reconGet_(req);
    else if (req.action === 'pulse_data') out = pulseData_(req);
    else if (req.action === 'news_list') out = newsList_(req);
    else if (req.action === 'news_add') out = newsAdd_(req);
    else if (req.action === 'news_file') out = newsFileGet_(req);
    else if (req.action === 'news_moderate') out = newsModerate_(req);
    else if (req.action === 'links_set') out = linksSet_(req);
    else if (req.action === 'auth_log') out = authLogGet_(req);
    else if (req.action === 'grants_get') out = grantsGet_(req);
    else if (req.action === 'dl_bonus') out = dlBonusGet_(req);
    else if (req.action === 'grants_set') out = grantsSet_(req);
    else if (req.action === 'import_checks') out = importChecks_(req);
    else if (req.action === 'vmr_mine') out = vmrMine_(req);
    else if (req.action === 'attach') out = attachReport_(req);
    else if (req.action === 'notif_list') out = notifList_(req);
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
  const toks = function (s) {
    return String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/[^0-9a-zа-я\- ]+/g, ' ')
      .replace(/\s+/g, ' ').trim().split(' ').filter(Boolean);
  };
  const t = toks(login);
  if (!t.length) return { status: 'none' };
  if (list.some(function (a) { return a.key === norm_(login); })) {
    const ex = list.filter(function (a) { return a.key === norm_(login); });
    if (ex.length === 1) return { status: 'ok', user: ex[0] };
  }
  const items = list.map(function (a) { return { user: a, p: toks(a.name) }; });
  const pick = function (rule) {
    const m = items.filter(rule);
    if (m.length === 1) return { status: 'ok', user: m[0].user };
    return m.length ? { status: 'ambiguous' } : null;
  };
  const like = function (x, p) {           // точное слово или инициал из списка («И.» ~ «Иван»)
    return p.indexOf(x) >= 0 || p.some(function (y) { return y.length === 1 && x.indexOf(y) === 0; });
  };
  return pick(function (it) { return t.every(function (x, i) { return it.p[i] === x; }); })                 // «Фамилия Имя» — начало ФИО
      || pick(function (it) { return t.every(function (x) { return it.p.indexOf(x) >= 0; }); })           // те же слова в любом порядке
      || pick(function (it) { return t.every(function (x) { return like(x, it.p); }); })                  // инициалы вместо полного имени
      || pick(function (it) { return t.length >= 2 && it.p[0] === t[0] && it.p[1] === t[1]; })             // лишнее слово (отчество) не из списка
      || (t.length === 1 ? pick(function (it) { return it.p[0] === t[0]; }) : null)                       // только фамилия, если она одна
      || { status: 'none' };
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
  authLog_(f.user, 'PIN задан', req.dev);
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
  if (hash_(f.user, String(req.pin || '').trim()) !== stored) { fail_(f.user); authLog_(f.user, 'ошибка PIN', req.dev); return { error: 'bad_credentials' }; }
  authLog_(f.user, 'вход', req.dev);
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
  try { maintain_(ss); } catch (e) {}
  try { const vk = 'vis:' + key, cc = CacheService.getScriptCache(); if (!cc.get(vk)) { cc.put(vk, '1', 3600); authLog_(user, 'визит', req.dev); } } catch (e) {}
  const out = cached_('u:' + key, TTL_, function () {
    const d = user.manager ? managerData_(ss, user, access) : sellerData_(ss, user, access);
    d.leaders = leaders_(ss, access); d.updated = new Date().toISOString(); return d;
  });
  const g = user.manager && user.roleTitle !== 'ТМ' ? grantActive_(ss, user) : null;     // временный доступ к отчётам — вне кэша данных
  if (!g) return out;
  const res = Object.assign({}, out, { guest: { until: new Date(g.until).toISOString(), sections: g.secs } });
  if (['dayres', 'dayplan', 'analytics'].some(function (k) { return g.secs.indexOf(k) >= 0; })) {      // «вид ТМ»: данные по всей сети ТМ
    const tmU = access.filter(function (a) { return a.roleTitle === 'ТМ' && a.manager; })[0];
    if (tmU) res.tmData = cached_('u:' + tmU.key, TTL_, function () { const d = managerData_(ss, tmU, access); d.leaders = leaders_(ss, access); d.updated = new Date().toISOString(); return d; });
  }
  return res;
}
// Обновление кэша: не по таймеру, а сразу после выгрузки продаж (importChecks_ ставит разовый запуск через минуту → warmAfterImport_).
// installTrigger() — один раз: убирает старый триггер «каждые 30 минут» и сразу обновляет кэш.
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'warmCache') ScriptApp.deleteTrigger(t); });
  warmCache();
}
function warmAfterImport_() { warmCache(); }
function scheduleWarm_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'warmAfterImport_') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('warmAfterImport_').timeBased().after(60000).create();
}
function warmCache() {
  FRESH_ = true;
  const ss = ss_();
  try { readAccess_(ss); readSalonSheet_(ss); readMonthPlan_(ss); readPlan_(ss); readIncome_(ss); readSchedule_(ss); } catch (e) {}
  try { kpiShopsDrop_(); } catch (e) {}
  const access = readAccess_(ss);
  try { kpiShops_(ss, access); kpiGrades_(ss); } catch (e) {}
  Object.keys(TM_REPORTS).forEach(function (k) { try { tmReportRaw_(ss, k); } catch (e) {} });
  try { salesModel_(ss); detailsGrid_(ss); incomeModel_(ss); scheduleHours_(ss); } catch (e) {}      // расчёты по чекам (вместо сводных таблиц)
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
const NEWS_HEAD = ['ID', 'Дата', 'Автор', 'Роль', 'Заголовок', 'Текст', 'Статус', 'Модератор', 'Вложения'];
const ST_PUB = 'опубликовано', ST_MOD = 'на модерации', ST_REJ = 'отклонено';
function newsSheet_(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_NEWS);
  if (!sh) { sh = ss.insertSheet(CONFIG.SHEET_NEWS); sh.getRange(1, 1, 1, 9).setValues([NEWS_HEAD]); sh.setFrozenRows(1); }
  return sh;
}
function safeCell_(t) { t = String(t); return /^[=+\-@]/.test(t) ? "'" + t : t; }
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
    let files = []; try { files = r[8] ? JSON.parse(r[8]) : []; } catch (e) {}
    items.push({ id: r[0], date: r[1], author: short_(r[2]), role: r[3], title: r[4], text: r[5], status: st, mine: mine, files: files });
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
    const meta = saveFiles_(ss, id, req.files);
    if (meta === null) return { error: 'bad_file' };
    if (!sh.getRange(1, 9).getValue()) sh.getRange(1, 9).setValue('Вложения');
    const rg = sh.getRange(row, 1, 1, 9); rg.setNumberFormat('@');
    rg.setValues([[id, when, u.name, u.roleTitle, safeCell_(title), safeCell_(text), tm ? ST_PUB : ST_MOD, tm ? u.name : '', meta.length ? JSON.stringify(meta) : '']]);
    if (tm) notifyMany_(ss, [{ to: '*', kind: 'news', title: '🚀 Скорее беги в приложение!', text: 'Там новая штука появилась 👀', tab: 'news' }]);
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
      if (d === 'delete') { sh.deleteRow(i + 1); dropFiles_(ss, String(req.id)); }
      else { sh.getRange(i + 1, 7).setValue(d === 'approve' ? ST_PUB : ST_REJ); sh.getRange(i + 1, 8).setValue(u.name); if (d === 'approve') notifyMany_(ss, [{ to: '*', kind: 'news', title: '🚀 Скорее беги в приложение!', text: 'Там новая штука появилась 👀', tab: 'news' }]); }
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


/* ---------- временный доступ к отчётам (выдаёт только ТМ) ----------
   Открываются отчёты ТМ: «Отчёты», «Сотрудники ДЛ», «План на день», полная «Аналитика» и «ВМР».
   НЕ открываются: «Бонус ТМ», «Входы», «Ссылки», «Weekly Pulse», «Доступы». Хранится на листе «Гостевой доступ». */
const GRANTS_SHEET = 'Гостевой доступ';
const GRANT_SECS_ = ['reports', 'dl', 'dayplan', 'dayres', 'cards', 'analytics', 'vmr'];           // разделы, которые ТМ отмечает галочками
const GRANT_DEF_ = ['reports', 'dl', 'dayplan', 'analytics', 'vmr'];                               // старые доступы (без отметки) — как раньше
function grantsSheet_(ss) {
  let sh = ss.getSheetByName(GRANTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(GRANTS_SHEET);
    sh.getRange(1, 1, 1, 8).setValues([['Ключ', 'ФИО', 'Должность', 'Доступ до', 'Выдал', 'Когда выдан', 'Статус', 'Разделы']]).setFontWeight('bold');
    sh.setFrozenRows(1); sh.getRange('D:D').setNumberFormat('dd.MM.yyyy HH:mm'); sh.getRange('F:F').setNumberFormat('dd.MM.yyyy HH:mm');
  }
  return sh;
}
function grantsAll_(ss) {
  const cc = CacheService.getScriptCache(), hit = cc.get('grants:v1');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const sh = grantsSheet_(ss), out = [];
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().forEach(function (r, i) {
      if (!r[0]) return;
      const until = r[3] instanceof Date ? r[3].getTime() : Date.parse(r[3]), when = r[5] instanceof Date ? r[5].getTime() : Date.parse(r[5]);
      out.push({ row: i + 2, key: String(r[0]), name: String(r[1]), role: String(r[2]), until: isNaN(until) ? 0 : until, by: String(r[4]), when: isNaN(when) ? 0 : when, status: String(r[6] || ''),
        secs: (function () { const a = String(r[7] || '').split(',').map(function (x) { return x.trim(); }).filter(function (x) { return GRANT_SECS_.indexOf(x) >= 0; }); return a.length ? a : GRANT_DEF_; })() });
    });
  }
  try { cc.put('grants:v1', JSON.stringify(out), 120); } catch (e) {}
  return out;
}
function grantActive_(ss, u) {
  const now = Date.now();
  const m = grantsAll_(ss).filter(function (g) { return g.key === u.key && norm_(g.status) !== 'отозван' && g.until > now; });
  return m.length ? m.sort(function (a, b) { return b.until - a.until; })[0] : null;
}
function grantHas_(ss, u, sec) { if (u.roleTitle === 'ТМ') return true; const g = u.manager ? grantActive_(ss, u) : null; return !!(g && g.secs.indexOf(sec) >= 0); }
function repAllowed_(ss, u, sec) { return u.roleTitle === 'ТМ' || (sec ? grantHas_(ss, u, sec) : !!(u.manager && grantActive_(ss, u))); }
function grantsGet_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (u.roleTitle !== 'ТМ') return { error: 'forbidden' };
  const now = Date.now();
  const people = readAccess_(ss).filter(function (a) { return a.manager && a.roleTitle !== 'ТМ'; })
    .map(function (a) { return { key: a.key, name: a.name, role: a.roleTitle }; })
    .sort(function (a, b) { return a.name.localeCompare(b.name, 'ru'); });
  const list = grantsAll_(ss).slice().sort(function (a, b) { return b.when - a.when; }).slice(0, 80)
    .map(function (g) { return { key: g.key, name: g.name, role: g.role, until: new Date(g.until).toISOString(), by: g.by, when: g.when ? new Date(g.when).toISOString() : '', active: norm_(g.status) !== 'отозван' && g.until > now, revoked: norm_(g.status) === 'отозван', sections: g.secs }; });
  return { people: people, grants: list, now: new Date(now).toISOString() };
}
function grantsSet_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (u.roleTitle !== 'ТМ') return { error: 'forbidden' };
  const target = readAccess_(ss).filter(function (a) { return a.key === String(req.key || ''); })[0];
  if (!target || !target.manager || target.roleTitle === 'ТМ') return { error: 'bad_request' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = grantsSheet_(ss), now = Date.now();
    if (!sh.getRange(1, 8).getValue()) sh.getRange(1, 8).setValue('Разделы').setFontWeight('bold');
    const secs = (Array.isArray(req.sections) ? req.sections : []).filter(function (x) { return GRANT_SECS_.indexOf(x) >= 0; });
    if (!req.revoke && !secs.length) return { error: 'bad_request' };
    CacheService.getScriptCache().remove('grants:v1');
    const act = grantsAll_(ss).filter(function (g) { return g.key === target.key && norm_(g.status) !== 'отозван' && g.until > now; });
    if (req.revoke) {
      act.forEach(function (g) { sh.getRange(g.row, 4).setValue(new Date(now)); sh.getRange(g.row, 7).setValue('отозван'); });
      CacheService.getScriptCache().remove('grants:v1');
      return { ok: true };
    }
    let until;
    if (req.untilDate) {                                              // «до даты включительно»: до конца этого дня по Москве
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(req.untilDate))) return { error: 'bad_request' };
      until = Date.parse(req.untilDate + 'T23:59:59+03:00');
      if (isNaN(until) || until <= now || until > now + 120 * 864e5) return { error: 'bad_date' };
    } else {
      const dn = Number(req.days);
      if (!(dn > 0)) return { error: 'bad_request' };
      until = now + Math.min(60, dn) * 864e5;
    }
    if (act.length) { sh.getRange(act[0].row, 4).setValue(new Date(until)); sh.getRange(act[0].row, 5).setValue(u.name); sh.getRange(act[0].row, 6).setValue(new Date(now)); sh.getRange(act[0].row, 8).setValue(secs.join(',')); }
    else sh.appendRow([target.key, target.name, target.roleTitle, new Date(until), u.name, new Date(now), 'активен', secs.join(',')]);
    CacheService.getScriptCache().remove('grants:v1');
    try { notifyMany_(ss, [{ to: target.key, kind: 'access', title: '🔓 Открыт доступ к отчётам', text: 'До ' + Utilities.formatDate(new Date(until), 'Europe/Moscow', 'dd.MM.yyyy HH:mm') + ' вам доступны отчёты ТМ (Отчёты, Сотрудники ДЛ и др.).', tab: 'reports' }]); } catch (e) {}
    return { ok: true, until: new Date(until).toISOString() };
  } finally { lock.releaseLock(); }
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
  const sh = ss.getSheetByName(CONFIG.SHEET_SCHEDULE), out = {}, cal = { today: 0, dim: 0, month: 0, year: 0 };
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
  cal.today = today; cal.dim = (month && year) ? new Date(year, month, 0).getDate() : 0; cal.month = month || 0; cal.year = year || 0;
  const shift = /^(i{1,3}|iv|v|vi{0,3}|ix|x)$/i, time = /^\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}$/;
  // блоки салонов: заголовок — строка с «№» в колонке C, название салона в A; ниже расшифровка смен (I — 10:00 - 19:00 …) именно этого салона
  const blocks = []; let cur = null;
  for (let r = 0; r < v.length; r++) {
    if (String(v[r][2]).trim() === '№' && String(v[r][0]).trim()) { cur = { salon: String(v[r][0]).trim(), legend: {} }; blocks.push(cur); v[r].blk = cur; continue; }
    if (cur && shift.test(String(v[r][2]).trim()) && time.test(String(v[r][4]).trim())) cur.legend[String(v[r][2]).trim().toUpperCase()] = String(v[r][4]).replace(/\s+/g, ' ').trim();
    v[r].blk = cur;
  }
  for (let r = dayRow + 1; r < v.length; r++) {
    const name = String(v[r][4] || '').trim();
    if (!/^\d+$/.test(String(v[r][2]).trim()) || name.split(' ').length < 2) continue;
    const blk = v[r].blk, o = out[norm_(name)] = out[norm_(name)] || { planned: 0, done: 0, left: 0, days: {} };
    for (let c = 9; c < v[r].length; c++) {
      const d = c - 8; if (d > 31) break;
      const code = String(v[r][c]).trim().replace(/\u0406/g, 'I').replace(/\u0425/g, 'X').replace(/\u0456/g, 'i').replace(/\u0445/g, 'x'); if (!code) continue;
      const isShift = shift.test(code);
      if (isShift || /^[^\d\s]{1,3}$/.test(code)) (o.days[d] = o.days[d] || []).push([blk ? blk.salon : '', code.toUpperCase(), isShift && blk ? (blk.legend[code.toUpperCase()] || '') : '']);
      if (!isShift) continue;
      o.planned++;
      if (d < today) o.done++; else o.left++;            // сегодняшняя смена ещё не попала в факт продаж — считаем её «осталось»
    }
  }
  return { people: out, cal: cal };
}
// График на 3 дня (сегодня, завтра, послезавтра): смена, время, точка; «З» (замена) или выход на другую точку → «Замена» и точка работы
const CODE_NAMES_ = { 'О': 'Отпуск', 'Б': 'Больничный', 'У': 'Уволен', 'Т': 'Тренинг', 'ГУ': 'Генуборка', 'И': 'Инвентаризация', 'С': 'Стажировка', 'М': 'Командировка' };
function sameSalon_(a, b) {
  const opOf = function (t) { const m = norm_(t).match(/(^|\s)(мтс|мф|блн|т2|тд)(\s|$)/); return m ? (m[2] === 'тд' ? 'т2' : m[2]) : ''; };
  if (opOf(a) && opOf(b) && opOf(a) !== opOf(b)) return false;                    // «Ямашева 71а МТС» ≠ «Ямашева 71а БЛН»
  const clean = function (t) { return norm_(t).replace(/\b(кзн|блн|тц|трц|тк)\b/g, ' ').replace(/\s+/g, ' ').trim(); };
  const x = clean(a), y = clean(b); if (!x || !y) return false;
  if (x === y) return true;
  const sh = x.length <= y.length ? x : y, lg = x.length <= y.length ? y : x;      // вхождение только целыми словами: «Зорге 7» ≠ «Зорге 75»
  return sh.length > 4 && (' ' + lg + ' ').indexOf(' ' + sh + ' ') >= 0;
}
function next3_(sched, person, mainSalon) {
  const cal = sched.cal; if (!cal.month || !cal.year) return [];
  const WD = ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'], out = [];
  // «сегодня» считаем в момент запроса (кэш графика живёт долго), по московскому времени
  const nowS = Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyyy-M-d').split('-').map(Number);
  for (let i = 0; i < 3; i++) {
    const dt = new Date(Date.UTC(nowS[0], nowS[1] - 1, nowS[2] + i)), y = dt.getUTCFullYear(), m = dt.getUTCMonth() + 1, d = dt.getUTCDate();
    const o = { date: ('0' + d).slice(-2) + '.' + ('0' + m).slice(-2), wd: WD[dt.getUTCDay()], day: i === 0 ? 'сегодня' : (i === 1 ? 'завтра' : 'послезавтра') };
    if (y !== cal.year || m !== cal.month) { o.kind = 'nodata'; out.push(o); continue; }       // лист «График» — на другой месяц
    const e = (person && person.days && person.days[d]) || [];
    const isSh = function (x) { return /^(I|II|III|IV|V|VI|VII|VIII|IX|X)$/.test(x[1]); };
    const shifts = e.filter(isSh), zs = e.filter(function (x) { return x[1] === 'З'; });
    const oth = e.filter(function (x) { return x[1] !== 'З' && !isSh(x); });
    if (shifts.length) {
      const s0 = shifts[0]; o.kind = 'shift'; o.shift = s0[1]; o.time = s0[2]; o.salon = s0[0];
      if (zs.length || (mainSalon && !sameSalon_(s0[0], mainSalon))) { o.repl = true; o.from = zs.length ? zs[0][0] : mainSalon; }
    } else if (zs.length) { o.kind = 'repl'; o.repl = true; o.from = zs[0][0]; }
    else if (oth.length) { o.kind = 'other'; o.label = CODE_NAMES_[oth[0][1]] || ('Отметка «' + oth[0][1] + '» в графике'); }
    else o.kind = 'off';
    out.push(o);
  }
  return out;
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
  const v0 = summaryGrid_(ss);                                  // «Общий свод» считаем на сайте
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
  return { name: CONFIG.SHEET_SUMMARY, head: trim(head), rows: trim(body), found: body.length };
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
    schedule: mySched ? { planned: mySched.planned, done: mySched.done, left: mySched.left } : null, cal: sched.cal,
    next3: next3_(sched, mySched, primarySalon_(ss, meKey, me.tt || (kpiShops_(ss, access)[meKey] || [])[0] || '')),
    pointRes: salonResults_(ss, primarySalon_(ss, meKey, me.tt || (kpiShops_(ss, access)[meKey] || [])[0] || '')),
    bonus: { total: total, shifts: shifts, avgDaily: shifts ? total / shifts : 0, days: worked, groups: inc.groups || null, groupsTotal: inc.gtotal === undefined ? null : inc.gtotal },
    rating: rating, ratings: ratingsFor_(sellers, meKey),
  };
}


/* ---------- «Результаты торговой точки»: цели салона продавца ---------- */
// Основной салон продавца — где у него больше всего плановых часов («Свод графика»); при равенстве — колонка «Торговая точка».
function primarySalon_(ss, personKey, fallback) {
  try {
    const h = scheduleHours_(ss)[personKey]; if (!h) return fallback || '';            // часы по салонам считаем из «График» (раньше — сводная «Свод графика»)
    let best = '', bh = -1, tie = false;
    Object.keys(h).forEach(function (sl) { if (h[sl] > bh) { bh = h[sl]; best = sl; tie = false; } else if (h[sl] === bh && bh > 0) tie = true; });
    if (tie && fallback) return fallback;
    return bh > 0 ? best : (fallback || '');
  } catch (e) { return fallback || ''; }
}
// Универсальный шаблон «Результаты торговой точки» (как лист «Шаблон»): ПЛАН — основная таблица, лист «План»;
// ФАКТ — книга KPI ОСС («Экономические KPI» и «Операторские таргеты»); остаток / на день / вып.% / прогноз% — формулы шаблона.
// Колонки заданы буквами (проверены по факту на салоне КЗН Зорге 75 БЛН). eco = «Экономические KPI», op = «Операторские таргеты».
const RES_SPEC_ = [
  { title: 'Экономика', rows: [
    { id: 'phones', label: 'Телефоны руб', plan: 'I',  fact: ['eco', 'AE'] },
    { id: 'acc',    label: 'Аксессуары',   plan: 'J',  fact: ['eco', 'AK'] },
    { id: 'serv',   label: 'Услуги',       plan: 'M',  fact: ['eco', 'AR'] },
    { id: 'accserv', label: 'АКС+Услуги',  sum: ['acc', 'serv'] },
    { id: 'ins',    label: 'Страховки',    plan: 'N',  fact: ['eco', 'BI'] },
    { id: 'insShare', label: 'Страховки доля', share: ['ins', 'phones'], planConst: 9 },
    { id: 'ya',     label: 'ЯА',           plan: 'AE', fact: ['op', 'HF'] },
    { id: 'attach', label: 'Аттачмент',    fact: ['eco', 'BP'], planConst: 2, pctUnit: true },
    { id: 'credRev', label: 'Кредитная выручка', plan: 'O', fact: ['eco', 'BR'] },
    { id: 'credApp', label: 'Кредитная заявка',  planFact: ['eco', 'BY'], fact: ['eco', 'BZ'] },
    { id: 'sellout', label: 'Sell out и трудный сток', plan: 'Q', fact: ['eco', 'CD'] } ] },
  { title: '@OPER', rows: [
    { id: 'sim1', label: 'Сим декада 1', plan: 'G', fact: ['op', 'K'] },
    { id: 'sim2', label: 'Сим декада 2', plan: 'H', fact: ['op', 'P'] },
    { id: 'simT', label: 'Сим итого',    plan: 'F', sum: ['sim1', 'sim2'], sumPlanOwn: true },
    { id: 'mnp',  label: 'МНП',          plan: 'S', fact: ['op', 'BH'] },
    { id: 'shpd', label: 'ШПД',          plan: 'AV', fact: ['op', 'GH'] },
    { id: 'act',  label: 'Актив 1-15',   plan: 'AS', fact: ['op', 'HB'] } ] },
  { title: 'Фокусное направление', rows: [
    { id: 'combo', label: 'Комбо М',     plan: 'BK', fact: ['eco', 'BF'], planFact: ['eco', 'BE'] },
    { id: 'bp',    label: 'БП и Пакетные услуги', fact: ['eco', 'AU'] } ] },
];
function colIdx_(L) { let n = 0; String(L).split('').forEach(function (ch) { n = n * 26 + ch.charCodeAt(0) - 64; }); return n - 1; }
function resNum_(v) { return typeof v === 'number' ? v : num_(v); }
function readResSheet_(book, name, key) {
  return cached_(key, TTL_, function () {
    const sh = book && sheetByName_(book, name); if (!sh) return [];
    return sh.getDataRange().getValues().map(function (r) { return r.map(function (c) { return (c instanceof Date) ? '' : c; }); });
  });
}
function resFindRow_(grid, nameCol, salon, from) {
  const key = norm_(salon);
  for (let r = from || 0; r < grid.length; r++) if (norm_(grid[r][nameCol]) === key) return grid[r];
  return null;
}
function salonResults_(ss, salon) {
  if (!salon) return { found: false, salon: '' };
  try {
    const plan = readResSheet_(ss, 'План', 'res:plan'), prow = resFindRow_(plan, colIdx_('E'), salon, 2);
    if (!prow) return { found: false, salon: salon };
    const kb = kpiBook_(ss);
    const grids = { eco: readResSheet_(kb, 'Экономические KPI', 'res:eco'), op: readResSheet_(kb, 'Операторские таргеты', 'res:op') };
    const frow = { eco: resFindRow_(grids.eco, colIdx_('H'), salon, 56), op: resFindRow_(grids.op, colIdx_('H'), salon, 56) };
    const cal = readSchedule_(ss).cal, now = new Date();
    const dim = cal.dim || new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
    const day = cal.today >= 99 ? dim : (cal.today || now.getDate());   // 99 — график прошлого месяца
    const left = Math.max(1, dim - day + 1);
    const val = {};   // id -> { plan, fact }
    const out = [];
    RES_SPEC_.forEach(function (blk) {
      const rows = [];
      blk.rows.forEach(function (sp) {
        let p = null, f = null;
        if (sp.sum) {
          const a = sp.sum.map(function (i) { return val[i]; });
          if (a.every(function (x) { return x && x.fact != null; })) f = a.reduce(function (t, x) { return t + x.fact; }, 0);
          if (sp.sumPlanOwn) p = sp.plan ? resNum_(prow[colIdx_(sp.plan)]) : null;
          else if (a.every(function (x) { return x && x.plan != null; })) p = a.reduce(function (t, x) { return t + x.plan; }, 0);
        } else if (sp.share) {
          const ins = val[sp.share[0]], ph = val[sp.share[1]];
          if (ins && ph && ins.fact != null && ph.fact > 0) f = ins.fact / ph.fact * 100;
          p = sp.planConst;
        } else {
          if (sp.plan) p = resNum_(prow[colIdx_(sp.plan)]);
          if (p == null && sp.planFact && frow[sp.planFact[0]]) p = resNum_(frow[sp.planFact[0]][colIdx_(sp.planFact[1])]);
          if (sp.planConst != null) p = sp.planConst;
          if (sp.fact && frow[sp.fact[0]]) f = resNum_(frow[sp.fact[0]][colIdx_(sp.fact[1])]);
          if (sp.pctUnit && f != null && Math.abs(f) <= 1.5) f = f * 100;   // доля из листа приходит дробью
        }
        val[sp.id] = { plan: p, fact: f };
        if (p == null && f == null) return;                                   // нет ни плана, ни факта — строку не показываем
        const unit = (sp.pctUnit || sp.share) ? 'pct' : '';
        const row = { label: sp.label, unit: unit, plan: p, fact: f, rest: null, perDay: null, pct: null, prog: null };
        if (p != null && f != null) {
          row.rest = sp.share ? (val[sp.share[1]].fact * 0.06 - val[sp.share[0]].fact) : p - f;   // «Страховки доля»: остаток = телефоны×6% − страховки (как в шаблоне)
          row.perDay = row.rest / left;
          row.pct = p ? f / p * 100 : null;
          row.prog = row.pct != null && day ? row.pct / day * dim : null;
        }
        rows.push(row);
      });
      if (rows.length) out.push({ title: blk.title === '@OPER' ? ('Операторка ' + (/блн/i.test(salon) ? 'БЛН' : /мтс/i.test(salon) ? 'МТС' : /т2|теле2/i.test(salon) ? 'Т2' : /мф|мегафон/i.test(salon) ? 'МФ' : '')).trim() : blk.title, rows: rows });
    });
    const dateTxt = (frow.eco && grids.eco[1] && String(grids.eco[1][7] || '').trim()) || '';
    return { found: true, salon: salon, label: salon, daysPassed: day, daysTotal: dim, date: dateTxt || Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy'),
             noFact: !frow.eco && !frow.op, blocks: out, updated: new Date().toISOString() };
  } catch (e) { return { found: false, salon: salon, error: String(e && e.message || e) }; }
}

/* ---------- данные калькулятора бонуса РТТ/УТТ/ТМ: экономический и операторский блоки из KPI ОСС ---------- */
// «Экономические KPI» (экономика), «Операторские таргеты» (декады, HB++), «Операторские РДО ММБ» (прогноз SIM, Gross AS 1–15, выручка 1М).
// Колонки ищутся по заголовкам, а не по буквам — вставка столбцов в таблице не ломает привязку.
function pctVal_(v) {                                          // '54%' → 0.54 ; '1 762' → 1762 ; '#DIV/0!', '' → null
  if (typeof v === 'number') return v;
  let t = String(v == null ? '' : v).replace(/[\s ]/g, '').replace(',', '.');
  if (!t || t.charAt(0) === '#') return null;
  const p = /%$/.test(t); if (p) t = t.slice(0, -1);
  if (!/^-?\d+(\.\d+)?$/.test(t)) return null;
  const n = parseFloat(t); return p ? n / 100 : n;
}
function ecoCols_(g) {                                         // индексы колонок «Экономических KPI» по группе + подзаголовку
  let top = -1; for (let r = 0; r < Math.min(g.length, 30); r++) if (String(g[r][8] || '').trim() === 'Трафик') { top = r; break; }
  if (top < 0) return null;
  const grp = g[top], sub = g[top + 1] || [], cols = []; let cur = '';
  for (let i = 0; i < grp.length; i++) { if (String(grp[i] || '').trim()) cur = String(grp[i]).trim(); cols.push({ g: cur, s: String(sub[i] || '').trim() }); }
  const find = function (gr, sb, nth) { let k = 0; for (let i = 0; i < cols.length; i++) if (cols[i].g === gr && (sb == null || cols[i].s === sb)) { if (k === (nth || 0)) return i; k++; } return -1; };
  return { telPlan: find('Тел, руб', 'План, руб'), telFact: find('Тел, руб', 'Факт, руб'), tel: find('Тел, руб', 'Прогн, %'),
    accPlan: find('Аксесс', 'План, руб'), acc: find('Акс + Усл', 'Прогн, %'), insPlan: find('Страховки', 'План, руб'), ins: find('Страховки', 'Вып доли %'),
    attach: find('Аттачмент', ''), credPlan: find('Кредиты', 'План, руб'), cred: find('Кредиты', 'Выпол-е,%', 0),
    sellPlan: find('Sell_Out + TI', 'План'), sell: find('Sell_Out + TI', 'Выполнение'),
    comboPlan: find('Комбо', 'План шт Комбо'), combo: find('Комбо', 'Прогноз вып доли Комбо'),
    trPlan: find('Трафик', 'План, чел'), trFact: find('Трафик', 'Факт, чел'), trDev: find('Трафик', 'Откл, %'),
    cvPlan: find('Конв', 'План, %'), cvFact: find('Конв', 'Факт, %') };
}
function opColsTargets_(g) {                                   // «Операторские таргеты»: колонки по подписям второй строки заголовка
  let h = -1; for (let r = 0; r < Math.min(g.length, 120); r++) if (String(g[r][0] || '').trim() === 'ТМ' && String(g[r][7] || '').trim() === 'Наименование ТТ') { h = r; break; }
  if (h < 0) return null;
  const lab = g[h + 1] || [], top = g[h] || [], res = { h: h };
  for (let i = 0; i < lab.length; i++) {
    const t = String(lab[i] || '').trim().toLowerCase();
    if (t === '1 декада') res.dec1 = i; else if (t === '2 декада') res.dec2 = i;
    else if (t === 'сим') res.sim = i; else if (t === 'hb++') { res.hb = i - 2; res.hbProg = i; }
    else if (/^выручка 1м\s*бл/.test(t)) res.rev = i;
  }
  res.top = top; return res;
}
function rdoCols_(g) {                                         // «Операторские РДО ММБ»: код ТТ → прогноз SIM (GI), Gross AS 1–15, выручка 1М (Билайн)
  let hdr = -1; for (let r = 0; r < g.length; r++) { const a = String(g[r][0] || '').trim(), c = String(g[r][2] || ''), e = String(g[r][4] || ''); if (a === 'ТМ' && /оператор/i.test(c) && /название тт/i.test(e)) { hdr = r; break; } }
  if (hdr < 1) return null;
  const lab = g[hdr - 1] || [], sub = g[hdr] || [], res = { hdr: hdr };
  const pct = function (i) { for (let j = i; j < Math.min(i + 4, sub.length); j++) if (/%\s*прогноз/i.test(sub[j] || '')) return j; return -1; };
  for (let i = 5; i < lab.length; i++) {
    const t = String(lab[i] || '').trim(), q = /["«»]/.test(t);
    if (/gross as 1-15/i.test(t)) res.act = pct(i); else if (/^["«]?\s*выручка 1м/i.test(t) && !q) res.rev = pct(i);
  }
  for (let i = 0; i < sub.length; i++) if (/прогноз gi/i.test(sub[i] || '')) { res.gi = i; break; }
  return res;
}
function kpiPrefill_(ss, names) {
  const out = { stores: {}, eco: null, asOf: '' };
  const kb = kpiBook_(ss); if (!kb || !names || !names.length) return out;
  const eco = readResSheet_(kb, 'Экономические KPI', 'res:eco'), tg = readResSheet_(kb, 'Операторские таргеты', 'res:op'), rdo = readResSheet_(kb, 'Операторские РДО ММБ', 'kpi:rdo');
  const ec = ecoCols_(eco), tc = opColsTargets_(tg), rc = rdoCols_(rdo);
  out.asOf = String((eco[1] && eco[1][7]) || '').trim();
  const rows = [];
  names.forEach(function (n) {
    const er = resFindRow_(eco, 7, n, 50), tr = tc ? resFindRow_(tg, 7, n, tc.h) : null;
    let rr = null; if (rc) { const code = er ? String(er[4] || '').trim() : ''; for (let r = rc.hdr + 1; r < rdo.length; r++) { const x = rdo[r]; if ((code && String(x[3] || '').trim() === code) || (!code && norm_(x[4]) === norm_(n))) { rr = x; break; } } }
    const g = function (row, i) { return row && i != null && i >= 0 ? pctVal_(row[i]) : null; };
    const st = { name: n, brand: er ? String(er[3] || '').trim() : '', code: er ? String(er[4] || '').trim() : '' };
    st.sim = g(rr, rc && rc.gi); if (st.sim == null) st.sim = g(tr, tc && tc.sim);
    st.dec1 = g(tr, tc && tc.dec1); st.dec2 = g(tr, tc && tc.dec2);
    st.hb = g(tr, tc && tc.hb);                                           // доля HB++, порог 25 %
    st.rev = g(rr, rc && rc.rev); if (st.rev == null) st.rev = g(tr, tc && tc.rev);
    st.act = g(rr, rc && rc.act);
    out.stores[norm_(n)] = st;
    if (er && ec) rows.push(er);
  });
  if (ec && rows.length) {                                                  // экономика по салонам руководителя: среднее, взвешенное планом
    const w = function (vk, wk) { let sw = 0, sv = 0; rows.forEach(function (r) { const v = pctVal_(r[ec[vk]]), p = pctVal_(r[ec[wk]]); if (v != null && p > 0) { sv += v * p; sw += p; } }); return sw ? sv / sw : null; };
    out.eco = { tel: w('tel', 'telPlan'), sell: w('sell', 'sellPlan'), acc: w('acc', 'accPlan'), ins: w('ins', 'insPlan'), credit: w('cred', 'credPlan'), attach: w('attach', 'telFact'), combo: w('combo', 'comboPlan') };
  }
  return out;
}

/* ---------- Weekly Pulse (ТМ): инвентаризации и блоки «Экономических KPI» (L4L, M2M, трафик, конверсия, средняя цена, кредит) ---------- */
// ID таблицы инвентаризаций можно переопределить в свойствах скрипта: INV_SPREADSHEET_ID.
const INV_SPREADSHEET_ID = '1TxmgRko5O3h5gSPKUT-AlBRP4xZKbmQ9tSGI9Sbuzzs';
function pulseData_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (u.roleTitle !== 'ТМ') return { error: 'forbidden' };
  return cached_('pulse:' + u.key, TTL_, function () {
    const out = { inv: null, eco: null };
    try {
      const id = PropertiesService.getScriptProperties().getProperty('INV_SPREADSHEET_ID') || INV_SPREADSHEET_ID;
      const sh = SpreadsheetApp.openById(id).getSheetByName('Лист1') || SpreadsheetApp.openById(id).getSheets()[0];
      const g = sh.getDataRange().getDisplayValues().filter(function (r) { return r.some(function (c) { return String(c).trim(); }); });
      out.inv = g.map(function (r) { return r.slice(0, 40); });
    } catch (e) { out.invErr = String(e && e.message || e); }
    try {
      const kb = kpiBook_(ss), eco = kb ? readResSheet_(kb, 'Экономические KPI', 'res:eco') : null;
      if (eco) {
        let hdr = -1; for (let r = 0; r < Math.min(eco.length, 60); r++) if (String(eco[r][0] || '').trim() === 'ТМ' && String(eco[r][3] || '').trim() === 'Бренд') { hdr = r; break; }
        if (hdr >= 0) {
          const grp = eco[hdr], sub = eco[hdr + 1] || [], cols = []; let cur = '';
          for (let i = 0; i < grp.length; i++) { if (String(grp[i] || '').trim()) cur = String(grp[i]).trim(); cols.push(cur + '|' + String(sub[i] || '').trim()); }
          const rows = [];
          for (let r = hdr + 2; r < eco.length; r++) { const nm = String(eco[r][7] || '').trim(); if (!nm || /^закрыт/i.test(nm)) continue; rows.push(eco[r].map(function (c) { return typeof c === 'number' ? c : String(c == null ? '' : c); })); }
          out.eco = { cols: cols, rows: rows, dim: Number(eco[1] && eco[1][0]) || 0, passed: Number(eco[1] && eco[1][4]) || 0, asOf: String((eco[1] && eco[1][7]) || '').trim() };
        }
      }
    } catch (e) { out.ecoErr = String(e && e.message || e); }
    return out;
  });
}

/* ---------- «Салоны рядом»: смежные салоны (тот же адрес, другой бренд: БЛН / МТС / МФ / Т2) + трафик, прогнозы — из KPI ОСС ---------- */
const NEAR_BRANDS_ = ['блн', 'мтс', 'мф', 'т2'];
function brandOf_(t) { const m = norm_(t).replace(/(^|\s)тд(?=\s|$)/g, ' т2').match(/(^|\s)(блн|мтс|мф|т2)(?=\s|$)/); return m ? m[2] : ''; }
function baseOf_(t) { return norm_(t).replace(/(^|\s)(блн|мтс|мф|т2|тд)(?=\s|$)/g, ' ').replace(/\s+/g, ' ').trim(); }
function nearSalons_(ss, names) {
  const kb = kpiBook_(ss); if (!kb || !names || !names.length) return null;
  const eco = readResSheet_(kb, 'Экономические KPI', 'res:eco'), ec = ecoCols_(eco); if (!ec) return null;
  let hdr = -1; for (let r = 0; r < eco.length; r++) if (String(eco[r][0] || '').trim() === 'ТМ' && String(eco[r][3] || '').trim() === 'Бренд') { hdr = r; break; }
  if (hdr < 0) return null;
  const dim = Number(eco[1] && eco[1][0]) || 0, passed = Number(eco[1] && eco[1][4]) || 0;   // A2 — дней в месяце, E2 — прошло дней
  const g = function (row, i) { return i != null && i >= 0 ? pctVal_(row[i]) : null; };
  const rows = [];
  for (let r = hdr + 2; r < eco.length; r++) {
    const nm = String(eco[r][7] || '').trim(); if (!nm || /^закрыт/i.test(nm) || nm === 'ТЦ' || nm === 'Стрит') continue;
    const br = brandOf_(String(eco[r][3] || '').trim()) || brandOf_(nm); if (!br) continue;
    rows.push({ name: nm, brand: br, base: baseOf_(nm), row: eco[r] });
  }
  const item = function (x) {
    const w = x.row, fact = g(w, ec.trFact), plan = g(w, ec.trPlan);
    return { name: x.name, brand: x.brand, code: String(w[4] || '').trim(), mgr: String(w[2] || '').trim(),
      tr: { plan: plan, fact: fact, dev: g(w, ec.trDev), fc: (fact != null && passed > 0 && dim > 0) ? Math.round(fact / passed * dim) : null },
      cv: { plan: g(w, ec.cvPlan), fact: g(w, ec.cvFact) },
      fc: { tel: g(w, ec.tel), acc: g(w, ec.acc), ins: g(w, ec.ins), combo: g(w, ec.combo) } };
  };
  const out = [], seen = {};
  names.forEach(function (n) {
    const own = rows.filter(function (x) { return norm_(x.name) === norm_(n); })[0];
    const base = own ? own.base : baseOf_(n), ob = own ? own.brand : brandOf_(n);
    if (seen[base + '|' + ob]) return; seen[base + '|' + ob] = 1;
    const near = rows.filter(function (x) { return x.base === base && x.brand !== ob; });
    if (!near.length) return;
    out.push({ salon: n, own: own ? item(own) : null, near: near.map(item) });
  });
  return { groups: out, asOf: String((eco[1] && eco[1][7]) || '').trim(), dim: dim, passed: passed };
}

/* ---------- руководитель (РТТ / УТТ / ТМ) ---------- */
function managerData_(ss, sess, access) {
  const all = readSalonSheet_(ss);                         // { салон: { phones:{...}, acc:{...} ... } }
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
      const x = all.salons[key][d];
      const mp = mplan[norm_(key)] && mplan[norm_(key)][d];
      let plan = mp > 0 ? mp : 0, src = 'месяц';
      if (!plan) { plan = x.plan; src = 'на дату'; }               // запасной вариант: «факт + отставание» из деталей продаж
      return { key: d, label: DIR_LABELS[d], unit: DIR_UNITS[d], plan: plan || 0, planSrc: src, fact: x.fact || 0,
        planDay: plan && all.daysInMonth ? plan / all.daysInMonth : 0,
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
    kpi: (function () { try { return kpiPrefill_(ss, names); } catch (e) { return { stores: {}, eco: null, error: String(e && e.message || e) }; } })(),   // экономика и операторка из KPI ОСС для калькулятора бонуса
    near: (function () { try { return nearSalons_(ss, names); } catch (e) { return null; } })(),   // смежные салоны другого бренда + трафик и прогнозы из KPI ОСС
    canLinks: sess.roleTitle === 'ТМ',
    prefill: me ? { grade: grade, hours: me.hours.fact, att: me.quality.attach, vmr: me.quality.vmr, credits: me.quality.credits } : { grade: grade },
  };
}

function readSalonSheet_(ss) { return cached_('salons', TTL_, function () { return readSalonSheetRaw_(ss); }); }
function readSalonSheetRaw_(ss) {
  const v = detailsGrid_(ss);                                // «Детали продаж салоны» — считаем из чеков, без сводных таблиц
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
// ВМР: сырые проверки берутся из отдельной таблицы (лист «Лист1»). ID можно переопределить в свойствах скрипта: VMR_SPREADSHEET_ID.
const VMR_SPREADSHEET_ID = '1Y0erbIsvDKxKALsvk2MC3cW0bUcgN2XfhA2uoPzyzqY';
function vmrSheet_() {
  try {
    const id = PropertiesService.getScriptProperties().getProperty('VMR_SPREADSHEET_ID') || VMR_SPREADSHEET_ID;
    const book = SpreadsheetApp.openById(id), all = book.getSheets();
    return all.filter(function (x) { return norm_(x.getName()) === norm_('Лист1'); })[0] || all[0] || null;
  } catch (e) { return null; }   // нет доступа — используем лист «отчет по вмр ММБ» в основной таблице
}
const TM_REPORTS = { details: 'Детали продаж салоны', summary: 'общий свод', fact: 'Факт продаж', dl: 'Сотрудники ДЛ', sim: 'Детализация СИМ', recon: 'Сопоставления 1с и рег', vmr: 'отчет по вмр ММБ' };
function tmReport_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (!TM_REPORTS[req.name]) return { error: 'bad_request' };
  if (!(req.name === 'dl' && u.manager) && !repAllowed_(ss, u, req.name === 'dl' ? 'dl' : req.name === 'vmr' ? 'vmr' : 'reports')) return { error: 'forbidden' };   // «Сотрудники ДЛ» — и РТТ/УТТ (только свои)   // ТМ или временный доступ (по отмеченным разделам)
  if (req.name === 'vmr' && u.roleTitle !== 'ТМ') { const m = vmrMine_(req); return m && m.grid ? Object.assign({ at: new Date().toISOString() }, m) : m; }   // ВМР у РТТ/УТТ — только их ТТ (в т.ч. при выданном доступе)
  const rep = tmReportRaw_(ss, req.name);
  if (req.name !== 'dl' || u.roleTitle === 'ТМ' || !rep.grid) return rep;                // «Сотрудники ДЛ» у РТТ/УТТ — только их сотрудники
  const shops = kpiShops_(ss, readAccess_(ss))[u.key] || [];
  const grid = rep.grid.filter(function (r, i) { return i < 2 || shops.some(function (sh) { return salonShort_(sh) === String(r[1] || '').trim() || sameSalon_(String(r[1] || ''), sh); }); });
  return Object.assign({}, rep, { grid: grid });
}
function tmReportRaw_(ss, name) {
  const title = TM_REPORTS[name];
  return cached_('rep:' + name, TTL_, function () {
    const virt = { details: detailsGrid_, summary: summaryGrid_, fact: factGrid_, sim: simGrid_, dl: dlGrid_ }[name];      // эти страницы считаем на сайте из чеков
    if (virt) { let g0 = virt(ss); if (name === 'summary') g0 = summaryMonthPlan_(ss, g0); return { name: title, grid: g0, at: new Date().toISOString(), calc: true }; }       // «План на день» → «План на месяц»
    let sh = null;
    if (name === 'vmr') sh = vmrSheet_();                                   // выгрузка ВМР — напрямую из таблицы по ссылке (см. VMR_SPREADSHEET_ID)
    if (!sh) sh = sheetByName_(ss, title);
    if (!sh && name === 'vmr') sh = ss.getSheets().filter(function (x) { return /отч[её]т.*вмр/i.test(x.getName()); })[0] || null;   // «отчет по вмр ММБ» — допускаем небольшие отличия в названии
    if (!sh) return { missing: true, name: title };
    const rows = Math.min(sh.getLastRow(), name === 'vmr' ? 6000 : 800), cols = Math.min(sh.getLastColumn(), 130);
    if (rows < 1 || cols < 1) return { missing: true, name: title };
    let g = sh.getRange(1, 1, rows, cols).getDisplayValues();
    let lr = g.length; while (lr > 0 && !g[lr - 1].some(function (c) { return String(c).trim(); })) lr--;
    let lc = 0; g.forEach(function (r) { for (let i = r.length; i > lc; i--) if (String(r[i - 1]).trim()) { lc = i; break; } });
    g = g.slice(0, lr).map(function (r) { return r.slice(0, lc); });
    return { name: sh.getName(), grid: g, at: new Date().toISOString() };
  });
}


/* ---------- «Сопоставления 1С и рег»: те же формулы, что на листе, но сразу по каждому салону и итогом ---------- */
// Источники (как в формулах листа): Сим 1С — «Детали продаж салоны»; Трафик — лист «Трафик»; План — «Цели» (план по дням);
// плановая конверсия — «Цели» (блок «План по конверсии»); регистрации — «Детали оператор» (блоки «Регистрации», «Не актив», «Актив»).
// Все блоки ищутся по заголовкам, а не по номерам строк и столбцов — привязка по оператору «скакать» не будет.
function reconGet_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (!u.manager) return { error: 'forbidden' };
  return cached_('recon:' + u.key, TTL_, function () { return reconBuild_(ss, u); });
}
function dateParts_(s) { const m = /^(\d{1,2})\.(\d{2})\.(\d{4})/.exec(String(s).trim()); return m ? { d: +m[1], m: +m[2], y: +m[3] } : null; }
function readTrafficByName_(ss) {
  const sh = sheetByName_(ss, 'Трафик'); if (!sh) return {};
  const v = sh.getDataRange().getDisplayValues(), out = {};
  let hr = -1; for (let r = 0; r < Math.min(v.length, 12); r++) if (norm_(v[r][1]) === 'магазин') { hr = r; break; }
  for (let r = (hr < 0 ? 1 : hr + 1); r < v.length; r++) {
    const nm = String(v[r][1] || '').trim(); if (!nm || norm_(nm) === 'название тт') continue;
    out[norm_(nm)] = v[r].slice(2, 33).map(num_);
  }
  const head = v[1] || []; if (hr < 0 && head[1]) out[norm_(head[1])] = head.slice(2, 33).map(num_);   // строка итога сверху (фамилия руководителя)
  if (v[1] && String(v[1][1]).trim()) out[norm_(v[1][1])] = v[1].slice(2, 33).map(num_);
  return out;
}
function readGoals_(ss) {
  const sh = sheetByName_(ss, 'Цели'); const res = { plan: {}, conv: {}, cum: [] };
  if (!sh) return res;
  const v = sh.getDataRange().getDisplayValues();
  // шапка плана по дням: ячейка «Дата», справа «Сим» — столбец ищем по всей строке (служебные столбцы слева могут сдвигаться)
  let hr = -1, dc = -1;
  for (let r = 0; r < v.length && hr < 0; r++) for (let c = 0; c < v[r].length - 1; c++) if (norm_(v[r][c]) === 'дата' && norm_(v[r][c + 1]) === 'сим') { hr = r; dc = c; break; }
  for (let r = 0; r < Math.min(v.length, hr >= 0 ? hr : 60); r++) {         // плановая конверсия: имя в C, процент в E (только выше таблицы по дням)
    const nm = String(v[r][2] || '').trim(), pc = String(v[r][4] || '').trim();
    if (nm && !dateParts_(nm) && /%$/.test(pc) && norm_(nm) !== 'сим' && num_(pc) !== null) res.conv[norm_(nm)] = num_(pc);
  }
  if (hr >= 0) {
    for (let c = dc + 2; c < v[hr].length; c++) { const nm = String(v[hr][c] || '').trim(); if (nm) res.plan[norm_(nm)] = []; }
    for (let r = hr + 1; r < v.length; r++) {
      const dp = dateParts_(v[r][dc]); if (!dp) break;
      res.cum[dp.d - 1] = num_(v[r][dc + 2]);                                  // накопительная доля плана на день, %
      for (let c = dc + 2; c < v[hr].length; c++) { const nm = String(v[hr][c] || '').trim(); if (nm) res.plan[norm_(nm)][dp.d - 1] = num_(v[r][c]); }
    }
  }
  return res;
}
function readOperator_(ss, month, year) {
  const sh = sheetByName_(ss, 'Детали оператор'), res = { simOp: {}, inactive: {}, active: {}, stale: false, any: false };
  if (!sh) return res;
  const v = sh.getDataRange().getDisplayValues();
  const hdr = [];
  for (let r = 0; r < v.length; r++) { const ci = v[r].findIndex(function (c) { return /^тт\s*kpi/i.test(String(c).trim()); }); if (ci >= 0) hdr.push({ r: r, ci: ci }); }
  hdr.forEach(function (h, k) {
    let kind = '';
    for (let r = h.r; r >= Math.max(0, h.r - 8) && !kind; r--) { const a = norm_(v[r][0]); if (/^не актив/.test(a)) kind = 'inactive'; else if (/^актив/.test(a)) kind = 'active'; else if (/^регистрац/.test(a)) kind = 'simOp'; }
    if (!kind) return;
    const cols = {}; let anyDate = false;
    v[h.r].forEach(function (c, i) { const dp = dateParts_(c); if (!dp) return; anyDate = true; if (dp.m === month && dp.y === year) { cols[i] = dp.d; res.any = true; } });
    if (anyDate && !Object.keys(cols).length) res.stale = true;
    const end = k + 1 < hdr.length ? hdr[k + 1].r - 1 : v.length;
    for (let r = h.r + 1; r < end; r++) {
      const nm = String(v[r][h.ci] || '').trim(); if (!nm || /^всего|^итого/i.test(nm)) continue;
      const key = norm_(nm), arr = res[kind][key] = res[kind][key] || [];
      Object.keys(cols).forEach(function (i) { const n = num_(v[r][i]); if (n !== null) arr[cols[i] - 1] = (arr[cols[i] - 1] || 0) + n; });
    }
  });
  return res;
}
function reconBuild_(ss, u) {
  const names = kpiShops_(ss, readAccess_(ss))[u.key] || [];
  const cal = readSchedule_(ss).cal, salonSheet = readSalonSheet_(ss);
  const dim = cal.dim || salonSheet.daysInMonth || 30, last = cal.today >= 99 ? dim : Math.max(0, (cal.today || 1) - 1);
  const mm = ('0' + (cal.month || 1)).slice(-2);
  const days = []; for (let d = 1; d <= last; d++) days.push(('0' + d).slice(-2) + '.' + mm);
  const traffic = readTrafficByName_(ss), goals = readGoals_(ss), op = readOperator_(ss, cal.month, cal.year);
  const surname = String(u.name || u.key).split(' ')[0], sn = norm_(surname), isTM = u.roleTitle === 'ТМ';
  const pick = function (a) { const o = []; for (let i = 0; i < last; i++) o.push(a && a[i] != null ? a[i] : null); return o; };
  const sumArr = function (list) { const o = []; for (let i = 0; i < last; i++) { let t = 0, any = false; list.forEach(function (a) { if (a && a[i] != null) { t += a[i]; any = true; } }); o.push(any ? t : null); } return o; };
  const simOf = function (n) { let x = salonSheet.salons[n]; if (!x) { const k = norm_(n); Object.keys(salonSheet.salons).some(function (z) { if (norm_(z) === k) { x = salonSheet.salons[z]; return true; } return false; }); } const a = []; if (x && x.sim && x.sim.days) x.sim.days.forEach(function (z) { const dp = dateParts_(z.d); if (dp) a[dp.d - 1] = z.v; }); return a; };
  const one = function (n) {
    const k = norm_(n);
    return { sim1c: pick(simOf(n)), traffic: pick(traffic[k]), plan: pick(goals.plan[k]), planConv: goals.conv[k] != null ? goals.conv[k] : null,
      simOp: pick(op.simOp[k]), inactive: pick(op.inactive[k]), active: pick(op.active[k]) };
  };
  const per = names.map(function (n) { return { key: n, label: n, series: one(n) }; });
  const useSn = isTM && (traffic[sn] || goals.plan[sn]);
  const tot = {
    sim1c: sumArr(per.map(function (o) { return o.series.sim1c; })),
    traffic: useSn && traffic[sn] ? pick(traffic[sn]) : sumArr(per.map(function (o) { return o.series.traffic; })),
    plan: useSn && goals.plan[sn] ? pick(goals.plan[sn]) : sumArr(per.map(function (o) { return o.series.plan; })),
    planConv: null, simOp: sumArr(per.map(function (o) { return o.series.simOp; })), inactive: sumArr(per.map(function (o) { return o.series.inactive; })), active: sumArr(per.map(function (o) { return o.series.active; })),
  };
  if (isTM && goals.conv[sn] != null) tot.planConv = goals.conv[sn];
  else { let w = 0, t = 0; per.forEach(function (o) { const tr = o.series.traffic.reduce(function (a, b) { return a + (b || 0); }, 0); if (o.series.planConv != null && tr > 0) { w += o.series.planConv * tr; t += tr; } }); tot.planConv = t ? w / t : null; }
  const nearOpts = [];                                                         // смежные салоны (тот же адрес, другой бренд) — для РТТ/УТТ
  if (!isTM) { try {
    const have = {}; names.forEach(function (n) { have[norm_(n)] = 1; });
    const ng = nearSalons_(ss, names);
    ((ng && ng.groups) || []).forEach(function (g) { g.near.forEach(function (x) { if (have[norm_(x.name)]) return; have[norm_(x.name)] = 1; nearOpts.push({ key: x.name, label: x.name, near: true, nearTo: g.salon, series: one(x.name) }); }); });
  } catch (e) {} }
  return { month: cal.month, days: days, opStale: !!op.stale && !op.any,
    options: [{ key: 'all', label: 'Всего по ' + surname, series: tot }].concat(per, nearOpts), updated: new Date().toISOString() };
}

/* ---------- «Остальные карточки» из «Деталей продаж салоны»: всё, кроме телефонов/аксессуаров/услуг/страховок/сим ---------- */
// Находит на листе все таблицы «дата × салоны» (и сводные «Значения») правее и ниже основных блоков: защита экрана, ПЗУ, часы,
// наушники, БП, чистка, пакет, трафик и т. п. Результат фильтруется по салонам руководителя (итоги пересчитываются).
function cardsAll_(ss) {
  return cached_('cards:all', TTL_, function () {
    return { cards: cardsModel_(ss) };                          // карточки считаем из чеков (раньше — сводные правее основных блоков)
  });
}
function cardsGet_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (!u.manager) return { error: 'forbidden' };
  const acc = readAccess_(ss), shopsBy = kpiShops_(ss, acc);
  const tmU = u.roleTitle !== 'ТМ' && grantHas_(ss, u, 'cards') ? acc.filter(function (a) { return a.roleTitle === 'ТМ'; })[0] : null;   // гостевой доступ: карточки по всей сети ТМ
  const names = shopsBy[tmU ? tmU.key : u.key] || shopsBy[u.key] || [];
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


/* ---------- чтение листов продавцов ---------- */
function readPlan_(ss) { return cached_('plan', TTL_, function () { return readPlanRaw_(ss); }); }
function readPlanRaw_(ss) {                                       // план, прогноз, место и рейтинг продавца — та же модель, что у ТМ во вкладке «Сотрудники ДЛ» (dlModel_), без листа «До сотрудника ДЛ»
  return dlModel_(ss).rows.map(function (r) {
    const metrics = METRICS.map(function (m) {
      const k = EMP_KEYS_[m.key], pm = r.plan[k] || 0, fc = r.fc[k], f = r.fact[k] || 0;
      return { key: m.key, label: m.label, planDay: Math.max(0, r.day[k] || 0), planMonth: pm, fact: Math.round(f),
        pct: pm > 0 ? f / pm * 100 : 0, forecast: fc === null ? null : pm * fc, forecastPct: fc === null ? null : fc * 100 };
    });
    return { name: r.name, place: r.place, rating: r.score, tt: salonShort_(r.tt),
      hours: { plan: r.planH || 0, fact: r.factH || 0, fact1c: r.wt },
      metrics: metrics,
      quality: { vmr: num_(r.vmr), credits: r.cn, creditsRub: r.cs, attach: r.att, simLeft: r.rem.sim } };
  });
}

function readIncome_(ss) { return cached_('income', TTL_, function () { return readIncomeRaw_(ss); }); }
function readIncomeRaw_(ss) { return incomeModel_(ss); }          // бонусы по чекам считаем на сайте (раньше — сводные «Доход сотрудника»)

/* ---------- утилиты ---------- */
function num_(s) {
  s = String(s === undefined || s === null ? '' : s).replace(/[\s ]/g, '').replace('%', '').replace(',', '.');
  if (s === '' || s.charAt(0) === '#') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function norm_(s) { return String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' ').trim(); }
function short_(full) { return String(full).trim().split(/\s+/).slice(0, 2).join(' '); }

/* =====================================================================================================================
 * РАСЧЁТЫ НА САЙТЕ ВМЕСТО СВОДНЫХ ТАБЛИЦ
 * ---------------------------------------------------------------------------------------------------------------------
 * Раньше сайт читал готовые «сводные» (pivot) на листах «Детали продаж салоны», «Детали продаж сотрудники»,
 * «Доход сотрудника», «Детализация СИМ», «Свод графика» (и формулы листов «общий свод», «Факт продаж»,
 * «До сотрудника ДЛ», которые на эти сводные ссылаются). Все эти сводные строились из ОДНОГО листа «Группа товаров»
 * (выгрузка чеков: A Группа, B Номенклатура, D ТМ, E Торговая точка, F Продавец, H Дата-время, J Количество, K Сумма)
 * и, для графика, из листа «График» (блок CB:CJ). Теперь сайт сам читает эти исходные листы и считает то же самое:
 *
 *   сводная                          строки × столбцы           значение     фильтр (как в сводной)
 *   ────────────────────────────────────────────────────────────────────────────────────────────────────────────────
 *   Телефоны                         дата × салон               Σ Сумма      Группа ∈ {пусто, «Телефоны и планшеты»}; ТМ = «Миллер»
 *   Аксессуары                       дата × салон               Σ Сумма      Группа = «Аксессуары»
 *   Услуги                           дата × салон               Σ Сумма      Группа = «Услуги»; не «ЦУП» (2-е слово номенклатуры)
 *   Сим                              дата × салон               Σ Кол-во     Группа = «Сим-карты»; нет в листе «Сим исключения»
 *   Страховки                        дата × салон               Σ Сумма      Группа = «Страхование и защита»
 *   Защита экрана                    дата × салон               Σ Кол-во     2-е слово номенклатуры = «пленка» / «стекло»
 *   ПЗУ / Часы / Наушники / Чистка   дата × салон               число строк  ТМ = «Миллер»; 2-е слово = «аккумулятор» / «часы»(или 1-е «смарт-часы») /
 *   БП / Пакет                                                                «наушники» / «чистка»; 1-е слово = «Безлимитная» / «Пакет»
 *   «Детали продаж сотрудники»       продавец                   Σ Сумма/Кол-во   те же фильтры, что выше (Телефоны и Аксессуары — только ТМ «Миллер»)
 *   «Доход сотрудника»               продавец × дата / группа   Σ Бонус      строка чека × ставка из «Надстройки» (по категории салона) × 0,9; только «Учитывается» = 1
 *   «Детализация СИМ»                салон / продавец × группа тарифа   Σ Кол-во   Сим-карты, не в «Сим исключения»; группа — из «Разбивка по тарифам»
 *   «Свод графика»                   сотрудник × салон          Σ Итог план по часам   График!CB:CJ: роль ∈ {ПТТ, УТТ, РТТ, новичок}; «Учитывается» = 1
 *
 * План на месяц — листы «Цели» / «План»; доля плана «на сегодня» и по дням — «Цели» (накопительная %) и «План по конверсии» (по дням).
 * Всё складывается в кэш (gzip) так же, как остальные чтения; warmCache() прогревает его по триггеру.
 * ===================================================================================================================== */
const RAW_ = { SHEET: 'Группа товаров', TM: 'Миллер', EXCL: 'Сим исключения', RATES: 'Надстройка', TARIFFS: 'Разбивка по тарифам', GRAFIK: 'График' };
const ROLES_SVOD_ = { 'сп': 1, 'птт': 1, 'утт': 1, 'ртт': 1, 'новичок': 1 };
const SEASON_ = (new Date().getMonth() + 1) <= 6 ? 1.1 : 0.9;   // «коэфф. сезонности»: январь–июнь 1,1 · июль–декабрь 0,9 (по месяцу выгрузки)
const CARD_DEFS_ = [                                    // направления «Деталей продаж салоны» (основные пять + карточки)
  { key: 'phones', label: 'Телефоны',    unit: 'rub', main: true, idx: 4 },
  { key: 'acc',    label: 'Аксессуары',  unit: 'rub', main: true, idx: 7 },
  { key: 'serv',   label: 'Услуги',      unit: 'rub', main: true, idx: 10 },
  { key: 'sim',    label: 'Сим',         unit: 'pcs', main: true, idx: 16 },
  { key: 'ins',    label: 'Страховки',   unit: 'rub', main: true, idx: 13 },
];
const EXTRA_CARDS_ = [                                  // карточки (правее основных блоков на листе)
  { key: 'screen', title: 'Дней без продажи · Защита экрана', dir: 'Аксессуары' },
  { key: 'pzu',    title: 'ПЗУ',      dir: 'Аксессуары' },
  { key: 'watch',  title: 'Часы',     dir: 'Аксессуары' },
  { key: 'head',   title: 'Наушники', dir: 'Аксессуары' },
  { key: 'bp',     title: 'БП',       dir: 'Услуги' },
  { key: 'clean',  title: 'Чистка',   dir: 'Услуги' },
  { key: 'pack',   title: 'Пакет',    dir: 'Услуги' },
];
function words_(s) { return String(s || '').split(' ').filter(function (w) { return w !== ''; }); }
function lcase_(s) { return String(s || '').toLowerCase().replace(/ё/g, 'е'); }
function pad2_(n) { return ('0' + n).slice(-2); }
function dstr_(v, tz) {                                 // дата чека → «дд.мм.гггг» (строка либо Date)
  if (v instanceof Date) return Utilities.formatDate(v, tz || 'Europe/Moscow', 'dd.MM.yyyy');
  const m = /^(\d{1,2})\.(\d{2})\.(\d{4})/.exec(String(v || '').trim());
  return m ? pad2_(m[1]) + '.' + m[2] + '.' + m[3] : '';
}
// Колонки-признаки листа «Группа товаров» (L:AM) — те же формулы, что в таблице, но на сайте
function saleFlags_(r, excl) {
  const w = words_(r.n), L = lcase_(w[0]), M = lcase_(w[1]), tm = r.tm === RAW_.TM;
  const ex = tm ? (excl.hasOwnProperty(lcase_(String(r.n).trim())) ? excl[lcase_(String(r.n).trim())] : 0) : '';
  return {
    L: L, M: M, tm: tm,
    screen: (M === 'пленка' || M === 'стекло') ? 1 : 0,                         // «Фактр пленки»
    pzu: tm && M === 'аккумулятор' ? 1 : 0,                                     // «Факт ПЗУ»
    watch: tm && (L === 'смарт-часы' || M === 'часы') ? 1 : 0,                  // «Фактр часов»
    head: tm && (M === 'наушники' || L === 'наушники') ? 1 : 0,                 // «Фактр наушники»
    serv: tm && r.g === 'Услуги' && M !== 'цуп' ? 1 : 0,                        // «Факт услуги»
    bp: tm && L === 'безлимитная' ? 1 : 0,                                      // «Факт БП»
    pack: tm && L === 'пакет' ? 1 : 0,                                          // «Факт пакета»
    clean: tm && (L === 'чистка' || M === 'чистка') ? 1 : 0,                    // «Фактр чистки»
    excl: ex,                                                                   // «Исключение сим»: 0 — считаем, 1 — исключение, '' — не Миллер
    simOk: (ex === 0 || ex === '' || ex === null),                              // фильтр сводной «Сим»: AM ∈ {0, пусто}
    barakat: w[8] !== undefined && lcase_(w[8]) === '"баракат"' ? 1 : 0,        // «Детализация СИМ» → «Баракат»
  };
}
function hourOf_(v) {                                  // час чека: из Date или из строки «дд.мм.гггг ч:мм:сс»
  if (v instanceof Date) return v.getHours();
  const m = String(v || '').match(/\s(\d{1,2}):/); return m ? Number(m[1]) : -1;
}
// Чтение сырья (кэш): строки чеков + список исключений
function salesRows_(ss) {
  return cached_('sales:rows', TTL_, function () {
    const sh = sheetByName_(ss, RAW_.SHEET); if (!sh) return { rows: [], excl: {}, missing: true };
    const tz = ss.getSpreadsheetTimeZone() || 'Europe/Moscow';
    const lr = sh.getLastRow(); if (lr < 2) return { rows: [], excl: {} };
    const v = sh.getRange(2, 1, lr - 1, 11).getValues(), rows = [];
    v.forEach(function (x) {
      const d = dstr_(x[7], tz), b = String(x[1] || '').trim();
      if (!d || !b) return;
      rows.push({ g: String(x[0] || '').trim(), n: b, tm: String(x[3] || '').trim(), s: String(x[4] || '').trim(), e: String(x[5] || '').trim(),
        d: d, q: Number(x[9]) || 0, k: Number(x[10]) || 0,
        c: String(x[6] || '').trim(), h: hourOf_(x[7]), p: String(x[8] || '').trim() });
    });
    const excl = {}, es = sheetByName_(ss, RAW_.EXCL);
    if (es) es.getRange(1, 1, Math.max(1, es.getLastRow()), 2).getValues().forEach(function (x, i) {
      const nm = lcase_(String(x[0] || '').trim()); if (i === 0 || !nm || excl.hasOwnProperty(nm)) return;    // VLOOKUP берёт первое совпадение
      excl[nm] = (x[1] === '' || x[1] === null) ? '' : Number(x[1]);
    });
    return { rows: rows, excl: excl };
  });
}
// Агрегация: { cards:{key:{салон:{день:значение}}}, emp:{ФИО:{phones,acc,serv,sim,ins}}, month, year, maxDay, salons:{key:[имена]} }
function salesAgg_(data) {
  const rows = data.rows || [], excl = data.excl || {};
  let ym = 0;
  rows.forEach(function (r) { const p = r.d.split('.'); ym = Math.max(ym, Number(p[2]) * 12 + Number(p[1])); });
  const out = { cards: {}, emp: {}, month: ym ? ((ym - 1) % 12) + 1 : 0, year: ym ? Math.floor((ym - 1) / 12) : 0, maxDay: 0, salonsAll: {} };
  const put = function (key, salon, day, val) {
    const c = out.cards[key] = out.cards[key] || {}, s = c[salon] = c[salon] || {};
    s[day] = (s[day] || 0) + val;
  };
  const eadd = function (e, key, val) { const o = out.emp[e] = out.emp[e] || { phones: 0, acc: 0, serv: 0, sim: 0, ins: 0 }; o[key] += val; };
  rows.forEach(function (r) {
    const p = r.d.split('.'); if (Number(p[1]) + Number(p[2]) * 12 !== ym) return;      // только текущий месяц листа
    const day = Number(p[0]); out.maxDay = Math.max(out.maxDay, day); out.salonsAll[r.s] = 1;
    const f = saleFlags_(r, excl), g = r.g;
    if ((g === '' || g === 'Телефоны и планшеты') && f.tm) { put('phones', r.s, day, r.k); eadd(r.e, 'phones', r.k); }
    if (g === 'Аксессуары') { put('acc', r.s, day, r.k); if (f.tm) eadd(r.e, 'acc', r.k); }
    if (g === 'Услуги' && f.serv) { put('serv', r.s, day, r.k); eadd(r.e, 'serv', r.k); }
    if (g === 'Сим-карты' && f.simOk) { put('sim', r.s, day, r.q); eadd(r.e, 'sim', r.q); }
    if (g === 'Страхование и защита') { put('ins', r.s, day, r.k); eadd(r.e, 'ins', r.k); }
    if (f.screen) put('screen', r.s, day, r.q);
    ['pzu', 'watch', 'head', 'bp', 'clean', 'pack'].forEach(function (k) { if (f[k]) put(k, r.s, day, 1); });
  });
  return out;
}
function salesModel_(ss) { return cached_('sales:agg', TTL_, function () { return salesAgg_(salesRows_(ss)); }); }

/* ---------- План: «Цели» (план на месяц по салонам/ТМ + накопительная доля) и «План по конверсии» (доля по дням) ---------- */
const GOAL_DIRS_ = { 'телефоны': 'phones', 'аксессуары': 'acc', 'услуги': 'serv', 'страховки': 'ins', 'сим': 'sim' };
function readGoalsTable_(ss) {                          // { norm(имя): { phones, acc, serv, ins, sim } } — план на месяц
  return cached_('goals:table', TTL_, function () {
    const sh = sheetByName_(ss, 'Цели'), out = {};
    if (!sh) return out;
    const v = sh.getDataRange().getDisplayValues();
    let hr = -1;
    for (let r = 0; r < Math.min(v.length, 12) && hr < 0; r++) for (let c = 0; c < v[r].length - 1; c++) if (GOAL_DIRS_[norm_(v[r][c])] && norm_(v[r][c + 1]) === 'план на день') { hr = r; break; }
    if (hr < 0) return out;
    const cols = {}; v[hr].forEach(function (h, c) { const k = GOAL_DIRS_[norm_(h)]; if (k && norm_(v[hr][c + 1]) === 'план на день') cols[k] = c; });
    for (let r = hr + 1; r < v.length; r++) {
      const nm = String(v[r][2] || '').trim(); if (!nm) break;
      const o = out[norm_(nm)] = { name: nm };
      Object.keys(cols).forEach(function (k) { const n = num_(v[r][cols[k]]), d = num_(v[r][cols[k] + 1]); if (n !== null) o[k] = n; if (d !== null) o[k + 'Day'] = d; });
    }
    return out;
  });
}
function readDayWeights_(ss) {                          // { день: доля } — «План по конверсии»: «Дата» («1.11») и «%» рядом
  return cached_('goals:dayw', TTL_, function () {
    const sh = sheetByName_(ss, 'План по конверсии'), out = {};
    if (!sh) return out;
    const v = sh.getDataRange().getDisplayValues();
    let hr = -1, hc = -1;
    for (let r = 0; r < Math.min(v.length, 10) && hr < 0; r++) for (let c = 0; c < v[r].length - 1; c++) if (norm_(v[r][c]) === 'дата' && String(v[r][c + 1]).trim() === '%') { hr = r; hc = c; break; }
    if (hr < 0) return out;
    for (let r = hr + 1; r < v.length; r++) {
      const m = /^(\d{1,2})\.(\d{1,2})/.exec(String(v[r][hc]).trim()); if (!m) break;
      const w = num_(v[r][hc + 1]); if (w !== null) out[Number(m[1])] = w / 100;
    }
    return out;
  });
}
// контекст плана: сегодня (с учётом конца месяца), доля плана «к сегодняшнему дню», план ТМ
function planCtx_(ss, model) {
  const cal = readSchedule_(ss).cal, goals = readGoals_(ss), dim = cal.dim || new Date(model.year || 2000, model.month || 1, 0).getDate() || 30;
  const day = cal.today >= 99 ? dim : Math.max(1, Math.min(dim, cal.today || model.maxDay || 1));
  const cumArr = goals.cum || [], cum = cumArr[day - 1] != null ? cumArr[day - 1] / 100 : (cumArr.length ? 0 : day / dim);
  return { cal: cal, dim: dim, day: day, cum: cum, table: readGoalsTable_(ss), dayW: readDayWeights_(ss) };
}
function pctTxt_(x) { return x === null || x === undefined || !isFinite(x) ? '' : Math.round(x * 100) + '%'; }
function numTxt_(x) { return x === null || x === undefined || !isFinite(x) ? '' : String(Math.round(x)); }
function sortedKeys_(o) { return Object.keys(o || {}).sort(); }

// «Детали продаж салоны» — основные блоки в том же виде, что на листе (значения сводных считаются из чеков; план, отставание, прогноз — по плану «Цели»)
function detailsGrid_(ss) {
  return cached_('sales:details', TTL_, function () {
    const m = salesModel_(ss), pc = planCtx_(ss, m), tm = pc.table[norm_(RAW_.TM)] || {}, grid = [];
    CARD_DEFS_.forEach(function (def) {
      const data = m.cards[def.key]; if (!data) return;
      const salons = sortedKeys_(data), n = salons.length, W = 3 + n + 4;
      const row = function () { const r = []; for (let i = 0; i < W; i++) r.push(''); return r; };
      const should = function (name) { const p = (pc.table[norm_(name)] || {})[def.key]; return p ? (def.key === 'sim' ? Math.ceil(p * pc.cum - 1e-9) : Math.round(p * pc.cum)) : null; };
      const tot = function (s) { let t = 0; Object.keys(data[s] || {}).forEach(function (d) { t += data[s][d]; }); return t; };
      const gap = row(), fc = row(), hint = row(), head = row();
      gap[1] = String(def.idx); gap[2] = def.label; fc[1] = def.label; fc[2] = 'Прогноз'; hint[1] = 'План на день'; hint[2] = def.label; hint[3] = 'Торговая точка';
      head[2] = 'Дата'; head[3 + n] = 'Итого'; head[3 + n + 2] = 'Закрытие';
      let allT = 0;
      salons.forEach(function (s, i) {
        const t = tot(s), sh = should(s), plan = (pc.table[norm_(s)] || {})[def.key];
        allT += t; head[3 + i] = s; fc[2] = 'Прогноз';
        gap[3 + i] = sh === null ? '' : numTxt_(sh - t);
        fc[3 + i] = plan ? pctTxt_(t / (plan * pc.cum)) : '';
      });
      const tmPlan = tm[def.key], tmShould = should(RAW_.TM);
      gap[3 + n] = tmShould === null ? '' : numTxt_(tmShould - allT);
      fc[3 + n] = tmPlan ? pctTxt_(allT / (tmPlan * pc.cum)) : '';
      grid.push(gap, fc, hint, head);
      for (let d = 1; d <= m.maxDay; d++) {
        const r = row(), dd = pad2_(d) + '.' + pad2_(m.month) + '.' + m.year;
        const pd = tmPlan && pc.dayW[d] ? tmPlan * pc.dayW[d] : null;
        let dt = 0, any = false;
        salons.forEach(function (s, i) { const x = data[s] && data[s][d]; if (x !== undefined) { r[3 + i] = numTxt_(x); dt += x; any = true; } });
        r[1] = numTxt_(pd); r[2] = dd; r[3 + n] = any ? numTxt_(dt) : ''; r[3 + n + 2] = pd ? pctTxt_(dt / pd) : '';
        grid.push(r);
      }
      const t = row(); t[0] = def.label; t[1] = '#N/A'; t[2] = 'Итого';
      salons.forEach(function (s, i) { t[3 + i] = numTxt_(tot(s)); }); t[3 + n] = numTxt_(allT);
      grid.push(t);
      for (let i = 0; i < 3; i++) { const f = row(); f[1] = '#N/A'; grid.push(f); }
      grid.push(row());
    });
    return grid;
  });
}
// Карточки правее основных блоков («Защита экрана», ПЗУ, часы, наушники, БП, чистка, пакет)
function cardsModel_(ss) {
  const m = salesModel_(ss), out = [];
  EXTRA_CARDS_.forEach(function (c) {
    const data = m.cards[c.key]; if (!data) return;
    const salons = sortedKeys_(data), days = [], v = [], plan = [];
    for (let d = 1; d <= m.maxDay; d++) {
      days.push(pad2_(d) + '.' + pad2_(m.month) + '.' + m.year); plan.push(null);
      v.push(salons.map(function (s) { const x = data[s] && data[s][d]; return x === undefined ? null : x; }));
    }
    const tot = salons.map(function (s, j) { let t = 0; v.forEach(function (r) { t += r[j] || 0; }); return t; });
    out.push({ title: c.title, dir: c.dir, kind: 'days', pct: false, salons: salons, days: days, plan: plan, v: v, tot: tot, sums: [] });
  });
  return out;
}

/* ---------- «Свод графика»: часы сотрудника по салонам — из График!CB:CJ ---------- */
function scheduleBlock_(ss) {                          // строки блока «Сотрудник · Салон · … · Учитывается · Итог план по часам»
  return cached_('sched:block', TTL_, function () {
    const sh = sheetByName_(ss, RAW_.GRAFIK), out = [];
    if (!sh) return out;
    const lr = sh.getLastRow(), lc = Math.min(sh.getLastColumn(), 110);
    if (lr < 1 || lc < 60) return out;
    const g = sh.getRange(1, 60, lr, lc - 59).getDisplayValues();
    let hr = -1, c0 = -1;
    for (let r = 0; r < Math.min(g.length, 6) && hr < 0; r++) for (let c = 1; c < g[r].length - 1; c++) if (norm_(g[r][c]) === 'сотрудник' && norm_(g[r][c + 1]) === 'салон') { hr = r; c0 = c; break; }
    if (hr < 0) return out;
    let ci = -1, hi = -1;
    g[hr].forEach(function (h, c) { if (norm_(h) === 'учитывается') ci = c; if (norm_(h) === 'итог план по часам') hi = c; });
    if (ci < 0 || hi < 0) return out;
    for (let r = hr + 1; r < g.length; r++) {
      const name = String(g[r][c0] || '').trim(), salon = String(g[r][c0 + 1] || '').trim();
      if (!name && !salon) continue;
      out.push({ role: String(g[r][c0 - 1] || '').trim(), name: name, salon: salon, use: num_(g[r][ci]), hours: num_(g[r][hi]) });
    }
    return out;
  });
}
function scheduleHours_(ss) {                          // { norm(ФИО): { салон: часы } }
  const out = {};
  scheduleBlock_(ss).forEach(function (x) {
    if (!ROLES_SVOD_[norm_(x.role)] || x.use !== 1 || !x.name || !x.salon || x.hours === null) return;
    const o = out[norm_(x.name)] = out[norm_(x.name)] || {};
    o[x.salon] = (o[x.salon] || 0) + x.hours;
  });
  return out;
}

/* ---------- Ставки и справочники для бонуса ---------- */
function ratesBook_(ss) {
  return cached_('sales:rates', TTL_, function () {
    const out = { rates: {}, grade: {}, tariffs: {}, fil: {} };
    const rs = sheetByName_(ss, RAW_.RATES);
    if (rs) {
      const v = rs.getDataRange().getDisplayValues(), cats = [v[0][1], v[0][2], v[0][3]].map(function (x) { return String(x).trim(); });
      const pr = function (s) { s = String(s || '').trim(); if (!s) return null; const n = num_(s); return n === null ? null : (/%/.test(s) ? n / 100 : n); };
      for (let r = 2; r < Math.min(v.length, 40); r++) {
        const k = lcase_(String(v[r][0]).trim()); if (!k || out.rates[k]) continue;
        const o = {}; cats.forEach(function (c, i) { o[c] = pr(v[r][1 + i]); }); out.rates[k] = o;
      }
      for (let r = 1; r < v.length; r++) { const s = norm_(v[r][6]); if (s && !out.grade[s]) out.grade[s] = String(v[r][8]).trim(); }
    }
    const ts = sheetByName_(ss, RAW_.TARIFFS);
    if (ts) ts.getDataRange().getDisplayValues().forEach(function (x, i) {
      const k = lcase_(String(x[0]).trim()); if (i === 0 || !k || out.tariffs[k]) return;
      out.tariffs[k] = { grp: String(x[1] || '').trim(), abon: String(x[2] || '').trim() };
    });
    scheduleBlock_(ss).forEach(function (x) { const k = norm_(x.name); if (k && out.fil[k] === undefined) out.fil[k] = x.use; });   // VLOOKUP — первое совпадение
    return out;
  });
}
const GOLD_ = 'Золото / Комбо сим 3 мес (доп. к оплате за ТП Sim)';
// Бонус по одной строке чека (колонки AP…AX листа «Группа товаров»)
// Категория салона для ставок: точное название → тот же салон (sameSalon_) → тот же город (первое слово названия: КЗН, НЧ, ПРМ…) — берётся самая частая категория города
function cityTok_(name) { const t = norm_(name).split(' ')[0] || ''; return t; }
function catFor_(T, salon) {
  const k = norm_(salon); if (!k) return undefined;
  if (T.grade[k]) return T.grade[k];
  T._cc = T._cc || {};
  if (T._cc[k] !== undefined) return T._cc[k] || undefined;
  let res = '';
  const keys = Object.keys(T.grade);
  for (let i = 0; i < keys.length && !res; i++) if (sameSalon_(keys[i], salon)) res = T.grade[keys[i]];
  if (!res) {
    const ct = cityTok_(salon), cnt = {};
    keys.forEach(function (x) { if (cityTok_(x) === ct && T.grade[x]) cnt[T.grade[x]] = (cnt[T.grade[x]] || 0) + 1; });
    Object.keys(cnt).forEach(function (c) { if (!res || cnt[c] > cnt[res]) res = c; });
  }
  T._cc[k] = res; return res || undefined;
}
function rowBonus_(r, f, T) {
  const cat = catFor_(T, r.s), rt = function (lbl) { const o = T.rates[lcase_(lbl)]; const x = o && o[cat]; return x === null || x === undefined ? 0 : x; };
  const g = r.g; let b = 0;
  if (f.serv === 1) b += r.k * rt('Услуги | Настройки') * SEASON_;
  if (g === 'Страхование и защита') b += r.k * rt('Страхование остальное / КСЖ /') * SEASON_;
  if (g === 'Сим-карты' && f.excl === 0) { const t = T.tariffs[lcase_(r.n.trim())]; if (t && t.grp) b += rt(t.grp) * SEASON_; }
  if (g === 'Телефоны и планшеты') b += (f.L === 'мобильный' || f.L === 'сотовый') ? r.k * 0.1 : r.k * 0.01 * SEASON_;
  if (g === 'Аксессуары') b += r.k * rt('Аксессуары и носимая электроника 10%') * SEASON_;
  if (g === 'Финансовые продукты') b += r.k * 0.003;
  return b;
}
// «Доход сотрудника»: { byName: { norm(ФИО): { days:[{d,v}], groups:[{name,v}], gtotal } }, month }
function incomeModel_(ss) {
  return cached_('sales:income', TTL_, function () {
    const data = salesRows_(ss), m = salesModel_(ss), T = ratesBook_(ss), M = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
    const byE = {}, groups = {};
    data.rows.forEach(function (r) {
      const p = r.d.split('.'); if (Number(p[1]) + Number(p[2]) * 12 !== m.month + m.year * 12) return;
      if (T.fil[norm_(r.e)] !== 1) return;                                        // «Фильтр сотрудника» = 1
      const b = rowBonus_(r, saleFlags_(r, data.excl), T), e = byE[norm_(r.e)] = byE[norm_(r.e)] || { days: {}, grp: {} };
      const gn = /^\(не исп\)/i.test(r.g) ? 'Прочее' : (r.g || 'Прочее');
      e.days[Number(p[0])] = (e.days[Number(p[0])] || 0) + b; e.grp[gn] = (e.grp[gn] || 0) + b; groups[gn] = 1;
    });
    const gl = Object.keys(groups).sort(function (a, b) { return a < b ? -1 : a > b ? 1 : 0; }), byName = {};
    Object.keys(byE).forEach(function (k) {
      const e = byE[k], days = [];
      for (let d = 1; d <= m.maxDay; d++) days.push({ d: pad2_(d) + '.' + pad2_(m.month) + '.' + m.year, v: e.days[d] === undefined ? null : e.days[d] });
      let tot = 0; gl.forEach(function (g) { tot += e.grp[g] || 0; });
      byName[k] = { days: days, groups: gl.map(function (g) { return { name: g, v: e.grp[g] || 0 }; }), gtotal: tot };
    });
    return { byName: byName, month: m.month ? M[m.month - 1] + ' ' + m.year : '' };
  });
}

/* ---------- «Общий свод» и «Факт продаж»: те же столбцы, что на листах, но считаем на сайте ---------- */
const SUMMARY_DIRS_ = [['sim', 'Сим'], ['phones', 'Телефоны'], ['acc', 'Аксессуары'], ['serv', 'Услуги'], ['ins', 'Страховки']];
function summaryRows_(ss) {
  return cached_('sales:summary', TTL_, function () {
    const m = salesModel_(ss), pc = planCtx_(ss, m), names = Object.keys(pc.table);
    let days = 4;                                       // «Планирование на N дн.» — ячейка A3 листа «общий свод» (вводит человек)
    try { const sh = sheetByName_(ss, CONFIG.SHEET_SUMMARY); if (sh) { const x = num_(sh.getRange(3, 1).getDisplayValue()); if (x > 0) days = x; } } catch (e) {}
    const tot = function (key, s) { let t = 0; const d = (m.cards[key] || {})[s] || {}; Object.keys(d).forEach(function (k) { t += d[k]; }); return t; };
    const tmName = norm_(RAW_.TM);
    const rows = names.map(function (k) {
      const g = pc.table[k], isTM = k === tmName, o = { name: g.name, dirs: {} };
      SUMMARY_DIRS_.forEach(function (d) {
        const key = d[0], planM = g[key], perDay = g[key + 'Day'];
        let fact = 0; if (isTM) Object.keys(m.cards[key] || {}).forEach(function (s) { fact += tot(key, s); }); else fact = tot(key, g.name);
        const should = planM ? (key === 'sim' ? Math.ceil(planM * pc.cum - 1e-9) : Math.round(planM * pc.cum)) : null;
        const gap = should === null ? null : should - fact;
        o.dirs[key] = { planDay: perDay === undefined ? null : perDay + (gap !== null && gap > 0 ? gap / days : 0), fact: fact, gap: gap, fc: planM ? fact / (planM * pc.cum) : null, planMonth: planM || null };
      });
      return o;
    });
    return { rows: rows, days: days, today: pc.day };
  });
}
function summaryGrid_(ss) {
  const S = summaryRows_(ss), grid = [];
  const r0 = ['Общий свод'], r1 = ['Планирование на ', 'Общая'], r2 = [String(S.days)], r3 = [''];
  SUMMARY_DIRS_.forEach(function (d) { r2.push(d[1], '', '', ''); r3.push('План на день', 'Факт', 'Отставание', 'Прогноз'); });
  grid.push(r0, r1, r2, r3);
  S.rows.forEach(function (o) {
    const r = [o.name];
    SUMMARY_DIRS_.forEach(function (d) { const x = o.dirs[d[0]]; r.push(numTxt_(x.planDay), numTxt_(x.fact), numTxt_(x.gap), pctTxt_(x.fc)); });
    grid.push(r);
  });
  return grid;
}
function factGrid_(ss) {
  const S = summaryRows_(ss), m = salesModel_(ss), tz = ss.getSpreadsheetTimeZone() || 'Europe/Moscow', now = new Date();
  const today = Utilities.formatDate(now, tz, 'dd.MM.yyyy'), td = Number(today.slice(0, 2));
  const T = ratesBook_(ss), promo = 'Промо Sim от оператора / Акция «Бустер»', data = salesRows_(ss);
  const todayOf = function (key, salon) { const d = (m.cards[key] || {})[salon]; return d && d[td] !== undefined && today.slice(3) === pad2_(m.month) + '.' + m.year ? d[td] : null; };
  const promoToday = {};
  data.rows.forEach(function (r) { if (r.d !== today || r.g !== 'Сим-карты') return; const f = saleFlags_(r, data.excl), t = T.tariffs[lcase_(r.n.trim())]; if (f.excl === 0 && t && t.grp === promo) promoToday[r.s] = (promoToday[r.s] || 0) + r.q; });
  const grid = [[today], ['Минммальный прогноз', 'Актуально на', '', today + ' ' + Utilities.formatDate(now, tz, 'HH:mm:ss')]];
  const h2 = ['Планирование на '], h3 = [String(S.days)];
  SUMMARY_DIRS_.forEach(function (d) { h2.push(d[1], '', ''); h3.push('План', 'Факт', '%'); });
  h2.push('Пленки', 'Часы', 'ПЗУ', 'Чистки', 'Промо сим', 'Пакет услуг', 'БП'); grid.push(h2, h3);
  const sum = {};
  S.rows.forEach(function (o, i) {
    const r = [o.name], isTM = i === 0;
    SUMMARY_DIRS_.forEach(function (d) {
      const x = o.dirs[d[0]]; let f = 0;
      if (isTM) Object.keys(m.cards[d[0]] || {}).forEach(function (s) { f += todayOf(d[0], s) || 0; }); else f = todayOf(d[0], o.name) || 0;
      r.push(numTxt_(x.planDay), numTxt_(f), x.planDay ? pctTxt_(f / x.planDay) : '');
    });
    ['screen', 'watch', 'pzu', 'clean'].forEach(function (k) { let f = 0; if (isTM) Object.keys(m.cards[k] || {}).forEach(function (s) { f += todayOf(k, s) || 0; }); else f = todayOf(k, o.name) || 0; r.push(numTxt_(f)); });
    let pr = 0; if (isTM) Object.keys(promoToday).forEach(function (s) { pr += promoToday[s]; }); else pr = promoToday[o.name] || 0; r.push(numTxt_(pr));
    ['pack', 'bp'].forEach(function (k) { let f = 0; if (isTM) Object.keys(m.cards[k] || {}).forEach(function (s) { f += todayOf(k, s) || 0; }); else f = todayOf(k, o.name) || 0; r.push(numTxt_(f)); });
    grid.push(r);
  });
  return grid;
}

/* ---------- «Детализация СИМ»: салоны и продавцы × группа тарифов ---------- */
function cmpRu_(a, b) { const la = /^[A-Za-z]/.test(a), lb = /^[A-Za-z]/.test(b); if (la !== lb) return la ? 1 : -1; return a < b ? -1 : a > b ? 1 : 0; }
function simGrid_(ss) {
  return cached_('sales:simgrid', TTL_, function () {
    const data = salesRows_(ss), m = salesModel_(ss), T = ratesBook_(ss);
    const bySalon = {}, bySeller = {}, byGrpDay = {}, groups = {}, barakat = {};
    data.rows.forEach(function (r) {
      const p = r.d.split('.'); if (Number(p[1]) + Number(p[2]) * 12 !== m.month + m.year * 12) return;
      const f = saleFlags_(r, data.excl);
      if (r.g !== 'Сим-карты') return;
      const t = T.tariffs[lcase_(r.n.trim())], grp = t && t.grp ? t.grp : '#N/A';
      if (f.excl === 0) {
        groups[grp] = 1;
        (bySalon[r.s] = bySalon[r.s] || {})[grp] = ((bySalon[r.s] || {})[grp] || 0) + r.q;
        (bySeller[r.e] = bySeller[r.e] || {})[grp] = ((bySeller[r.e] || {})[grp] || 0) + r.q;
      }
      if (f.excl === 0 || f.excl === 1) { groups[grp] = 1; (byGrpDay[grp] = byGrpDay[grp] || {})[Number(p[0])] = (byGrpDay[grp][Number(p[0])] || 0) + r.q; }
      if (f.barakat) (barakat[r.s] = barakat[r.s] || {})[Number(p[0])] = ((barakat[r.s] || {})[Number(p[0])] || 0) + r.q;
    });
    const gl = Object.keys(groups).sort(cmpRu_), grid = [];
    const table = function (title, head, map, pct) {
      if (title) grid.push([title, 'Группа тарифов']);
      grid.push([head].concat(gl, ['Итого']));
      const keys = Object.keys(map).sort(), colT = {}; let all = 0;
      keys.forEach(function (k) {
        let t = 0; gl.forEach(function (g) { t += map[k][g] || 0; });
        const row = [k]; gl.forEach(function (g) { const x = map[k][g] || 0; colT[g] = (colT[g] || 0) + x; row.push(pct ? (t ? Math.round(x / t * 100) + '%' : '') : (x ? String(x) : '')); });
        row.push(pct ? '' : String(t)); all += t; grid.push(row);
      });
      const tr = ['Итого']; gl.forEach(function (g) { tr.push(pct ? (all ? Math.round((colT[g] || 0) / all * 100) + '%' : '') : String(colT[g] || 0)); }); tr.push(pct ? '' : String(all)); grid.push(tr);
    };
    table('SUM из Количество', 'Торговая точка', bySalon, false);
    for (let i = 0; i < 6; i++) grid.push([]);
    table('', 'Торговая точка', bySalon, true);
    for (let i = 0; i < 6; i++) grid.push([]);
    table('SUM из Количество', 'Продавец', bySeller, false);
    for (let i = 0; i < 3; i++) grid.push([]);
    // группа тарифов × дата (с учётом исключений, как в сводной)
    const dh = ['Группа тарифов']; for (let d = 1; d <= m.maxDay; d++) dh.push(pad2_(d) + '.' + pad2_(m.month) + '.' + m.year); dh.push('Итого');
    grid.push(['SUM из Количество', 'Дата']); grid.push(dh);
    const dt = {}; let aa = 0;
    gl.forEach(function (g) { const row = [g]; let t = 0; for (let d = 1; d <= m.maxDay; d++) { const x = (byGrpDay[g] || {})[d] || 0; dt[d] = (dt[d] || 0) + x; t += x; row.push(x ? String(x) : ''); } aa += t; row.push(String(t)); grid.push(row); });
    const tr = ['Итого']; for (let d = 1; d <= m.maxDay; d++) tr.push(String(dt[d] || 0)); tr.push(String(aa)); grid.push(tr);
    const bs = Object.keys(barakat).sort();
    if (bs.length) {
      grid.push([]); grid.push(['SUM из Количество', '«БАРАКАТ»']); grid.push(['Торговая точка'].concat(dh.slice(1)));
      bs.forEach(function (s) { const row = [s]; let t = 0; for (let d = 1; d <= m.maxDay; d++) { const x = barakat[s][d] || 0; t += x; row.push(x ? String(x) : ''); } row.push(String(t)); grid.push(row); });
    }
    return grid;
  });
}

/* ---------- факт продавцов («Детали продаж сотрудники» → чеки) ---------- */
const EMP_KEYS_ = { sim: 'sim', smart: 'phones', acc: 'acc', serv: 'serv', ins: 'ins' };
function empFacts_(ss) { const o = {}, e = salesModel_(ss).emp || {}; Object.keys(e).forEach(function (k) { o[norm_(k)] = e[k]; }); return o; }
/* ---------- «До сотрудника ДЛ»: план на месяц / на день, отставание, прогноз и место — считаем на сайте, механика листа ----------
   План сотрудника на месяц по направлению = Σ по салонам ( план салона на месяц [лист «План»] ÷ часы всех сотрудников салона × часы сотрудника в салоне ).
   Часы — «График» CB:CJ (учитывается = 1, «Итог план по часам»). ТТ — салон, где у сотрудника больше всего часов.
   План на день (отставание) = (план на месяц − факт) ÷ (плановых смен − отработанных смен); смены — «График» (AO / BZ).
   Прогноз = факт ÷ отработанные смены × плановые смены ÷ план на месяц.  Рейтинг = Σ мест по прогнозу (5 направлений, лучший = 1) − аттачмент − кредиты.  Место — по рейтингу. */
const DL_DIRS_ = [['sim', 'SIM'], ['phones', 'Телефоны'], ['acc', 'Аксессуары'], ['serv', 'Услуги'], ['ins', 'Страховки']];
function salonWord_(s) {                                // «КЗН Ершова 62 БЛН ТЦ Арт Центр» → «ершова»
  const w = norm_(s).replace(/\(.*?\)/g, ' ').split(/\s+/).filter(function (x) { return x && x !== 'кзн' && x !== 'блн' && !/^[\d\/а-я]?\d[\d\/а-я]*$/.test(x) && !/^\d/.test(x); });
  return w[0] || norm_(s);
}
function salonShort_(s) { return String(s || '').replace(/^\s*КЗН\s+/i, '').replace(/\s+БЛН.*$/i, '').replace(/\s*\(.*?\)\s*/g, ' ').trim(); }
function dlMonthPlanFor_(ss, salon) {                   // план салона на месяц; название в «График» и в «План» может отличаться оформлением
  const mp = readMonthPlan_(ss), k = norm_(salon);
  if (mp[k]) return mp[k];
  const w = salonWord_(salon), ks = Object.keys(mp);
  for (let i = 0; i < ks.length; i++) if (salonWord_(ks[i]) === w) return mp[ks[i]];
  return null;
}
function dlRankAvg_(arr, i, desc) {                     // RANK.AVG: ничьи — среднее место
  let gt = 0, eq = 0; const x = arr[i];
  for (let j = 0; j < arr.length; j++) { if (arr[j] === null) continue; if (arr[j] === x) eq++; else if (desc ? arr[j] > x : arr[j] < x) gt++; }
  return gt + (eq + 1) / 2;
}
function dlModel_(ss) {
  return cached_('dl:model', TTL_, function () {
    const sh = sheetByName_(ss, RAW_.GRAFIK); if (!sh) return { rows: [] };
    const lr = Math.min(sh.getLastRow(), 400), g = sh.getRange(1, 1, lr, Math.min(sh.getLastColumn(), 88)).getDisplayValues();
    const cE = colIdx_('E'), cAO = colIdx_('AO'), cAS = colIdx_('AS'), cBZ = colIdx_('BZ');
    const cRole = colIdx_('CB'), cName = colIdx_('CC'), cSal = colIdx_('CD'), cPlanH = colIdx_('CE'), cUse = colIdx_('CI'), cCJ = colIdx_('CJ');
    const sh_ = {};                                                      // смены и факт часов по ФИО (строки «График» — по одной на сотрудника и салон)
    for (let r = 5; r < g.length; r++) {
      const k = norm_(g[r][cE]); if (!k) continue;
      const o = sh_[k] = sh_[k] || { plan: 0, work: 0, fact: 0 };
      o.plan += num_(g[r][cAO]) || 0; o.work += num_(g[r][cBZ]) || 0; o.fact += num_(g[r][cAS]) || 0;
    }
    const emps = {}, order = [], salonH = {};
    for (let r = 5; r < g.length; r++) {                                 // блок CB:CJ — часы по салонам
      const role = norm_(g[r][cRole]), name = String(g[r][cName] || '').trim(), salon = String(g[r][cSal] || '').trim();
      if (!name || !salon || !ROLES_SVOD_[role] || num_(g[r][cUse]) !== 1) continue;
      const h = num_(g[r][cCJ]), hp = num_(g[r][cPlanH]); if (h === null || h <= 0) continue;
      const k = norm_(name), e = emps[k] = emps[k] || (order.push(k), { name: name, role: g[r][cRole], hours: {}, planH: 0 });
      e.hours[salon] = (e.hours[salon] || 0) + h; e.planH += hp || 0;
      salonH[salon] = (salonH[salon] || 0) + h;
    }
    const num0 = function (s) { const n = num_(s); return n === null ? 0 : n; };
    const wt = {}, w = sheetByName_(ss, 'Рабочее время сотрудников');
    if (w) w.getDataRange().getDisplayValues().forEach(function (x) { const k = norm_(x[1]); if (k && wt[k] === undefined && num_(x[4]) !== null) wt[k] = num_(x[4]); });
    const cr = {}, cs = sheetByName_(ss, 'Кредиты текущий месяц');
    if (cs) cs.getDataRange().getDisplayValues().forEach(function (x) { const k = norm_(x[2]); if (!k) return; const o = cr[k] = cr[k] || { n: 0, sum: 0 }; o.n++; o.sum += num0(x[3]); });
    const vmr = {}, vs = sheetByName_(ss, 'ВМР'), mon = new Date().getMonth() + 1;
    if (vs) vs.getRange(4, 2, Math.max(1, Math.min(vs.getLastRow(), 400) - 3), 14).getDisplayValues().forEach(function (x) { const k = norm_(x[0]); if (k && vmr[k] === undefined) vmr[k] = (mon >= 7 && mon <= 12 ? String(x[mon - 6] || '').trim() : ''); });   // колонки месяцев ВМР — июль…декабрь; иначе пусто, а не чужая колонка
    const att = {}, as = sheetByName_(ss, 'Показатели сотрудников');
    if (as) {
      const hv = as.getRange(3, 8, 1, 80).getDisplayValues()[0]; let ac = -1; hv.forEach(function (h, i) { if (ac < 0 && norm_(h) === 'аттачмент') ac = i; });
      if (ac >= 0) as.getRange(4, 8, Math.max(1, Math.min(as.getLastRow(), 1000) - 3), ac + 1).getDisplayValues().forEach(function (x) { const k = norm_(x[0]); if (k && att[k] === undefined) att[k] = num_(x[ac]); });
    }
    const ef = empFacts_(ss), rows = order.map(function (k) {
      const e = emps[k], t = sh_[k] || { plan: 0, work: 0, fact: 0 }, f = ef[k] || {};
      let tt = '', best = -1; Object.keys(e.hours).forEach(function (s) { if (e.hours[s] > best) { best = e.hours[s]; tt = s; } });
      const plan = {}, fact = {}, day = {}, rem = {}, fc = {};
      const den = t.plan - t.work;
      DL_DIRS_.forEach(function (d) {
        let p = 0, any = false;
        Object.keys(e.hours).forEach(function (s) { const sp = dlMonthPlanFor_(ss, s); if (sp && sp[d[0]] > 0 && salonH[s] > 0) { p += sp[d[0]] * e.hours[s] / salonH[s]; any = true; } });
        plan[d[0]] = any ? p : null; fact[d[0]] = num0(f[d[0]]);
        rem[d[0]] = plan[d[0]] === null ? null : plan[d[0]] - fact[d[0]];
        day[d[0]] = rem[d[0]] !== null && den > 0 ? Math.max(0, rem[d[0]]) / den : null;
        fc[d[0]] = plan[d[0]] > 0 && t.work > 0 && t.plan > 0 ? fact[d[0]] / t.work * t.plan / plan[d[0]] : null;
      });
      return { k: k, name: e.name, role: e.role, rated: !/^(ртт|утт|тм)$/.test(norm_(e.role)), tt: tt, planH: e.planH, wt: wt[k] === undefined ? null : wt[k], factH: t.fact, plan: plan, fact: fact, rem: rem, day: day, fc: fc,
        vmr: vmr[k] || '', cn: (cr[k] || {}).n || 0, cs: (cr[k] || {}).sum || 0, att: att[k] === undefined || att[k] === null ? 0 : att[k] };
    });
    DL_DIRS_.forEach(function (d) {
      const arr = rows.map(function (r) { return r.rated ? r.fc[d[0]] : null; });
      rows.forEach(function (r, i) { r['rk_' + d[0]] = arr[i] === null ? null : dlRankAvg_(arr, i, true); });
    });
    rows.forEach(function (r) {
      let s = 0, ok = true; DL_DIRS_.forEach(function (d) { if (r['rk_' + d[0]] === null) ok = false; else s += r['rk_' + d[0]]; });
      r.score = ok && r.rated ? s - r.att - r.cn : null;
    });
    const sc = rows.map(function (r) { return r.score; });
    rows.forEach(function (r, i) { r.place = r.score === null ? null : dlRankAvg_(sc, i, false); });
    rows.sort(function (a, b) { return (a.place === null ? 1e9 : a.place) - (b.place === null ? 1e9 : b.place) || (a.name < b.name ? -1 : 1); });
    return { rows: rows };
  });
}
function dlGrid_(ss) {                                  // та же раскладка, что у листа «До сотрудника ДЛ» (ФИО — колонка H)
  const m = dlModel_(ss).rows, n0 = function (x, d) { return x === null || x === undefined || !isFinite(x) ? '' : (d ? (Math.round(x * 10) / 10).toString().replace('.', ',') : String(Math.round(x))); };
  const pc = function (x) { return x === null || x === undefined || !isFinite(x) ? '' : Math.round(x * 100) + '%'; };
  const lab = DL_DIRS_.map(function (d) { return d[1]; });
  const h0 = new Array(34).fill(''), h1 = new Array(34).fill('');
  ['', 'ТТ', 'План, ч', 'Факт 1С, ч', 'Факт график, ч', 'Разница, ч', 'Место', 'ФИО', 'Остаток по SIM'].forEach(function (t, i) { h1[i] = t; });
  lab.forEach(function (t, i) { h1[9 + i] = t; h1[14 + i] = t; h1[19 + i] = t; h1[28 + i] = t; });
  h0[8] = 'План на день'; h0[14] = 'План на месяц'; h0[19] = 'Факт'; h0[24] = 'Качество'; h0[28] = 'Прогноз'; h0[33] = 'Итог';
  h1[24] = 'ВМР'; h1[25] = 'Кредитов, шт'; h1[26] = 'Кредитов, руб'; h1[27] = 'Аттачмент'; h1[33] = 'Рейтинг';
  const out = [h0, h1];
  m.forEach(function (r) {
    const row = new Array(34).fill('');
    row[1] = salonShort_(r.tt); row[2] = n0(r.planH); row[3] = r.wt === null ? '' : n0(r.wt); row[4] = n0(r.factH); row[5] = r.wt === null ? '' : n0(r.factH - r.wt); row[6] = r.place === null ? '' : n0(r.place, 1); row[7] = r.name;
    row[8] = n0(r.rem.sim);
    DL_DIRS_.forEach(function (d, i) { row[9 + i] = n0(r.day[d[0]]); row[14 + i] = n0(r.plan[d[0]]); row[19 + i] = n0(r.fact[d[0]]); row[28 + i] = pc(r.fc[d[0]]); });
    row[24] = r.vmr; row[25] = n0(r.cn); row[26] = n0(r.cs); row[27] = n0(r.att); row[33] = r.score === null ? '' : n0(r.score, 1);
    out.push(row);
  });
  return out;
}


/* ---------- автовыгрузка «Товарные группы» из DataLens → лист «Группа товаров» (колонки A:K) ----------
   Раньше отчёт скачивали руками и вставляли в лист. Теперь страницу DataLens (под вашим входом) открывает задача на компьютере,
   забирает строки отчёта (период: с начала месяца по сегодня, ТМ «Миллер») и отправляет сюда действием import_checks.
   Защита: общий секрет — свойство скрипта IMPORT_SECRET (Настройки проекта → Свойства скрипта); без него запросы отклоняются.
   Пишутся ТОЛЬКО колонки A:K (Группа, Номенклатура, Характеристика, ТМ, Торговая точка, Продавец, Номер чека, Дата время чека,
   Тип оплаты, Количество, Сумма); формулы правее (L:…) не трогаются. Пустую выгрузку и выгрузку, сильно меньшую текущей, не принимаем. */
function importChecks_(req) {
  const secret = PropertiesService.getScriptProperties().getProperty('IMPORT_SECRET');
  if (!secret || String(req.secret || '') !== secret) return { error: 'forbidden' };
  const rows = req.rows, offset = Number(req.offset) || 0, total = Number(req.total) || (Array.isArray(rows) ? rows.length : 0);
  if (!Array.isArray(rows) || !rows.length || rows.length > 2000 || offset < 0 || total < 1 || total > 60000 || offset + rows.length > total) return { error: 'bad_request', message: 'rows' };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);       // выгрузка идёт порциями (по ~300 строк): offset — номер первой строки порции, total — всего строк
  try {
    const ss = ss_(), sh = ss.getSheetByName(RAW_.SHEET), tz = ss.getSpreadsheetTimeZone();
    if (!sh) return { error: 'server', message: 'no sheet' };
    const out = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!Array.isArray(r) || r.length < 11) return { error: 'bad_request', message: 'row ' + (offset + i) };
      let d = r[7];
      if (typeof d === 'string' && /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(d)) d = Utilities.parseDate(d, tz, 'yyyy-MM-dd HH:mm:ss');
      out.push([String(r[0] == null ? '' : r[0]), String(r[1] == null ? '' : r[1]), String(r[2] == null ? '' : r[2]), String(r[3] == null ? '' : r[3]),
        String(r[4] == null ? '' : r[4]), String(r[5] == null ? '' : r[5]), String(r[6] == null ? '' : r[6]), d,
        String(r[8] == null ? '' : r[8]), Number(r[9]) || 0, Number(r[10]) || 0]);
    }
    let was = 0;
    if (offset === 0) {                                                // первая порция: проверка и очистка колонок A:K
      const last = sh.getLastRow();
      if (last > 1) { const a = sh.getRange(2, 1, last - 1, 1).getValues(); for (let i = 0; i < a.length; i++) if (a[i][0] !== '') was++; }
      if (!req.force && was > 200 && total < was * 0.5) return { error: 'too_small', have: was, got: total };   // защита от сбойной выгрузки
      if (sh.getMaxRows() < total + 1) sh.insertRowsAfter(sh.getMaxRows(), total + 1 - sh.getMaxRows());
      if (last > 1) sh.getRange(2, 1, Math.max(last - 1, total), 11).clearContent();
    }
    sh.getRange(2 + offset, 1, out.length, 11).setValues(out);
    sh.getRange(2 + offset, 8, out.length, 1).setNumberFormat('dd.MM.yyyy HH:mm:ss');
    if (offset + out.length >= total) {
      PropertiesService.getScriptProperties().setProperty('import_checks_last', JSON.stringify({ at: new Date().toISOString(), rows: total, from: String(req.from || ''), to: String(req.to || '') }));
      try { scheduleWarm_(); } catch (e) {}                       // обновить кэш сайта сразу после выгрузки
    }
    return { ok: true, offset: offset, rows: out.length, total: total, was: was };
  } finally { lock.releaseLock(); }
}

/* ---------- «Сотрудники ДЛ» → Бонус: часы, бонус по группам и прогноз ЗП ----------
   Часы: план — «План, ч» (блок CB:CJ листа «График»), отработано — «Факт 1С, ч» (лист «Рабочее время сотрудников»; если нет — «Факт график, ч»).
   Бонус по группам — тот же расчёт по чекам, что у сотрудника на вкладке «Доход» (incomeModel_).
   Прогноз бонуса = бонус на сегодня ÷ отработанные часы × плановые часы. Оклад = факт часов × ставка города («Надстройка», G:H).
   Защита бонуса: бонус ÷ факт часов < ставки грейда сотрудника по салону («Надстройка», K:M по грейду из O:R) → бонус доводится до ставки × часы. */
// «Надстройка»: ставка города и ставки грейдов по салону (G:M), грейд сотрудника (O:R)
function salaryBook_(ss) {
  return cached_('sal:book', TTL_, function () {
    const out = { salons: [], grade: {} }, rs = sheetByName_(ss, RAW_.RATES); if (!rs) return out;
    const v = rs.getDataRange().getDisplayValues();
    for (let r = 1; r < v.length; r++) {
      const sal = String(v[r][6] || '').trim(), city = num_(v[r][7]);
      if (sal && city !== null) out.salons.push({ name: sal, key: norm_(sal), city: city, rates: { A: num_(v[r][10]), B: num_(v[r][11]), C: num_(v[r][12]) } });
      const nm = norm_(v[r][14]), g = String(v[r][17] || '').trim().toUpperCase();
      if (nm && /^[ABC]$/.test(g) && !out.grade[nm]) out.grade[nm] = g;
    }
    return out;
  });
}
function salonRate_(book, salon) {
  const k = norm_(salon); if (!k) return null;
  for (let i = 0; i < book.salons.length; i++) if (book.salons[i].key === k) return book.salons[i];
  for (let i = 0; i < book.salons.length; i++) if (sameSalon_(book.salons[i].name, salon)) return book.salons[i];
  const ct = cityTok_(salon);                                              // салона нет в «Надстройке» — ставка такая же, как у салонов того же города
  for (let i = 0; i < book.salons.length; i++) if (cityTok_(book.salons[i].name) === ct) return book.salons[i];
  return null;
}
function dlBonusGet_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  const team = u.roleTitle !== 'ТМ';                                        // РТТ/УТТ (в т.ч. с выданным доступом) видят только сотрудников своих салонов; вся сеть — только у ТМ
  if (team ? !u.manager : !repAllowed_(ss, u, 'dl')) return { error: 'forbidden' };
  const full = cached_('dl:bonus', TTL_, function () {
    const rows = dlModel_(ss).rows, inc = incomeModel_(ss), gset = {}, out = [], wkOf = {}, weeks = [], book = salaryBook_(ss);
    const MON = ['янв','фев','мар','апр','мая','июн','июл','авг','сен','окт','ноя','дек'];
    rows.forEach(function (r) {
      if (!r.rated) return;                                                   // РТТ / УТТ / ТМ в бонус сотрудников не входят
      const b = inc.byName[r.k]; if (!b) return;
      b.days.forEach(function (x) {                                           // недели месяца: пн–вс (первая и последняя могут быть неполными)
        if (wkOf[x.d] !== undefined) return;
        const p = x.d.split('.'), dt = new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0])), dow = (dt.getDay() + 6) % 7, key = Number(p[0]) - dow;
        let w = weeks.filter(function (q) { return q.key === key; })[0];
        if (!w) { w = { key: key, from: Number(p[0]), to: Number(p[0]), mon: MON[Number(p[1]) - 1] }; weeks.push(w); }
        w.from = Math.min(w.from, Number(p[0])); w.to = Math.max(w.to, Number(p[0]));
        wkOf[x.d] = weeks.indexOf(w);
      });
      const worked = r.wt !== null && r.wt !== undefined ? r.wt : r.factH, g = {};
      b.groups.forEach(function (x) { if (x.v) { g[x.name] = x.v; gset[x.name] = (gset[x.name] || 0) + x.v; } });
      const fc = worked > 0 && r.planH > 0 ? b.gtotal / worked * r.planH : null;
      const wk = weeks.map(function () { return 0; });
      b.days.forEach(function (x) { if (x.v) wk[wkOf[x.d]] += x.v; });
      // оклад = факт часов × ставка города; защита бонуса: бонус на час < ставки грейда (по салону) → бонус доводится до ставки × часы
      const sal = salonRate_(book, r.tt), gr = book.grade[r.k] || 'C', rate = sal && gr ? sal.rates[gr] : null, city = sal ? sal.city : null;
      const wH = worked || 0, pH = r.planH || 0, perH = wH > 0 ? b.gtotal / wH : null;
      const prot = rate !== null && rate !== undefined && perH !== null && perH < rate;
      const protNow = prot ? rate * wH - b.gtotal : 0, protFc = prot && pH > 0 ? rate * pH - fc : 0;
      const salary = city === null ? null : city * wH, salaryFc = city === null ? null : city * pH;
      const fcProt = fc === null ? null : fc + protFc;
      out.push({ name: r.name, role: r.role, tt: salonShort_(r.tt), ttFull: r.tt, planH: pH, workH: wH, factGraph: r.factH || 0, groups: g, total: b.gtotal, fc: fc, wk: wk,
        grade: gr, city: city, rate: rate === undefined ? null : rate, salary: salary, salaryFc: salaryFc, perH: perH, prot: prot, protNow: protNow, protFc: protFc,
        totalProt: b.gtotal + protNow, fcProt: fcProt, zp: fcProt === null || salaryFc === null ? null : salaryFc + fcProt, zpNow: salary === null ? null : salary + b.gtotal + protNow });
    });
    out.forEach(function (o) { while (o.wk.length < weeks.length) o.wk.push(0); });
    weeks.sort(function (a, b) { return a.key - b.key; });
    const groups = Object.keys(gset).sort(function (a, b) { return gset[b] - gset[a]; });
    out.sort(function (a, b) { return (b.fc === null ? -1 : b.fc) - (a.fc === null ? -1 : a.fc) || b.total - a.total; });
    return { groups: groups, rows: out, weeks: weeks.map(function (w) { return { label: (w.from === w.to ? w.from : w.from + '–' + w.to) + ' ' + w.mon }; }), month: inc.month, at: new Date().toISOString() };
  });
  if (!team) return full;
  const shops = kpiShops_(ss, readAccess_(ss))[u.key] || [];
  const mine = full.rows.filter(function (r) { return r.ttFull && shops.some(function (sh) { return sameSalon_(r.ttFull, sh); }); });
  return { groups: full.groups, rows: mine, weeks: full.weeks, month: full.month, at: full.at, team: true };
}

/* ---------- журнал входов (виден только ТМ) ---------- */
const AUTHLOG_SHEET = 'Журнал входов';
function authLog_(user, event, dev) {
  try {
    const ss = ss_();
    let sh = ss.getSheetByName(AUTHLOG_SHEET);
    if (!sh) { sh = ss.insertSheet(AUTHLOG_SHEET); sh.appendRow(['Время', 'Сотрудник', 'Роль', 'Событие', 'Устройство']); sh.setFrozenRows(1); }
    sh.appendRow([new Date(), user.name, user.roleTitle, event, safeCell_(String(dev || '').slice(0, 60))]);
  } catch (e) {}
}
function authLogGet_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (u.roleTitle !== 'ТМ') return { error: 'forbidden' };
  const sh = ss.getSheetByName(AUTHLOG_SHEET), rows = [];
  if (sh && sh.getLastRow() > 1) {
    const n = Math.min(4000, sh.getLastRow() - 1), v = sh.getRange(sh.getLastRow() - n + 1, 1, n, 5).getValues();
    v.forEach(function (r) { if (r[0] instanceof Date) rows.push([Utilities.formatDate(r[0], 'Europe/Moscow', "yyyy-MM-dd'T'HH:mm:ss'+03:00'"), String(r[1]), String(r[2]), String(r[3]), String(r[4])]); });
  }
  const props = PropertiesService.getScriptProperties();
  const users = readAccess_(ss).map(function (a) { return { name: a.name, role: a.roleTitle, pin: !!props.getProperty('pin:' + a.key) }; });
  return { rows: rows, users: users, now: new Date().toISOString() };
}

/* ---------- вложения новостей (хранятся кусками base64 на листе «Вложения новостей») ---------- */
const FILES_SHEET = 'Вложения новостей', FILE_CHUNK = 40000, FILE_MAX = 3, FILE_B64_MAX = 1400000, FILE_B64_TOTAL = 2800000;
const FILE_MIME = /^(image\/(jpeg|png|webp|gif)|application\/(pdf|msword|vnd\.openxmlformats-officedocument\.[a-z.]+|vnd\.ms-(excel|powerpoint)|zip)|text\/(plain|csv))$/;
function filesSheet_(ss) {
  let sh = ss.getSheetByName(FILES_SHEET);
  if (!sh) { sh = ss.insertSheet(FILES_SHEET); sh.getRange(1, 1, 1, 4).setValues([['ID файла', 'ID новости', 'Часть', 'Данные (base64)']]); sh.setFrozenRows(1); }
  return sh;
}
function saveFiles_(ss, newsId, files) {
  if (!files || !files.length) return [];
  if (!Array.isArray(files) || files.length > FILE_MAX) return null;
  let total = 0; const meta = [], rows = [];
  for (let i = 0; i < files.length; i++) {
    const f = files[i] || {}, data = String(f.data || ''), mime = String(f.mime || '');
    if (!FILE_MIME.test(mime) || !/^[A-Za-z0-9+\/=]+$/.test(data) || data.length > FILE_B64_MAX) return null;
    total += data.length; if (total > FILE_B64_TOTAL) return null;
    const fid = Utilities.getUuid().slice(0, 10), name = String(f.name || 'файл').replace(/[\\/:*?"<>|\r\n]/g, '_').slice(0, 80);
    for (let p = 0, n = 0; p < data.length; p += FILE_CHUNK, n++) rows.push([fid, newsId, n, data.substr(p, FILE_CHUNK)]);
    meta.push({ id: fid, name: name, mime: mime, size: Math.round(data.length * 3 / 4) });
  }
  const sh = filesSheet_(ss), r0 = sh.getLastRow() + 1;
  sh.getRange(r0, 1, rows.length, 4).setNumberFormat('@').setValues(rows);
  return meta;
}
function dropFiles_(ss, newsId) {
  const sh = ss.getSheetByName(FILES_SHEET); if (!sh || sh.getLastRow() < 2) return;
  const v = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === newsId) sh.deleteRow(i + 2);
}
function newsFileGet_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  const fid = String(req.id || ''), sh = ss.getSheetByName(FILES_SHEET);
  if (!sh || sh.getLastRow() < 2) return { error: 'not_found' };
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues(), parts = [];
  let nid = '';
  v.forEach(function (r) { if (String(r[0]) === fid) { nid = String(r[1]); parts.push([Number(r[2]), String(r[3])]); } });
  if (!parts.length) return { error: 'not_found' };
  const nv = newsSheet_(ss).getDataRange().getDisplayValues(); let ok = false, meta = null;
  for (let i = 1; i < nv.length; i++) if (String(nv[i][0]) === nid) {
    ok = norm_(nv[i][6]) === ST_PUB || u.roleTitle === 'ТМ' || norm_(nv[i][2]) === u.key;
    try { meta = JSON.parse(nv[i][8] || '[]').filter(function (m) { return m.id === fid; })[0]; } catch (e) {}
  }
  if (!ok || !meta) return { error: 'forbidden' };
  parts.sort(function (a, b) { return a[0] - b[0]; });
  return { name: meta.name, mime: meta.mime, data: parts.map(function (p) { return p[1]; }).join('') };
}


/* ---------- Attachment: проникновение допродаж в чеки со смартфоном ---------- */
const ATT_ = { PHONE_MIN: 4000, ITEM_MIN: 100, GROUPS: { 'Аксессуары': 1, 'Услуги': 1, 'Страхование и защита': 1 }, MAXCHECKS: 5000 };
const ATT_CATS_ = [
  { k: 'film', t: 'Плёнка / стекло' }, { k: 'case', t: 'Чехол' }, { k: 'chg', t: 'Зарядка / кабель' },
  { k: 'audio', t: 'Наушники / колонки' }, { k: 'acc', t: 'Прочие аксессуары' }, { k: 'serv', t: 'Услуги' }, { k: 'ins', t: 'Страховки' }
];
function attCat_(r) {
  if (r.g === 'Услуги') return 'serv';
  if (r.g === 'Страхование и защита') return 'ins';
  const n = lcase_(r.n);
  if (/плен|стекл/.test(n)) return 'film';
  if (/чехол|бампер|накладк/.test(n)) return 'case';
  if (/кабел|сзу|зарядн|адаптер|аккумул|powerbank|power bank/.test(n)) return 'chg';
  if (/наушник|колонк|гарнитур|airpods/.test(n)) return 'audio';
  return 'acc';
}
function attIsPhone_(r) {
  if (!(r.g === '' || r.g === 'Телефоны и планшеты') || r.q <= 0) return false;
  const n = lcase_(r.n);
  if (/кнопоч|планшет/.test(n)) return false;
  return (r.k / r.q) >= ATT_.PHONE_MIN;
}
// Разбор по чекам: rows уже отфильтрованы по роли
function attachModel_(rows, mgrOf) {
  const ck = {};
  rows.forEach(function (r) {
    if (!r.c) return;
    const key = r.c + '|' + r.s, o = ck[key] = ck[key] || { m: mgrOf ? mgrOf(r.s) : '', c: r.c, s: r.s, e: r.e, d: r.d, h: r.h, p: r.p, ph: null, it: [] };
    if (attIsPhone_(r)) { if (!o.ph) o.ph = { n: r.n, k: r.k, q: r.q, mx: r.k }; else { if (r.k > o.ph.mx) { o.ph.n = r.n; o.ph.mx = r.k; } o.ph.k += r.k; o.ph.q += r.q; } return; }
    if (ATT_.GROUPS[r.g] && r.q > 0 && (r.k / r.q) >= ATT_.ITEM_MIN) o.it.push({ c: attCat_(r), n: r.n, k: r.k, q: r.q });
  });
  const checks = []; Object.keys(ck).forEach(function (k) { if (ck[k].ph) checks.push(ck[k]); });
  const mk = function () { const z = { n: 0, phn: 0, pcs: 0, dsum: 0, sum: 0, none: 0, cats: {} }; ATT_CATS_.forEach(function (c) { z.cats[c.k] = { ch: 0, pcs: 0 }; }); return z; };
  const add = function (z, ch) {
    z.n++; z.phn += ch.ph.q; let pcs = 0; z.sum += ch.ph.k;
    const seen = {};
    ch.it.forEach(function (i) { pcs += i.q; z.dsum += i.k; const c = z.cats[i.c]; c.pcs += i.q; if (!seen[i.c]) { seen[i.c] = 1; c.ch++; } });
    z.pcs += pcs; if (!pcs) z.none++;
  };
  const tot = mk(), bySalon = {}, bySeller = {}, byDay = {}, byHour = {}, byPay = {}, byBand = {};
  const bandOf = function (v) { return v < 10000 ? '4–10 тыс.' : v < 20000 ? '10–20 тыс.' : v < 40000 ? '20–40 тыс.' : '40 тыс.+'; };
  const grp = function (m, k, ch) { add(m[k] = m[k] || mk(), ch); };
  checks.forEach(function (ch) {
    add(tot, ch); grp(bySalon, ch.s, ch); grp(bySeller, ch.e + '|' + ch.s, ch); grp(byDay, ch.d, ch);
    grp(byHour, String(ch.h), ch); grp(byPay, ch.p || 'нет', ch); grp(byBand, bandOf(ch.ph.k), ch);
  });
  const flat = function (m, nameKey) { return Object.keys(m).map(function (k) { const z = m[k]; z[nameKey || 'k'] = k; return z; }); };
  const list = checks.slice().sort(function (a, b) {
    const f = function (x) { const p = x.d.split('.'); return Number(p[2]) * 10000 + Number(p[1]) * 100 + Number(p[0]); };
    return f(b) - f(a) || b.h - a.h;
  }).slice(0, ATT_.MAXCHECKS).map(function (ch) {
    return { c: ch.c, d: ch.d, h: ch.h, s: ch.s, e: ch.e, p: ch.p, m: ch.m, ph: { n: ch.ph.n, k: ch.ph.k, q: ch.ph.q }, it: ch.it };
  });
  return { cfg: { phoneMin: ATT_.PHONE_MIN, itemMin: ATT_.ITEM_MIN }, cats: ATT_CATS_, tot: tot, salon: flat(bySalon), seller: flat(bySeller), day: flat(byDay), hour: flat(byHour), pay: flat(byPay), band: flat(byBand), checks: list };
}
function attachReport_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  const data = salesRows_(ss); if (data.missing) return { missing: true };
  let rows = data.rows || [], scope = 'all', mgrOf = null;
  if (u.roleTitle !== 'ТМ') {
    if (u.manager) {
      const shops = kpiShops_(ss, readAccess_(ss))[u.key] || [];
      rows = rows.filter(function (r) { return shops.some(function (sh) { return sameSalon_(r.s, sh); }); }); scope = 'tt';
    } else { rows = rows.filter(function (r) { return norm_(r.e) === u.key; }); scope = 'me'; }
  }
  if (scope === 'all') {                                   // ТМ: привязка салонов к руководителям (РТТ/УТТ)
    const acc = readAccess_(ss), ks = kpiShops_(ss, acc), mg = acc.filter(function (a) { return a.manager && a.roleTitle !== 'ТМ'; }), memo = {};
    mgrOf = function (salon) {
      if (memo.hasOwnProperty(salon)) return memo[salon];
      let nm = '';
      for (let i = 0; i < mg.length && !nm; i++) if ((ks[mg[i].key] || []).some(function (sh) { return sameSalon_(salon, sh); })) nm = mg[i].name;
      return (memo[salon] = nm);
    };
  }
  const out = attachModel_(rows, mgrOf); out.scope = scope; return out;
}

/* ---------- ВМР для РТТ/УТТ (по их ТТ) и ПТТ/СП (только их проверки) ---------- */
function vmrMine_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  if (u.roleTitle === 'ТМ') return { error: 'forbidden' };                       // у ТМ — полный отчёт на вкладке «Аналитика»
  const rep = tmReportRaw_(ss, 'vmr');
  if (!rep || !rep.grid) return rep || { missing: true };
  const g = rep.grid; let hi = -1;
  for (let r = 0; r < Math.min(g.length, 15); r++) {
    const h = g[r].map(norm_);
    if (h.indexOf('тт') >= 0 && h.indexOf('оценка') >= 0) { hi = r; break; }
  }
  if (hi < 0) return { grid: [] };
  const H = g[hi].map(norm_), iTT = H.indexOf('тт'), iFio = H.indexOf('фио сотрудника');
  let pred;
  if (u.manager) {
    const shops = kpiShops_(ss, readAccess_(ss))[u.key] || [];
    pred = function (r) { return shops.some(function (sh) { return sameSalon_(r[iTT], sh); }); };
  } else {
    pred = function (r) { return iFio >= 0 && norm_(r[iFio]) === u.key; };
  }
  return { name: rep.name, scope: u.manager ? 'tt' : 'me', grid: g.slice(0, hi + 1).concat(g.slice(hi + 1).filter(pred)) };
}

/* =================== Уведомления и новости об обновлениях =================== */
const NOTIF_SHEET = 'Уведомления';
// ВАЖНО: при каждом обновлении сайта добавляйте сюда запись (новая версия `v`) — она сама станет новостью и уведомлением для всех.
const RELEASES = [
  { v: '2026-12-11', title: 'Бонус по салонам, «Сопоставления», доступы',
    text: 'Что нового:\n• Все таблицы приведены к единому стилю (общий класс и общий фон шапки, первого столбца, зебры и итога во всех оформлениях).\n• Исправления интерфейса на телефоне: кнопки не меньше 44px, текст не мельче 11px, шапка не перекрывается кнопками (на узких экранах скрыты название и бейдж «Демо»).\n• Новое оформление «Калькулятор» (как в «Бонус-калькуляторе»): фиолетовый акцент, лаймовые подсветки, шрифты Unbounded и Manrope, точечный фон. Включается в «Оформлении» → «Профиль интерфейса» → «Калькулятор» (светлая и тёмная тема).\n• Единая методика прогноза: темп по завершённым дням × дней в месяце (у сотрудников ДЛ — по сменам); сегодняшний неполный день не берётся.\n• Отчёты объединены: «План на день» и «Факт продаж» теперь внутри «Итогов дня»; «Диагностика» и тепловая карта — внутри «Прогноза»; у РТТ/УТТ вместо «Бонуса команды» — «Сотрудники ДЛ» (рейтинг и бонус только их сотрудников).\n• «Сопоставления»: алерт «Трафик есть — продаж нет».\n• Бонус: коэффициент сезонности по месяцу (янв–июнь 1,1; июль–дек 0,9); салону без категории ставка берётся по названию и городу (как у салонов того же города).\n• Проверка отчётов: сопоставление салонов строже (целые слова, Т2/ТД различаются), «Точка» корректно считает прогноз при графике прошлого месяца, ВМР в рейтинге ДЛ не берёт чужую колонку вне июля–декабря.\n• «Бонус»: если у сотрудника не найден грейд — ему присваивается грейд C (ставка защиты и расчёт по C).\n• «Бонус сотрудников» и «Бонус команды»: сотрудники сгруппированы по салонам, у каждого салона — подытог.\n• «Сотрудники ДЛ» (рейтинг и бонус) у РТТ/УТТ — только сотрудники их салонов, в том числе при выданном доступе; вся сеть — у ТМ.\n• «Сопоставления»: исправлена подгрузка «Сим 1с» — график «Сим 1с / Трафик / Конверсия» снова показывает данные.\n• «Сопоставления» у РТТ/УТТ: помимо своих салонов можно переключиться на смежные (тот же адрес, другой бренд) — трафик, сим 1с и конверсия.\n• «Сотрудники ДЛ» переработаны: вместо широкой таблицы — карточки сотрудников (рейтинг, часы, полосы выполнения плана по направлениям с прогнозом, качество), группировка по салонам, сортировка и поиск; таблица осталась во вкладке «Таблица». Бонус — карточки с раскрывающимся расчётом (оклад, защита, группы, недели).\n• ВМР у РТТ/УТТ — только проверки по их салонам, в том числе при выданном доступе.\n• Временный доступ: ТМ отмечает галочками разделы (Отчёты, Сотрудники ДЛ, План на день, Итоги дня, Карточки, Аналитика, ВМР); «Итоги дня», «План на день», «Карточки» и «Аналитика» открываются в виде ТМ — по всей сети; «Сотрудники ДЛ» — всегда только свои сотрудники.' },
  { v: '2026-12-10', title: 'Бонус сотрудников: оклад и защита бонуса',
    text: 'Что нового:\n• В «Бонус сотрудников» и «Бонус команды» добавлены оклад (факт часов × ставка города) и защита бонуса: если бонус на час ниже ставки грейда сотрудника по салону, бонус доводится до ставки × часы.\n• Прогноз ЗП теперь = оклад + бонус с защитой. Ставки берутся с листа «Надстройка».' },
  { v: '2026-12-09', title: 'Бонус команды для РТТ и УТТ',
    text: 'Что нового:\n• У РТТ и УТТ появилась вкладка «Бонус команды»: часы, бонус по группам, прогноз ЗП и динамика по неделям — только по сотрудникам их салонов.' },
  { v: '2026-12-08', title: 'Бонус сотрудников: читабельнее и по неделям',
    text: 'Что нового:\n• В «Сотрудники ДЛ → Бонус» убраны РТТ и УТТ.\n• Таблица разбита на блоки: часы и бонус (с ₽ за час), бонус по группам, динамика по неделям.\n• Внизу — динамика бонуса по неделям (пн–вс) с изменением к прошлой неделе.' },
  { v: '2026-12-07', title: 'Сотрудники ДЛ: бонус',
    text: 'Что нового:\n• В разделе «Сотрудники ДЛ» появилась вкладка «Бонус»: плановые и отработанные часы, бонус по группам товаров и прогноз ЗП на месяц по каждому сотруднику.' },
  { v: '2026-12-06', title: 'Временный доступ к отчётам',
    text: 'Что нового:\n• ТМ может открыть РТТ/УТТ отчёты (в т. ч. «Сотрудники ДЛ») на заданный срок или до выбранной даты включительно (кроме Бонуса ТМ, Входов, Ссылок): вкладка «Доступы».\n• По истечении срока доступ закрывается сам.' },
  { v: '2026-12-05', title: 'Бонус ТМ: новая мотивация',
    text: 'Что нового:\n• Бонус ТМ пересчитан по листу «Калькулятор»: КВП до 110%, коэффициент SIM без ×1,1, шаг 75–89% = ×0,66.\n• Эконом: вместо Sell Out — «Выполнение плана по заявкам на кредит» (10%); коэффициент кредитов без исключения по заявкам.\n• Админ: «Фотоотчёт» вместо «Конверсии SIM», «Цель от РДО» 90/80%, текучесть от 10% — 0.\n• Операторы: МТС — Скидка на связь + Карты и Voicetech «Комбо 2»; Мегафон — выручка 35%; Билайн — HB++ 25% (активные 50%, выручка 30%).\n• Годовой бонус в «Итого к выплате» не входит.' },
  { v: '2026-12-04', title: 'Презентация в стиле Apple и новые слайды',
    text: 'Что нового:\n• Презентация Weekly Pulse в стиле Apple: светлые карточки, крупная типографика, тёмные титульные слайды и плавные анимации появления с переходами Morph.\n• Инвентаризация: таблица «Да / Нет» по последним 4 периодам.\n• L4L и M2M: отдельные слайды, в рублях и в процентах, с графиком процентов.\n• Услуги: сколько дней из прошедших были с услугами и ₽ услуг в день по салонам.\n• Боковое меню на ПК закреплено на месте — перетаскивать панель и кнопки больше нельзя.\n• Блок «Салоны рядом» на сайте стал понятнее: карточки с планом, фактом, прогнозом и конверсией.' },
  { v: '2026-12-03', title: 'Weekly Pulse: все салоны в презентации',
    text: 'Что нового:\n• На слайдах про салоны теперь все салоны ТМ, без обрезки списков.\n• Если салонов много, таблицы уплотняются и делятся на две колонки, подписи на графиках уменьшаются.\n• Названия салонов сопоставляются с «Экономическими KPI» устойчивее: меньше пустых ячеек.' },
  { v: '2026-12-02', title: 'Новый экран входа',
    text: 'Что нового:\n• Экран входа полностью переработан: тёмная «аврора» с неоном, стеклянная карточка, градиентная кнопка, стикеры и подсказки.\n• Логика входа и PIN не менялась.' },
  { v: '2026-12-01', title: 'Поиск по отчётам и выбор отчётов для презентации',
    text: 'Что нового:\n• В шапке появилась лупа (или клавиша «/», Ctrl+K): быстрый поиск разделов и отчётов с учётом вашего доступа.\n• Weekly Pulse (ТМ): на странице презентации можно выбрать отчёты. По умолчанию выбраны все, новые отчёты сайта появляются в списке сами. Основные 33 слайда не меняются, выбранные дополнительные отчёты добавляются в конец.' },
  { v: '2026-11-30', title: 'Weekly Pulse: презентация в формате «Собрания с управленцами»',
    text: 'Что нового:\n• Кнопка «Скачать» собирает презентацию из 33 слайдов в дизайне сайта, с анимациями и переходами Morph.\n• Данные на момент формирования: графики, инвентаризации, L4L, M2M, трафик и конверсия из KPI ОСС.\n• Везде используется полный логотип SUPER PROдажник by Miller.' },
  { v: '2026-11-29', title: 'Weekly Pulse: презентация недели для ТМ',
    text: 'Что нового:\n• Вкладка «Weekly Pulse» (только ТМ): одной кнопкой собирается презентация .pptx — итоги месяца, неделя к неделе, светофор салонов, лидеры и аутсайдеры, потери и что делать, Attachment, ВМР, фокус на следующую неделю.\n• Можно выбрать слайды и светлую или тёмную тему.' },
  { v: '2026-11-28', title: 'Attachment: допы на один телефон и разбивка по привязке',
    text: 'Что нового:\n• Главный показатель Attachment — все допы (аксессуары, услуги, страховки от 100 ₽) из чеков со смартфоном ÷ все проданные телефоны (от 4 000 ₽): шт и ₽ на один телефон.\n• Новый блок «По привязке»: у ТМ — руководители → салоны → сотрудники, у РТТ/УТТ — их салоны → сотрудники, у остальных — только свои чеки.\n• Убрана карточка «Когда провал: время».' },
  { v: '2026-11-27', title: 'Новое меню на телефоне',
    text: 'Что нового:\n• Внизу — 4 главных раздела с иконками и кнопка «Ещё».\n• «Ещё» открывает шторку со всеми разделами по группам: Итоги и темп, Диагностика, Команда и бонус, Качество, Контроль. Закрывается свайпом вниз, нажатием на фон или Esc.\n• Меню подстраивается под профиль: iOS 27 — плавающая стеклянная панель, iOS 16 — классическая панель вкладок, Pixel — навигация Material 3 с индикатором-пилюлей.' },
  { v: '2026-11-26', title: 'ВМР — отдельная страница, фильтры в Attachment',
    text: 'Что нового:\n• ВМР у всех теперь на отдельной странице. У ТМ полный отчёт ВМР вынесен из «Аналитики» в свою вкладку «ВМР».\n• В «Attachment» — выбор периода выпадающим списком и выбор торговой точки.' },
  { v: '2026-11-25', title: 'Новая вкладка Attachment: допродажи в чеках со смартфоном',
    text: 'Что нового:\n• Вкладка «Attachment» для всех ролей: считаем каждый чек со смартфоном (от 4 000 ₽, без кнопочных) и смотрим, что в него добавили — плёнка/стекло, чехол, зарядка/кабель, наушники, услуги, страховки (позиции от 100 ₽, в штуках).\n• Проникновение по группам, доля чеков без допродаж, где проваливаемся: по салонам, сотрудникам, дням, часам, цене телефона и типу оплаты.\n• Список чеков с разбором: что продали и чего не хватило.' },
  { v: '2026-11-24', title: 'Диагностика, группы отчётов и новая индикация',
    text: 'Что нового:\n• Новая вкладка «Диагностика»: вердикт по плану, светофор салонов, где теряем деньги, достижимость плана, что делать, дни недели, трафик и конверсия.\n• Меню разбито на группы: Итоги и темп · Диагностика · Команда и бонус · Качество · Контроль.\n• «По дням» и «План на день» стали наглядными: тепловая подложка, дисциплина по дням, шкалы нужного темпа.\n• Копирование отчёта в картинку: теперь только сама карточка со скруглением, без рамки 16:9.' },
  { v: '2026-11-23', title: 'Профиль Pixel и новые анимации iOS 16',
    text: 'Что нового:\n• В «Оформлении» — профиль «Pixel» в стиле Google Material You: тональные поверхности, скруглённые карточки, пилюля-индикатор вкладок, переключатели Material 3, круговая волна (ripple) на кнопках и фирменные плавные анимации.\n• Профиль «iOS 16» получил новые анимации: скользящий переключатель вкладок, листы снизу на телефоне, всплывающие панели, уведомления-баннеры, каскадное появление карточек.' },
  { v: '2026-11-22', title: 'Салоны рядом: смежные салоны, трафик и профиль iOS 16',
    text: 'Что нового:\n• «Салоны рядом»: смежные салоны (тот же адрес, другие бренды МТС / МФ / Т2) с трафиком, конверсией и прогнозами по телефонам, аксессуарам, страховкам и комбо — из KPI ОСС.\n• Прогноз трафика на конец месяца.\n• В «Оформлении» — профиль интерфейса «iOS 16»: классический плоский вид, системные цвета, переключатели и простая анимация. Включается одной кнопкой.' },
  { v: '2026-11-21', title: 'Рейтинг по направлениям и титулы лидеров',
    text: 'Что нового:\n• Рейтинг команды: общий и отдельно по сим, телефонам, аксессуарам, услугам и страховкам.\n• Топ-5 — единая рамка «лидеры продаж».\n• Первое место по направлению — титул везде, где видно имя: «ценитель своего времени», «железный человек», «любитель побрякушек», «на страже защиты», «алло, я на связи».' },
  { v: '2026-11-20', title: 'Разбор ВМР: порог 80% и полная анкета',
    text: 'Что нового:\n• «На разбор» теперь — результат ниже 80%.\n• По клику на проверку открывается полная анкета: каждый критерий, итоговая оценка и комментарий проверяющего.\n• Для ТМ и РТТ/УТТ в быстром просмотре — отметки «🎁 Комбо» и «💳 Рассрочка» и динамика их предложения по периодам и ТТ.' },
  { v: '2026-11-19', title: 'Обновление: цвета, фон, входы, вложения и разбор ВМР',
    text: 'Что нового:\n• Оформление — кнопка-палитра в шапке: свои цвета, статусы, тема и фон (цвет или своя картинка).\n• Вкладка «Входы» для ТМ — история авторизации по каждому сотруднику.\n• Новости — можно прикладывать файлы и вставлять картинки из буфера.\n• «Проверки на разбор» (ВМР) — фильтр по месяцам и неделям и динамика: у ТМ — по всем, у РТТ/УТТ — по их ТТ, у ПТТ/СП — по их проверкам.\n• В аналитике темпа добавлен «Весь товарооборот».\n• Исправлен «График на 3 дня».\n• Уведомления (колокольчик): об обновлениях и о новых проверках ВМР.' }
];
function h_(t) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(t))).slice(0, 10); }
function notifSheet_(ss) {
  let sh = ss.getSheetByName(NOTIF_SHEET);
  if (!sh) { sh = ss.insertSheet(NOTIF_SHEET); sh.getRange(1, 1, 1, 7).setValues([['ID', 'Время', 'Кому', 'Тип', 'Заголовок', 'Текст', 'Вкладка']]); sh.setFrozenRows(1); }
  return sh;
}
function notifyMany_(ss, list) {            // list: [{to, kind, title, text, tab}] — to: ключ сотрудника или '*'
  if (!list.length) return;
  const sh = notifSheet_(ss), now = new Date().toISOString();
  const rows = list.map(function (n) { return [Utilities.getUuid().slice(0, 8), now, n.to, n.kind || '', safeCell_(n.title || ''), safeCell_(n.text || ''), n.tab || '']; });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 7).setNumberFormat('@').setValues(rows);
  try { pushSend_(list); } catch (e) {}
}
/* ---------- настоящие push-уведомления через OneSignal (REST). Нужны свойства скрипта ONESIGNAL_APP_ID и ONESIGNAL_API_KEY ---------- */
function pushSend_(list) {
  const props = PropertiesService.getScriptProperties(), cfg = (typeof PUSH_CFG_ !== 'undefined' && PUSH_CFG_) || {};
  const app = props.getProperty('ONESIGNAL_APP_ID') || cfg.appId || '', key = props.getProperty('ONESIGNAL_API_KEY') || cfg.apiKey || '';
  if (!app || !key || !list.length) return;
  const reqs = list.slice(0, 60).map(function (n) {
    const body = { app_id: app, target_channel: 'push', headings: { en: n.title, ru: n.title }, contents: { en: n.text, ru: n.text } };
    if (PUSH_URL_()) body.url = PUSH_URL_();
    if (n.to === '*') body.included_segments = ['Total Subscriptions']; else body.include_aliases = { external_id: [n.to] };
    return { url: 'https://api.onesignal.com/notifications', method: 'post', contentType: 'application/json', headers: { Authorization: 'Key ' + key }, payload: JSON.stringify(body), muteHttpExceptions: true };
  });
  UrlFetchApp.fetchAll(reqs);
}
function PUSH_URL_() { const cfg = (typeof PUSH_CFG_ !== 'undefined' && PUSH_CFG_) || {}; return PropertiesService.getScriptProperties().getProperty('SITE_URL') || cfg.siteUrl || ''; }
function notifList_(req) {
  const ss = ss_(), u = sessionUser_(ss, req.token);
  if (!u) return { error: 'auth' };
  try { maintain_(ss); } catch (e) {}
  const sh = ss.getSheetByName(NOTIF_SHEET), items = [];
  if (sh && sh.getLastRow() > 1) {
    const n = Math.min(400, sh.getLastRow() - 1), v = sh.getRange(sh.getLastRow() - n + 1, 1, n, 7).getValues(), lim = new Date(Date.now() - 30 * 864e5).toISOString();
    for (let i = v.length - 1; i >= 0 && items.length < 40; i--) {
      const r = v[i], to = String(r[2]);
      if ((to === '*' || to === u.key) && String(r[1]) >= lim) items.push({ id: String(r[0]), t: String(r[1]), kind: String(r[3]), title: String(r[4]), text: String(r[5]), tab: String(r[6]) });
    }
  }
  return { items: items, now: new Date().toISOString(), me: u.key };
}
// Фоновое обслуживание (не чаще раза в 5 минут): релизы → новость + уведомление всем; ВМР → уведомления сотруднику и РТТ/УТТ
function maintain_(ss) {
  const cc = CacheService.getScriptCache();
  if (cc.get('maint')) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return;
  try {
    cc.put('maint', '1', 300);
    try { releasesSync_(ss); } catch (e) {}
    try { vmrWatch_(ss); } catch (e) {}
  } finally { lock.releaseLock(); }
}
function releasesSync_(ss) {
  const props = PropertiesService.getScriptProperties(), done = (props.getProperty('rel_done') || '').split(',').filter(String);
  const todo = RELEASES.filter(function (r) { return done.indexOf(r.v) < 0; });
  if (!todo.length) return;
  const sh = newsSheet_(ss), when = Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm'), notes = [];
  todo.forEach(function (r) {
    const row = sh.getLastRow() + 1, rg = sh.getRange(row, 1, 1, 9); rg.setNumberFormat('@');
    rg.setValues([['rel-' + r.v, when, 'SUPER PROдажник', 'Обновление', safeCell_(r.title), safeCell_(r.text), ST_PUB, 'авто', '']]);
    notes.push({ to: '*', kind: 'release', title: '🚀 Скорее беги в приложение!', text: 'Там новая штука появилась 👀', tab: 'news' });
    done.push(r.v);
  });
  notifyMany_(ss, notes);
  props.setProperty('rel_done', done.join(','));
}
const VMR_SNAP = 'Снимок ВМР';
function vmrWatch_(ss) {
  const rep = tmReportRaw_(ss, 'vmr'); if (!rep || !rep.grid) return;
  const g = rep.grid; let hi = -1;
  for (let r = 0; r < Math.min(g.length, 15); r++) { const h = g[r].map(norm_); if (h.indexOf('тт') >= 0 && h.indexOf('оценка') >= 0) { hi = r; break; } }
  if (hi < 0) return;
  const H = g[hi].map(norm_), iTT = H.indexOf('тт'), iFio = H.indexOf('фио сотрудника'), iS = H.indexOf('оценка'), iD = H.indexOf('дата проверки'), iSc = H.indexOf('сценарий');
  if (iTT < 0 || iS < 0) return;
  const cur = {}, order = [];
  g.slice(hi + 1).forEach(function (r) {
    if (!String(r[iTT]).trim() || !String(r[iS]).trim()) return;
    const k = h_([r[iTT], iFio >= 0 ? r[iFio] : '', iD >= 0 ? r[iD] : '', iSc >= 0 ? r[iSc] : ''].join('|'));
    if (!cur[k]) order.push(k);
    cur[k] = { h: h_(r.join('|')), r: r };
  });
  let sh = ss.getSheetByName(VMR_SNAP), first = false;
  if (!sh) { sh = ss.insertSheet(VMR_SNAP); first = true; try { sh.hideSheet(); } catch (e) {} }
  const old = {};
  if (!first && sh.getLastRow() > 0) sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(function (v) { old[String(v[0])] = String(v[1]); });
  else first = true;
  const changed = order.filter(function (k) { return old[k] !== cur[k].h; });
  const save = function () {
    sh.clearContents();
    if (order.length) sh.getRange(1, 1, order.length, 2).setNumberFormat('@').setValues(order.map(function (k) { return [k, cur[k].h]; }));
  };
  if (first || !changed.length || changed.length > 300) { if (first || changed.length) save(); return; }   // первый запуск и массовая перезаливка — без уведомлений
  const access = readAccess_(ss), shops = kpiShops_(ss, access);
  const mgrs = access.filter(function (a) { return a.manager && a.roleTitle !== 'ТМ'; }), emps = {};
  access.forEach(function (a) { if (!a.manager) emps[a.key] = a; });
  const box = {};                                                    // ключ получателя → { me: [], tt: [] }
  const add = function (key, kind, it) { (box[key] = box[key] || { me: [], tt: [] })[kind].push(it); };
  changed.forEach(function (k) {
    const r = cur[k].r, it = { tt: String(r[iTT]), fio: iFio >= 0 ? String(r[iFio]) : '', d: iD >= 0 ? String(r[iD]) : '', sc: iSc >= 0 ? String(r[iSc]) : '', s: parseFloat(String(r[iS]).replace('%', '').replace(',', '.')) || 0, isNew: !old[k] };
    const ek = norm_(it.fio); if (emps[ek]) add(ek, 'me', it);
    mgrs.forEach(function (m) { if ((shops[m.key] || []).some(function (sh2) { return sameSalon_(it.tt, sh2); })) add(m.key, 'tt', it); });
  });
  const shortTT = function (t) { return String(t).replace(/^КЗН\s+/, '').replace(/\s+(БЛН|МТС|МФ)\b.*$/, '').trim() || t; };
  const out = [], em = function (x) { return x >= 100 ? '🏆' : x >= 80 ? '✅' : '⚠️'; };
  Object.keys(box).forEach(function (key) {
    const b = box[key];
    if (b.me.length > 5) out.push({ to: key, kind: 'vmr', tab: 'vmr', title: '📋 Новые анкеты ВМР', text: 'На тебя пришло анкет: ' + b.me.length + ' — загляни в приложение 👀' });
    else b.me.forEach(function (x) { out.push({ to: key, kind: 'vmr', tab: 'vmr', title: '📋 Пришла анкета ВМР', text: 'На тебя пришла анкета с результатом ' + em(x.s) + ' ' + Math.round(x.s) + '%' }); });
    if (b.tt.length > 5) out.push({ to: key, kind: 'vmr', tab: 'vmr', title: '👥 Анкеты по твоим сотрудникам', text: 'На твоих сотрудников пришло анкет: ' + b.tt.length + ' — загляни в приложение 👀' });
    else b.tt.forEach(function (x) { out.push({ to: key, kind: 'vmr', tab: 'vmr', title: '👥 Анкета на твоего сотрудника', text: 'На твоего сотрудника ' + (x.fio.split(' ').slice(0, 2).join(' ') || '—') + ' (' + shortTT(x.tt) + ') пришла анкета с результатом ' + em(x.s) + ' ' + Math.round(x.s) + '%' }); });
  });
  notifyMany_(ss, out);
  save();
}

/* ---------- рейтинг по направлениям и «лидеры» (первое место по группе) ---------- */
function ratingsFor_(sellers, meKey) {
  const out = {};
  METRICS.forEach(function (m) {
    const l = sellers.map(function (s) { return { s: s, m: s.metrics.filter(function (x) { return x.key === m.key; })[0] }; })
      .filter(function (o) { return o.m && o.m.planMonth > 0; })
      .sort(function (a, b) { return b.m.pct - a.m.pct || b.m.fact - a.m.fact; });
    out[m.key] = l.map(function (o, i) { return { place: i + 1, name: short_(o.s.name), score: Math.round(o.m.pct), me: norm_(o.s.name) === meKey }; });
  });
  return out;
}
function leaders_(ss, access) {
  const allowed = {}; access.forEach(function (a) { if (!a.manager) allowed[a.key] = true; });
  const sellers = readPlan_(ss).filter(function (s) { return allowed[norm_(s.name)]; }), r = ratingsFor_(sellers, ''), out = {};
  Object.keys(r).forEach(function (k) { if (r[k].length && r[k][0].score > 0) out[k] = r[k][0].name; });
  return out;
}
