/**
 * KK — Telegram-бот передзамовлень. Безкоштовно: Google Таблиця + Google Apps Script.
 *
 * Аркуші таблиці:
 *   Магазини  : id | name | addr | hours | chat_id        (chat_id — Telegram-група магазину, куди падають замовлення)
 *   Товари    : id | name | vol | price | cat | img | note
 *   Залишки   : item_id | <id магазину 1> | <id магазину 2> | …   (кількість у штуках)
 *   Замовлення: створюється автоматично
 *
 * Властивості скрипта (Project Settings → Script properties):
 *   BOT_TOKEN  — токен від @BotFather
 *   WEBAPP_URL — адреса міні-застосунку, напр. https://andreyfedorovskiy.github.io/kk-design/tg/
 */

const P = PropertiesService.getScriptProperties();
const TOKEN = () => P.getProperty('BOT_TOKEN');
const SS = () => SpreadsheetApp.getActive();

/* ---------- GET: магазини, товари, залишки для міні-застосунку ---------- */
function doGet() {
  const stores = rows_('Магазини').map(r => ({ id: r.id, name: r.name, addr: r.addr, hours: r.hours }));
  const items = rows_('Товари').map(r => ({ id: r.id, name: r.name, vol: r.vol, price: +r.price, cat: r.cat, img: r.img || null, note: r.note || '' }));
  const stock = {};
  rows_('Залишки').forEach(r => stores.forEach(s => { (stock[s.id] = stock[s.id] || {})[r.item_id] = +r[s.id] || 0; }));
  const updated = Utilities.formatDate(new Date(), 'Europe/Kyiv', 'yyyy-MM-dd HH:mm');
  return json_({ stores, items, stock, updated });
}

/* ---------- POST: або замовлення з міні-застосунку, або вебхук Telegram ---------- */
function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  if (body.update_id) return handleUpdate_(body);
  return handleOrder_(body);
}

function handleOrder_(o) {
  const user = verifyInitData_(o.initData);          // захист від підробки
  if (!user) return json_({ ok: false, error: 'bad initData' });
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    // 1) списуємо залишок (резерв)
    const sh = SS().getSheetByName('Залишки'), v = sh.getDataRange().getValues();
    const col = v[0].indexOf(o.storeId);
    o.lines.forEach(l => {
      const r = v.findIndex(row => row[0] === l.id);
      if (r > 0 && col > 0) { const have = +v[r][col] || 0; sh.getRange(r + 1, col + 1).setValue(Math.max(have - l.q, 0)); }
    });
    // 2) пишемо в журнал
    const log = SS().getSheetByName('Замовлення') || SS().insertSheet('Замовлення');
    if (log.getLastRow() === 0) log.appendRow(['№', 'Створено', 'Магазин', 'Коли', 'Клієнт', 'Телефон', 'tg_id', 'Склад', 'Сума', 'Статус']);
    log.appendRow([o.no, new Date(), o.store, o.when, o.name, o.phone, user.id,
      o.lines.map(l => `${l.name} ${l.vol} ×${l.q}`).join('; '), o.total, 'Нове']);
  } finally { lock.releaseLock(); }
  // 3) повідомлення в чат магазину з кнопками
  const st = rows_('Магазини').find(s => s.id === o.storeId);
  if (st && st.chat_id) tg_('sendMessage', { chat_id: st.chat_id, text: orderText_(o, user), reply_markup: { inline_keyboard: [[
    { text: '✅ Підтвердити', callback_data: `ok|${o.no}|${user.id}` },
    { text: '📦 Зібрано', callback_data: `ready|${o.no}|${user.id}` },
    { text: '✖ Скасувати', callback_data: `cancel|${o.no}|${user.id}` }]] } });
  // 4) підтвердження клієнту
  tg_('sendMessage', { chat_id: user.id, text: `Дякуємо! Передзамовлення ${o.no} прийнято.\n📍 ${o.store}\n🕒 ${o.when}\nНапишемо, коли все буде зібрано.` });
  return json_({ ok: true, no: o.no });
}

function handleUpdate_(u) {
  const webapp = P.getProperty('WEBAPP_URL');
  if (u.message && /^\/start/.test(u.message.text || '')) {
    const storeParam = (u.message.text.split(' ')[1] || '');         // t.me/bot?start=kyiv-1 → одразу потрібний магазин
    tg_('sendMessage', { chat_id: u.message.chat.id,
      text: 'Вітаємо в KK! Оберіть магазин, перегляньте наявність і відкладіть товар до свого приходу.',
      reply_markup: { inline_keyboard: [[{ text: '🛒 Наявність і передзамовлення', web_app: { url: webapp + (storeParam ? '?store=' + encodeURIComponent(storeParam) : '') } }]] } });
  }
  if (u.callback_query) {
    const [act, no, clientId] = u.callback_query.data.split('|');
    const label = { ok: 'Підтверджено', ready: 'Зібрано — чекаємо', cancel: 'Скасовано' }[act];
    setStatus_(no, label);
    const toClient = {
      ok: `✅ Замовлення ${no} підтверджено магазином.`,
      ready: `📦 Замовлення ${no} зібрано і чекає на вас. Резерв діє 24 години.`,
      cancel: `Замовлення ${no} скасовано. Якщо це помилка — напишіть нам.`
    }[act];
    if (toClient) tg_('sendMessage', { chat_id: clientId, text: toClient });
    tg_('answerCallbackQuery', { callback_query_id: u.callback_query.id, text: label });
    tg_('editMessageText', { chat_id: u.callback_query.message.chat.id, message_id: u.callback_query.message.message_id,
      text: u.callback_query.message.text + `\n\nСтатус: ${label} (${u.callback_query.from.first_name})`,
      reply_markup: act === 'cancel' || act === 'ready' ? { inline_keyboard: [] } : u.callback_query.message.reply_markup });
  }
  return json_({ ok: true });
}

/* ---------- Нагадування: щогодини (тригер за часом) знімає прострочені резерви ---------- */
function expireReservations() {
  const log = SS().getSheetByName('Замовлення'); if (!log) return;
  const v = log.getDataRange().getValues(), now = Date.now();
  for (let r = 1; r < v.length; r++) {
    if (['Нове', 'Підтверджено', 'Зібрано — чекаємо'].includes(v[r][9]) && now - new Date(v[r][1]).getTime() > 48 * 3600e3) {
      log.getRange(r + 1, 10).setValue('Прострочено');   // залишок магазин повертає вручну або доробимо автоповернення
    }
  }
}

/* ---------- Разове налаштування: запустити 1 раз після публікації веб-застосунку ---------- */
function setWebhook() {
  const url = ScriptApp.getService().getUrl();
  Logger.log(tg_('setWebhook', { url, drop_pending_updates: true }));
  tg_('setChatMenuButton', { menu_button: { type: 'web_app', text: 'Наявність', web_app: { url: P.getProperty('WEBAPP_URL') } } });
}

/* ---------- службове ---------- */
function rows_(name) {
  const v = SS().getSheetByName(name).getDataRange().getValues(), h = v.shift();
  return v.filter(r => r[0] !== '').map(r => Object.fromEntries(h.map((k, i) => [String(k).trim(), r[i]])));
}
function setStatus_(no, label) {
  const log = SS().getSheetByName('Замовлення'), v = log.getDataRange().getValues();
  const r = v.findIndex(row => row[0] === no); if (r > 0) log.getRange(r + 1, 10).setValue(label);
}
function orderText_(o, user) {
  return `🧾 Передзамовлення ${o.no}\n📍 ${o.store}\n🕒 ${o.when}\n👤 ${o.name} · ${o.phone}${user.username ? ' · @' + user.username : ''}\n\n` +
    o.lines.map(l => `• ${l.name} · ${l.vol} × ${l.q} = ${l.sum} грн${l.toOrder ? `  (під замовлення: ${l.toOrder})` : ''}`).join('\n') +
    `\n\nРазом: ${o.total} грн${o.comment ? '\n💬 ' + o.comment : ''}`;
}
function tg_(method, payload) {
  const r = UrlFetchApp.fetch(`https://api.telegram.org/bot${TOKEN()}/${method}`,
    { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  return r.getContentText();
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Перевірка підпису Telegram initData (https://core.telegram.org/bots/webapps#validating-data-received-via-the-mini-app) */
function verifyInitData_(initData) {
  if (!initData) return null;
  const params = {}; initData.split('&').forEach(p => { const i = p.indexOf('='); params[p.slice(0, i)] = decodeURIComponent(p.slice(i + 1)); });
  const hash = params.hash; delete params.hash;
  const check = Object.keys(params).sort().map(k => `${k}=${params[k]}`).join('\n');
  const secret = Utilities.computeHmacSha256Signature(Utilities.newBlob(TOKEN()).getBytes(), Utilities.newBlob('WebAppData').getBytes());
  const sig = Utilities.computeHmacSha256Signature(Utilities.newBlob(check).getBytes(), secret);
  const hex = sig.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  return hex === hash ? JSON.parse(params.user || '{}') : null;
}
