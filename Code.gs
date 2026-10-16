/**
 * „Къде ми е?“ — сървърна част в твоята Google таблица.
 *
 * Как се ползва (еднократно):
 *  1. Отвори таблицата → Разширения → Apps Script, изтрий всичко и постави този код.
 *  2. Избери функцията setup отгоре и натисни „Изпълни“. Разреши достъпа.
 *     В „Дневник на изпълнението“ ще видиш ТВОЯ КЛЮЧ — копирай го.
 *  3. Внедряване → Ново внедряване → тип „Уеб приложение“
 *     Изпълни като: „Аз“, Кой има достъп: „Всеки“ → Внедри. Копирай адреса (URL).
 *  4. В приложението → ⚙ Настройки постави адреса и ключа.
 *
 * Допълнителни функции (избираш ги отгоре и натискаш „Изпълни“):
 *  - enableWeeklyEmail   включва имейл всеки понеделник с изтичащите срокове
 *  - disableWeeklyEmail  изключва имейла
 *  - newKey              сменя ключа — всички свързани устройства губят достъп,
 *                        докато не им дадеш новия ключ (нова покана)
 */

const ITEM_SHEET = 'Предмети';
const PLACE_SHEET = 'Места';

// Редът на колоните е важен — не ги размествай.
const ITEM_F = ['id', 'name', 'category', 'expiry', 'note', 'location', 'code',
                'addedBy', 'createdAt', 'updatedAt', 'deleted', 'serverTs'];
const ITEM_H = ['ID', 'Име', 'Категория', 'Годно до', 'Бележка', 'Описание на мястото', 'Код на мястото',
                'Добавил', 'Създадено', 'Променено', 'Изтрито', 'служебно'];
const PLACE_F = ['code', 'description', 'addedBy', 'updatedAt', 'deleted', 'serverTs'];
const PLACE_H = ['Код', 'Описание', 'Добавил', 'Променено', 'Изтрито', 'служебно'];

/* ───────────── Настройка ───────────── */

function setup() {
  const ss = SpreadsheetApp.getActive();
  prepSheet_(ss, ITEM_SHEET, ITEM_H);
  prepSheet_(ss, PLACE_SHEET, PLACE_H);
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('KEY');
  if (!key) { key = makeKey_(); props.setProperty('KEY', key); }
  Logger.log('ТВОЯТ КЛЮЧ: ' + key);
}

function newKey() {
  const key = makeKey_();
  PropertiesService.getScriptProperties().setProperty('KEY', key);
  Logger.log('НОВ КЛЮЧ: ' + key);
}

function makeKey_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 20);
}

function prepSheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  // Всичко като текст, за да не превръща Google датите и кодовете в числа.
  sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
  sh.hideColumns(headers.length); // служебната колона
}

/* ───────────── Връзка с приложението ───────────── */

function doGet() {
  return ContentService.createTextOutput('„Къде ми е?“ — връзката работи.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    const req = JSON.parse(e.postData.contents);
    const key = PropertiesService.getScriptProperties().getProperty('KEY');
    if (!key || req.key !== key) return out_({ ok: false, error: 'Грешен ключ' });

    lock.waitLock(25000);
    locked = true;
    const ss = SpreadsheetApp.getActive();
    const itemSh = ss.getSheetByName(ITEM_SHEET);
    const placeSh = ss.getSheetByName(PLACE_SHEET);
    if (!itemSh || !placeSh) return out_({ ok: false, error: 'Не е изпълнена функцията setup' });

    if (req.action === 'ping') {
      return out_({ ok: true, items: Math.max(0, itemSh.getLastRow() - 1), places: Math.max(0, placeSh.getLastRow() - 1) });
    }

    const now = nextTs_();
    const since = Number(req.since) || 0;
    const rejI = upsert_(itemSh, ITEM_F, 'id', req.items || [], now);
    const rejP = upsert_(placeSh, PLACE_F, 'code', req.places || [], now);

    return out_({
      ok: true,
      serverTs: now,
      items: rowsSince_(itemSh, ITEM_F, since).concat(rejI),
      places: rowsSince_(placeSh, PLACE_F, since).concat(rejP)
    });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Строго нарастващ брояч на промените. */
function nextTs_() {
  const props = PropertiesService.getScriptProperties();
  const last = Number(props.getProperty('LAST_TS')) || 0;
  const ts = Math.max(Date.now(), last + 1);
  props.setProperty('LAST_TS', String(ts));
  return ts;
}

function cell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

function readAll_(sh, F) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, F.length).getValues().map(r => {
    const o = {};
    F.forEach((f, j) => { o[f] = cell_(r[j]); });
    return o;
  });
}

function time_(s) { return Date.parse(s) || 0; }

/** Записва промените; по-новата версия печели. Връща сървърните версии, които са били по-нови. */
function upsert_(sh, F, keyF, recs, now) {
  if (!recs.length) return [];
  const rows = readAll_(sh, F);
  const index = {};
  rows.forEach((r, i) => { if (r[keyF]) index[r[keyF]] = i; });
  const rejected = [];

  recs.forEach(rec => {
    let k = String(rec[keyF] || '').trim();
    if (keyF === 'code') k = k.toUpperCase();
    if (!k) return;
    const i = index[k];
    if (i !== undefined && time_(rows[i].updatedAt) > time_(rec.updatedAt)) {
      rejected.push(rows[i]);
      return;
    }
    const o = {};
    F.forEach(f => {
      let v = rec[f];
      if (f === 'deleted') v = v ? '1' : '';
      o[f] = v === null || v === undefined ? '' : String(v);
    });
    o[keyF] = k;
    o.serverTs = String(now);
    if (i !== undefined) rows[i] = o;
    else { index[k] = rows.length; rows.push(o); }
  });

  if (rows.length) {
    const rg = sh.getRange(2, 1, rows.length, F.length);
    rg.setNumberFormat('@');
    rg.setValues(rows.map(o => F.map(f => o[f])));
  }
  return rejected;
}

function rowsSince_(sh, F, since) {
  return readAll_(sh, F).filter(o => (Number(o.serverTs) || 0) > since);
}

/** Ако промениш нещо направо в таблицата, приложението ще го изтегли. */
function onEdit(e) {
  const sh = e.range.getSheet();
  const name = sh.getName();
  const F = name === ITEM_SHEET ? ITEM_F : name === PLACE_SHEET ? PLACE_F : null;
  if (!F || e.range.getRow() < 2) return;
  const ts = String(nextTs_());
  const iso = new Date().toISOString();
  const updCol = F.indexOf('updatedAt') + 1;
  const tsCol = F.indexOf('serverTs') + 1;
  for (let r = e.range.getRow(); r < e.range.getRow() + e.range.getNumRows(); r++) {
    const first = sh.getRange(r, 1);
    if (!first.getValue()) {
      if (name === ITEM_SHEET && sh.getRange(r, 2).getValue()) first.setValue(Utilities.getUuid());
      else continue;
    }
    sh.getRange(r, updCol).setValue(iso);
    sh.getRange(r, tsCol).setValue(ts);
  }
}

/* ───────────── Седмичен имейл ───────────── */

function enableWeeklyEmail() {
  disableWeeklyEmail();
  ScriptApp.newTrigger('weeklyEmail').timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  Logger.log('Имейлът е включен: всеки понеделник около 8:00.');
}

function disableWeeklyEmail() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'weeklyEmail')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function weeklyEmail() {
  const ss = SpreadsheetApp.getActive();
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const limit = Utilities.formatDate(new Date(Date.now() + 14 * 864e5), tz, 'yyyy-MM-dd');
  const places = {};
  readAll_(ss.getSheetByName(PLACE_SHEET), PLACE_F).forEach(p => { if (!p.deleted) places[p.code] = p.description; });

  const list = readAll_(ss.getSheetByName(ITEM_SHEET), ITEM_F)
    .filter(i => !i.deleted && i.expiry && i.expiry <= limit)
    .sort((a, b) => a.expiry < b.expiry ? -1 : 1);
  if (!list.length) return;

  const esc = s => String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  const fmt = d => d.split('-').reverse().join('.');
  const rows = list.map(i => {
    const where = [i.code ? 'код ' + i.code + (places[i.code] ? ' (' + places[i.code] + ')' : '') : '', i.location]
      .filter(Boolean).join(', ');
    const color = i.expiry < today ? '#c0392b' : '#b7791f';
    return '<tr><td style="padding:6px 10px"><b>' + esc(i.name) + '</b></td>' +
      '<td style="padding:6px 10px;color:' + color + '">' + (i.expiry < today ? 'изтекъл ' : 'до ') + fmt(i.expiry) + '</td>' +
      '<td style="padding:6px 10px;color:#555">' + esc(where) + '</td></tr>';
  }).join('');

  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: 'Къде ми е? — ' + list.length + ' неща изтичат скоро',
    htmlBody: '<p>Изтекли или изтичащи до 2 седмици:</p><table style="border-collapse:collapse;font-family:sans-serif">' + rows + '</table>'
  });
}
