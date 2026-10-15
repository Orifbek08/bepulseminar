/**
 * БУХОРО СЕМИНАРИ — БИЛЕТ ТИЗИМИ (Google Sheets backend)
 * Илҳом Бегимқулов · Бепул семинар · 2 ноябрь, 2026
 *
 * «Бухоро» листи: A: Сана | B: Исм фамилия | C: Телефон | D: ID рақам | E: Вилоят
 * Лист бўлмаса — ўзи яратилади. Бошқа листларга тегилмайди.
 * Битта телефон рақам — битта билет (фақат «Бухоро» листи ичида).
 */

var SHEET_NAME = 'Бухоро';
var ID_PREFIX  = 'BX-2026-';
var REGION     = 'Бухоро';
var TZ         = 'Asia/Tashkent';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var data  = JSON.parse(e.postData.contents);
    var name  = String(data.name || '').trim().replace(/\s+/g, ' ');
    var phone = normPhone(data.phone);

    if (!name || !phone) return json({ status: 'error', message: 'Исм ёки телефон хато' });

    var sheet  = getSheet();
    var values = sheet.getDataRange().getValues();

    // 1) Телефон бор бўлса — эски билет қайтарилади
    for (var i = 1; i < values.length; i++) {
      if (normPhone(values[i][2]) === phone) {
        return json(row(values[i], 'exists'));
      }
    }

    // 2) Янги билет: кейинги тартиб рақами (сарлавҳа ҳисобга олинмайди)
    var ticketId = ID_PREFIX + padLeft(nextNumber(values), 4);
    var date = Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm');
    sheet.appendRow([date, name, phone, ticketId, REGION]);
    sheet.getRange(sheet.getLastRow(), 3).setNumberFormat('@'); // телефон матн бўлиб қолсин

    return json({ status: 'created', ticketId: ticketId, name: name, phone: phone, region: REGION, date: date });

  } catch (err) {
    return json({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Текшириш: .../exec?phone=%2B998901234567  ёки  .../exec (ишлаяптими) */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!p.phone) return json({ status: 'ok', sheet: SHEET_NAME });

  var phone = normPhone(p.phone);
  if (!phone) return json({ status: 'error', message: 'Телефон хато' });

  var values = getSheet().getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (normPhone(values[i][2]) === phone) return json(row(values[i], 'exists'));
  }
  return json({ status: 'not_found' });
}

/* ---------- ёрдамчи функциялар ---------- */

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Сана', 'Исм фамилия', 'Телефон', 'ID рақам', 'Вилоят']);
    sheet.getRange('A1:E1').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('C:C').setNumberFormat('@');
  }
  return sheet;
}

/** Энг катта мавжуд ID дан кейингиси — қатор ўчирилса ҳам такрорланмайди */
function nextNumber(values) {
  var max = 0;
  for (var i = 1; i < values.length; i++) {
    var m = String(values[i][3]).match(/(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }
  return max + 1;
}

function row(r, status) {
  return {
    status: status, ticketId: r[3], name: r[1], phone: normPhone(r[2]),
    region: r[4], date: r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'dd.MM.yyyy HH:mm') : String(r[0])
  };
}

/** Ҳар хил ёзилган рақамларни битта кўринишга келтиради: +998901234567 */
function normPhone(raw) {
  var d = String(raw == null ? '' : raw).replace(/\D/g, '');
  if (d.indexOf('998') === 0 && d.length === 12) d = d.substring(3);
  return d.length === 9 ? '+998' + d : null;
}

function padLeft(n, len) {
  var s = String(n);
  while (s.length < len) s = '0' + s;
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
