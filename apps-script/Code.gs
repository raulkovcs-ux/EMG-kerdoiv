// sEMG adatbázis – önkéntes jelentkezés. A táblázathoz kötött Apps Script.

var SHEET_NAME = 'Jelentkezések';
var HEADERS = ['Időbélyeg', 'Név', 'Nem', 'Kor', 'Kezesség', 'Telefon', 'E-mail',
               'Dátum', 'Kezdés', 'Slot', 'Hozzájárulás', 'Emlékeztető elküldve'];
var COL = { DATE: 8, START: 9, KEY: 10, REMINDED: 12 };

// Napok és kezdési órák (2026). Egy slot = 1 óra, egy önkéntes.
var DAYS = [
  { date: '2026-10-23', hours: [9, 10, 11, 14, 15, 16] },
  { date: '2026-10-29', hours: [9, 10, 11, 14, 15, 16] },
  { date: '2026-11-02', hours: [14, 15, 16] },
  { date: '2026-11-03', hours: [9, 10, 11, 14, 15, 16] },
  { date: '2026-11-05', hours: [9, 10, 11, 14, 15, 16] }
];
var GENDERS = ['Nő', 'Férfi', 'Egyéb', 'Nem kívánok válaszolni'];
var HANDS = ['Jobbkezes', 'Balkezes'];
var ARRIVE_EARLY_MIN = 10;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('sEMG önkéntes jelentkezés')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function pad_(n) { return (n < 10 ? '0' : '') + n; }
function slotKey_(date, hour) { return date + ' ' + pad_(hour) + ':00'; }

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // Szövegként tároljuk, hogy a Sheets ne alakítsa át dátummá/számmá.
    [COL.DATE, COL.START, COL.KEY, 6].forEach(function (c) {
      sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  }
  return sh;
}

function bookedKeys_() {
  var sh = getSheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) return {};
  var vals = sh.getRange(2, COL.KEY, n, 1).getValues();
  var out = {};
  vals.forEach(function (r) { if (r[0]) out[String(r[0])] = true; });
  return out;
}

// A kliens ezt hívja: csak a szabad időpontokat adja vissza.
function getAvailableSlots() {
  var booked = bookedKeys_();
  var days = [];
  DAYS.forEach(function (d) {
    var slots = [];
    d.hours.forEach(function (h) {
      if (!booked[slotKey_(d.date, h)]) {
        slots.push({ hour: h, label: pad_(h) + ':00–' + pad_(h + 1) + ':00' });
      }
    });
    if (slots.length) days.push({ date: d.date, slots: slots });
  });
  return { days: days, earlyMin: ARRIVE_EARLY_MIN, genders: GENDERS, hands: HANDS };
}

function submitRegistration(f) {
  var name = String(f.name || '').trim();
  var age = parseInt(f.age, 10);
  var phone = String(f.phone || '').trim();
  var email = String(f.email || '').trim();
  var hour = parseInt(f.hour, 10);
  var date = String(f.date || '');

  if (!name) throw new Error('Add meg a neved.');
  if (GENDERS.indexOf(f.gender) < 0) throw new Error('Válassz nemet.');
  if (!(age >= 14 && age <= 99)) throw new Error('Adj meg érvényes életkort.');
  if (HANDS.indexOf(f.hand) < 0) throw new Error('Válaszd ki a domináns kezed.');
  if (phone.replace(/\D/g, '').length < 7) throw new Error('Adj meg érvényes telefonszámot.');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Adj meg érvényes e-mail címet.');
  if (f.consent !== true) throw new Error('Az adatkezelési hozzájárulás kötelező.');

  var valid = DAYS.some(function (d) { return d.date === date && d.hours.indexOf(hour) >= 0; });
  if (!valid) throw new Error('Érvénytelen időpont.');

  var key = slotKey_(date, hour);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (bookedKeys_()[key]) {
      throw new Error('Ezt az időpontot közben valaki lefoglalta. Válassz másikat.');
    }
    getSheet_().appendRow([new Date(), name, f.gender, age, f.hand, phone, email,
                           date, pad_(hour) + ':00', key, 'igen', '']);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  var arrive = pad_(hour) + ':00';
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'sEMG mérés – időpont visszaigazolás',
      body: 'Kedves ' + name + '!\n\nKöszönjük a jelentkezést az sEMG adatbázis építéséhez.\n\n' +
            'Időpontod: ' + date + ', ' + pad_(hour) + ':00–' + pad_(hour + 1) + ':00\n' +
            'Kérjük, legalább ' + ARRIVE_EARLY_MIN + ' perccel korábban érkezz.\n\n' +
            'A mérés előtt a nap folyamán emlékeztetőt is küldünk.'
    });
  } catch (e) { /* a jelentkezés ettől még érvényes */ }

  return { date: date, label: arrive + '–' + pad_(hour + 1) + ':00' };
}

// Opcionális: napi időzítővel (pl. 8:00) fusson. Az aznapi időpontokra küld emlékeztetőt.
function sendReminders() {
  var sh = getSheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) return;
  var today = Utilities.formatDate(new Date(), 'Europe/Budapest', 'yyyy-MM-dd');
  var rows = sh.getRange(2, 1, n, HEADERS.length).getValues();
  rows.forEach(function (r, i) {
    if (String(r[COL.DATE - 1]) === today && !r[COL.REMINDED - 1]) {
      MailApp.sendEmail({
        to: r[6],
        subject: 'Emlékeztető: ma sEMG mérés ' + r[COL.START - 1] + '-kor',
        body: 'Kedves ' + r[1] + '!\n\nEmlékeztető: ma ' + r[COL.START - 1] + '-kor van a mérésed. ' +
              'Kérjük, legalább ' + ARRIVE_EARLY_MIN + ' perccel előbb érkezz.'
      });
      sh.getRange(i + 2, COL.REMINDED).setValue(new Date());
    }
  });
}
