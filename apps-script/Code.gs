/**
 * ערב נשים בבתים – צד שרת להרשמה (Google Apps Script)
 * מחובר לגיליון "הרשמות – ערב נשים בבתים".
 *
 * GET  -> מחזיר מקומות פנויים לכל בית ומשבצות אוכל פנויות
 * POST -> רושם משתתפת (גוף JSON כ-text/plain): {name, phone, house, slot}
 */

const CAPACITY = 15; // נרשמות לכל בית

const HOUSES = [
  { id: '1', name: 'מאסטר שף' },
  { id: '2', name: 'רק בבריאות' },
  { id: '3', name: 'מלאכת מחשבת' },
  { id: '4', name: 'להראות בטוב' },
  { id: '5', name: 'הנאה ביצירה' },
];

// משבצות אוכל לכל בית (סה"כ 10), כל משבצת נלקחת פעם אחת
const SLOTS = [
  { id: 'soup',    name: 'מרק',                 count: 2 },
  { id: 'dessert', name: 'משהו מתוק ליד הקפה',  count: 2, aliases: ['קינוח'] },
  { id: 'cracker', name: 'חבילת קרקרים',        count: 2 },
  { id: 'cheese',  name: 'גבינה',               count: 2 },
  { id: 'veggies', name: 'מגש ירקות חתוכים',    count: 1 },
  { id: 'milk',    name: 'חלב',                 count: 1 },
];

const SHEET_NAME = 'הרשמות';
const HEADERS = ['חותמת זמן', 'שם מלא', 'טלפון', 'בית', 'משבצת אוכל'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

/** הרצה ידנית פעם אחת: מסדר כותרות, מקפיא שורה ראשונה, ומוסיף לשונית סיכום. */
function setup() {
  const sh = getSheet_();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setRightToLeft(true);
  sh.getRange('C:C').setNumberFormat('@'); // טלפון כטקסט, שלא יאבד 0 מוביל

  const ss = sh.getParent();
  let sum = ss.getSheetByName('סיכום');
  if (!sum) sum = ss.insertSheet('סיכום');
  sum.clear();
  sum.setRightToLeft(true);
  const rows = [['בית', 'נרשמות', 'מקומות פנויים']];
  HOUSES.forEach((h, i) => {
    const r = i + 2;
    rows.push([h.name, `=COUNTIF('${SHEET_NAME}'!D:D,A${r})`, `=${CAPACITY}-B${r}`]);
  });
  sum.getRange(1, 1, rows.length, 3).setValues(rows);
  sum.getRange(1, 1, 1, 3).setFontWeight('bold');
}

function readState_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  const data = last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
  const state = {};
  HOUSES.forEach(h => {
    const slots = {};
    SLOTS.forEach(s => (slots[s.id] = { name: s.name, remaining: s.count }));
    state[h.name] = { id: h.id, name: h.name, registered: 0, slots, phones: {} };
  });
  data.forEach(row => {
    const h = state[row[3]];
    if (!h) return;
    h.registered++;
    h.phones[normPhone_(row[2])] = true;
    const slot = SLOTS.find(s => s.name === row[4] || (s.aliases || []).includes(row[4]));
    if (slot) h.slots[slot.id].remaining--;
  });
  return state;
}

function publicState_(state) {
  return {
    capacity: CAPACITY,
    houses: HOUSES.map(h => {
      const s = state[h.name];
      return {
        id: h.id,
        name: h.name,
        remaining: Math.max(0, CAPACITY - s.registered),
        slots: SLOTS.map(sl => ({
          id: sl.id,
          name: sl.name,
          remaining: Math.max(0, s.slots[sl.id].remaining),
        })),
      };
    }),
  };
}

function normPhone_(p) {
  return String(p || '').replace(/\D/g, '');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_(publicState_(readState_()));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return json_({ ok: false, error: 'busy' });
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const name = String(body.name || '').trim().slice(0, 80);
    const phone = normPhone_(body.phone);
    const house = HOUSES.find(h => h.id === String(body.house));
    const slotDef = body.slot ? SLOTS.find(s => s.id === String(body.slot)) : null;

    if (!name) return json_({ ok: false, error: 'missing_name' });
    if (phone.length < 9 || phone.length > 12) return json_({ ok: false, error: 'bad_phone' });
    if (!house) return json_({ ok: false, error: 'bad_house' });
    if (body.slot && !slotDef) return json_({ ok: false, error: 'bad_slot' });

    const state = readState_();
    const h = state[house.name];
    if (h.registered >= CAPACITY) return json_({ ok: false, error: 'house_full', state: publicState_(state) });
    if (h.phones[phone]) return json_({ ok: false, error: 'already_registered' });
    if (slotDef && h.slots[slotDef.id].remaining <= 0)
      return json_({ ok: false, error: 'slot_taken', state: publicState_(state) });

    getSheet_().appendRow([new Date(), name, "'" + phone, house.name, slotDef ? slotDef.name : '']);
    return json_({ ok: true, state: publicState_(readState_()) });
  } catch (err) {
    return json_({ ok: false, error: 'server_error' });
  } finally {
    lock.releaseLock();
  }
}
