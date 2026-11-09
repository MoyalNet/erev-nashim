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

/* ===== תצוגות מסודרות (לא נוגע בשורות ההרשמה) ===== */
const HOUSE_COLORS = ['#f6d5c9', '#d9ead3', '#fce5cd', '#ead1dc', '#d0e0e3'];
const WINE = '#8c2f3a';

function setup() { buildViews(); }

function buildViews() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = getSheet_();
  const R = "'" + SHEET_NAME + "'!";

  // --- לשונית הנתונים: רק עיצוב, בלי לשנות ערכים ---
  raw.setRightToLeft(true);
  raw.setFrozenRows(1);
  raw.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
     .setFontWeight('bold').setBackground(WINE).setFontColor('#ffffff').setHorizontalAlignment('center');
  raw.getRange('A2:A').setNumberFormat('dd/MM HH:mm');
  raw.getRange('C2:C').setNumberFormat('@');
  [120, 170, 120, 130, 170].forEach((w, i) => raw.setColumnWidth(i + 1, w));
  raw.setTabColor('#999999');
  const rules = HOUSES.map((h, i) => SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="' + h.name + '"').setBackground(HOUSE_COLORS[i])
      .setRanges([raw.getRange('A2:E1000')]).build());
  raw.setConditionalFormatRules(rules);

  // --- לשונית סיכום ---
  let sum = ss.getSheetByName('סיכום') || ss.insertSheet('סיכום');
  sum.clear(); sum.clearConditionalFormatRules(); sum.setRightToLeft(true);
  const totalSlots = SLOTS.reduce((a, s) => a + s.count, 0);
  const rows = [['בית / סדנה', 'נרשמו', 'נותרו מקומות', 'משבצות אוכל שנלקחו', 'מתוך']];
  HOUSES.forEach((h, i) => {
    const r = i + 2;
    rows.push([h.name,
      '=COUNTIF(' + R + 'D2:D,A' + r + ')',
      '=' + CAPACITY + '-B' + r,
      '=COUNTIFS(' + R + 'D2:D,A' + r + ',' + R + 'E2:E,"<>")',
      totalSlots]);
  });
  rows.push(['סה״כ', '=SUM(B2:B6)', '=SUM(C2:C6)', '=SUM(D2:D6)', totalSlots * HOUSES.length]);
  sum.getRange(1, 1, rows.length, 5).setValues(rows).setHorizontalAlignment('center').setFontSize(12);
  sum.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground(WINE).setFontColor('#ffffff');
  sum.getRange(rows.length, 1, 1, 5).setFontWeight('bold').setBackground('#f3e7da');
  HOUSES.forEach((h, i) => sum.getRange(i + 2, 1).setBackground(HOUSE_COLORS[i]).setFontWeight('bold'));
  [170, 90, 110, 150, 70].forEach((w, i) => sum.setColumnWidth(i + 1, w));
  sum.setFrozenRows(1); sum.setTabColor(WINE);
  sum.getRange('A9').setValue('כל לשונית בית מתעדכנת אוטומטית מכל הרשמה חדשה. את הנתונים הגולמיים לא לערוך ידנית (לשונית "הרשמות").')
     .setFontColor('#7a5c55').setFontStyle('italic');

  // --- לשונית לכל בית ---
  HOUSES.forEach((h, i) => {
    let sh = ss.getSheetByName(h.name) || ss.insertSheet(h.name);
    sh.clear(); sh.clearConditionalFormatRules(); sh.setRightToLeft(true); sh.setTabColor(HOUSE_COLORS[i]);
    const q = '"' + h.name + '"';
    sh.getRange('A1').setValue(h.name).setFontSize(18).setFontWeight('bold').setFontColor(WINE);
    sh.getRange('A2').setFormula('="נרשמו "&COUNTIF(' + R + 'D2:D,' + q + ')&" מתוך ' + CAPACITY + ' · נותרו "&(' + CAPACITY + '-COUNTIF(' + R + 'D2:D,' + q + '))&" מקומות"')
      .setFontSize(12).setFontColor('#5e1c26');
    sh.getRange('A1:E1').merge().setHorizontalAlignment('right'); sh.getRange('A2:E2').merge().setHorizontalAlignment('right');
    // רשימת נרשמות
    sh.getRange('A4:E4').setValues([['#', 'שם מלא', 'טלפון', 'מביאה', 'נרשמה ב']])
      .setFontWeight('bold').setBackground(WINE).setFontColor('#ffffff').setHorizontalAlignment('center');
    sh.getRange('B5').setFormula('=IFERROR(SORT(FILTER({' + R + 'B2:B,' + R + 'C2:C,' + R + 'E2:E,' + R + 'A2:A},' + R + 'D2:D=' + q + '),4,TRUE),"עדיין אין נרשמות")');
    sh.getRange('A5').setFormula('=ARRAYFORMULA(IF(B5:B19="","",IF(B5:B19="עדיין אין נרשמות","",ROW(B5:B19)-4)))');
    sh.getRange('C5:C19').setNumberFormat('@');
    sh.getRange('E5:E19').setNumberFormat('dd/MM HH:mm');
    sh.getRange('A5:E19').setBackground('#fffaf3').setBorder(true, true, true, true, true, true, '#e6d6c6', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange('A5:A19').setHorizontalAlignment('center').setFontColor('#7a5c55');
    // מי מביאה מה
    sh.getRange('G4:H4').setValues([['משבצת אוכל', 'מי מביאה']])
      .setFontWeight('bold').setBackground(WINE).setFontColor('#ffffff').setHorizontalAlignment('center');
    let r = 5;
    SLOTS.forEach(s => {
      const names = [s.name].concat(s.aliases || []);
      const cond = names.map(n => '(' + R + 'E2:E="' + n + '")').join('+');
      for (let k = 1; k <= s.count; k++) {
        const label = s.count > 1 ? s.name + ' ' + k : s.name;
        sh.getRange(r, 7).setValue(label);
        sh.getRange(r, 8).setFormula('=IFERROR(INDEX(FILTER(' + R + 'B2:B,(' + R + 'D2:D=' + q + ')*((' + cond + ')*1>0)),' + k + '),"— פנוי —")');
        r++;
      }
    });
    sh.getRange(5, 7, r - 5, 2).setBackground('#fffaf3').setBorder(true, true, true, true, true, true, '#e6d6c6', SpreadsheetApp.BorderStyle.SOLID);
    sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('— פנוי —').setFontColor('#b9a7a2').setItalic(true).setRanges([sh.getRange(5, 8, r - 5, 1)]).build()]);
    [40, 170, 120, 170, 110, 24, 170, 150].forEach((w, c) => sh.setColumnWidth(c + 1, w));
    sh.setFrozenRows(4);
  });

  // --- סדר הלשוניות: סיכום, 5 בתים, ובסוף הנתונים הגולמיים ---
  const order = ['סיכום'].concat(HOUSES.map(h => h.name)).concat([SHEET_NAME]);
  order.forEach((name, idx) => { const s = ss.getSheetByName(name); ss.setActiveSheet(s); ss.moveActiveSheet(idx + 1); });
  ss.setActiveSheet(ss.getSheetByName('סיכום'));
  Logger.log('VIEWS OK rows=' + (raw.getLastRow() - 1));
}
