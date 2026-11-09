/**
 * ערב נשים בבתים – צד שרת להרשמה (Google Apps Script)
 * מחובר לגיליון "הרשמות – ערב נשים בבתים".
 *
 * GET  -> מחזיר מקומות פנויים לכל בית ומשבצות אוכל פנויות
 * POST -> רושם משתתפת (גוף JSON כ-text/plain): {name, phone, house, slot}
 *         או {action: lookup|update|cancel, phone, name, id, house, slot} לשינוי עצמי
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

function doGet(e) {
  if (e && e.parameter && e.parameter.admin) return adminPage_();   // ?admin=1 – דף ניהול (רק למנהלות)
  return json_(publicState_(readState_()));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return json_({ ok: false, error: 'busy' });
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action) return json_(selfAction_(body));   // lookup / update / cancel
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

/* ===== שינוי וביטול הרשמה (נרשמות ומנהלות) ===== */
const ADMINS = ['netanelm87@gmail.com', 'moyalnh@gmail.com'];
const CANCEL_SHEET = 'בוטלו';

function phoneKey_(p) {            // 0541234567 == 972541234567 == +972-54-123-4567
  let d = String(p || '').replace(/\D/g, '');
  if (d.indexOf('972') === 0) d = '0' + d.slice(3);
  if (d && d[0] !== '0') d = '0' + d;
  return d;
}
function firstName_(n) {
  return String(n || '').trim().split(/\s+/)[0].replace(/[^\u0590-\u05FFa-zA-Z]/g, '').toLowerCase();
}
function slotByName_(name) {
  return SLOTS.find(s => s.name === name || (s.aliases || []).indexOf(name) >= 0) || null;
}
function regId_(row) {
  const t = row[0] instanceof Date ? row[0].getTime() : String(row[0]);
  return t + '_' + phoneKey_(row[2]);
}
function allRegs_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  const vals = last > 1 ? sh.getRange(2, 1, last - 1, HEADERS.length).getValues() : [];
  return vals.map((r, i) => {
    const h = HOUSES.find(x => x.name === r[3]) || {};
    const s = slotByName_(r[4]);
    return { id: regId_(r), row: i + 2, name: String(r[1]), phone: phoneKey_(r[2]),
             house: h.id || '', houseName: String(r[3]), slot: s ? s.id : '', slotName: String(r[4] || ''),
             time: r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Jerusalem', 'dd/MM HH:mm') : '' };
  });
}

// שינוי בית ו/או משבצת לשורה קיימת, עם בדיקות מקום
function changeReg_(id, houseId, slotId, force) {
  const regs = allRegs_();
  const me = regs.find(r => r.id === id);
  if (!me) return { ok: false, error: 'not_found' };
  const house = HOUSES.find(h => h.id === String(houseId));
  if (!house) return { ok: false, error: 'bad_house' };
  const slotDef = slotId ? SLOTS.find(s => s.id === String(slotId)) : null;
  if (slotId && !slotDef) return { ok: false, error: 'bad_slot' };
  const others = regs.filter(r => r.id !== id && r.house === house.id);
  if (!force && others.length >= CAPACITY) return { ok: false, error: 'house_full' };
  if (others.some(r => r.phone === me.phone)) return { ok: false, error: 'already_registered' };
  if (slotDef && others.filter(r => r.slot === slotDef.id).length >= slotDef.count) return { ok: false, error: 'slot_taken' };
  getSheet_().getRange(me.row, 4, 1, 2).setValues([[house.name, slotDef ? slotDef.name : '']]);
  return { ok: true };
}

function cancelReg_(id, by) {
  const regs = allRegs_();
  const me = regs.find(r => r.id === id);
  if (!me) return { ok: false, error: 'not_found' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName(CANCEL_SHEET);
  if (!log) {
    log = ss.insertSheet(CANCEL_SHEET);
    log.setRightToLeft(true);
    log.appendRow(HEADERS.concat(['בוטל ב', 'בוטל ע״י']));
    log.getRange(1, 1, 1, HEADERS.length + 2).setFontWeight('bold').setBackground('#7a5c55').setFontColor('#ffffff');
    log.setFrozenRows(1); log.setTabColor('#b9a7a2');
  }
  const sh = getSheet_();
  const row = sh.getRange(me.row, 1, 1, HEADERS.length).getValues()[0];
  log.appendRow([row[0], row[1], "'" + String(row[2]).replace(/^'/, ''), row[3], row[4], new Date(), by]);
  log.getRange(log.getLastRow(), 1).setNumberFormat('dd/MM HH:mm');
  log.getRange(log.getLastRow(), HEADERS.length + 1).setNumberFormat('dd/MM HH:mm');
  sh.deleteRow(me.row);
  return { ok: true };
}

// --- נרשמת: זיהוי לפי טלפון + שם פרטי ---
function selfAction_(body) {
  const phone = phoneKey_(body.phone), first = firstName_(body.name);
  if (phone.length < 9 || !first) return { ok: false, error: 'not_found' };
  const mine = allRegs_().filter(r => r.phone === phone && firstName_(r.name) === first);
  if (body.action === 'lookup') {
    if (!mine.length) return { ok: false, error: 'not_found' };
    return { ok: true, regs: mine.map(r => ({ id: r.id, name: r.name, house: r.house, slot: r.slot })), state: publicState_(readState_()) };
  }
  const target = mine.find(r => r.id === String(body.id));
  if (!target) return { ok: false, error: 'not_found' };
  const res = body.action === 'cancel' ? cancelReg_(target.id, 'הנרשמת')
            : body.action === 'update' ? changeReg_(target.id, body.house, body.slot || '', false)
            : { ok: false, error: 'bad_action' };
  res.state = publicState_(readState_());
  return res;
}

// --- מנהלות: תפריט בגיליון ---
function isAdmin_() {
  const e = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  return ADMINS.indexOf(e) >= 0;
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🛠 ניהול הרשמות').addItem('פתיחת חלון ניהול', 'openAdmin').addToUi();
}
function openAdmin() {
  if (!isAdmin_()) { SpreadsheetApp.getUi().alert('אין לך הרשאת ניהול. פני לנתנאל.'); return; }
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(ADMIN_HTML).setTitle('ניהול הרשמות'));
}
function adminPage_() {
  const meta = o => o.setTitle('ניהול הרשמות – ערב נשים').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (!isAdmin_()) {
    const who = Session.getActiveUser().getEmail() || 'לא מחוברת';
    return meta(HtmlService.createHtmlOutput('<div dir="rtl" style="font-family:Arial;padding:30px;text-align:center"><h2 style="color:#8c2f3a">אין הרשאת ניהול</h2><p>מחוברת כעת: ' + who + '</p><p>יש להתחבר עם חשבון הגוגל המורשה.</p></div>'));
  }
  return meta(HtmlService.createHtmlOutput(ADMIN_HTML.replace('<body>', '<body class="page">')));
}
function adminData() {
  if (!isAdmin_()) throw new Error('אין הרשאה');
  const st = publicState_(readState_());
  return { regs: allRegs_(), houses: st.houses, slots: SLOTS.map(s => ({ id: s.id, name: s.name, count: s.count })), capacity: CAPACITY };
}
function adminChange(id, houseId, slotId, force) {
  if (!isAdmin_()) throw new Error('אין הרשאה');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try { return changeReg_(id, houseId, slotId, !!force); } finally { lock.releaseLock(); }
}
function adminCancel(id) {
  if (!isAdmin_()) throw new Error('אין הרשאה');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try { return cancelReg_(id, Session.getActiveUser().getEmail()); } finally { lock.releaseLock(); }
}

const ADMIN_HTML = `<!doctype html><html dir="rtl"><head><base target="_top"><meta charset="utf-8">
<style>
body{font-family:Arial,sans-serif;margin:0;padding:10px;color:#3a1f1f;font-size:13px;background:#fffaf3}
input,select,button{font:inherit}
#q{width:100%;box-sizing:border-box;padding:8px;border:1px solid #d9c6b4;border-radius:8px;margin-bottom:8px}
.sum{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:8px}
.sum span{background:#f3e7da;border-radius:99px;padding:2px 8px;font-size:12px}
h3{margin:12px 0 4px;color:#8c2f3a;font-size:14px;border-bottom:2px solid #8c2f3a;padding-bottom:2px}
.r{background:#fff;border:1px solid #eadbcc;border-radius:8px;padding:8px;margin:6px 0}
.n{font-weight:bold}.m{color:#7a5c55;font-size:12px}
.ed{display:none;margin-top:6px}.r.open .ed{display:block}
.ed select{width:100%;margin:3px 0;padding:4px}
.bt{display:flex;gap:6px;margin-top:4px}
.bt button{flex:1;border:0;border-radius:6px;padding:6px;cursor:pointer;color:#fff;background:#8c2f3a}
.bt .x{background:#b9a7a2}.bt .c{background:#fff;color:#8c2f3a;border:1px solid #8c2f3a}
.page{max-width:640px;margin:0 auto;padding:16px;font-size:15px}
.page:before{content:'ניהול הרשמות – ערב נשים';display:block;font-size:22px;font-weight:bold;color:#8c2f3a;margin-bottom:10px}
.msg{position:sticky;bottom:0;background:#3a1f1f;color:#fff;padding:6px;border-radius:6px;display:none;margin-top:8px}
</style></head><body>
<input id="q" placeholder="חיפוש לפי שם או טלפון…" oninput="draw()">
<div class="sum" id="sum"></div><div id="list">טוען…</div><div class="msg" id="msg"></div>
<script>
let D=null;
const ERR={house_full:'הבית מלא',slot_taken:'המשבצת תפוסה בבית הזה',already_registered:'הטלפון הזה כבר רשום בבית הזה',not_found:'ההרשמה לא נמצאה (אולי השתנתה), מרעננת'};
function say(t){const m=document.getElementById('msg');m.textContent=t;m.style.display='block';clearTimeout(say.t);say.t=setTimeout(()=>m.style.display='none',3500)}
function load(){google.script.run.withSuccessHandler(d=>{D=d;draw()}).withFailureHandler(e=>say(e.message)).adminData()}
function esc(s){return String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
function draw(){
  if(!D)return; const q=document.getElementById('q').value.trim();
  document.getElementById('sum').innerHTML=D.houses.map(h=>'<span>'+esc(h.name)+': '+(D.capacity-h.remaining)+'/'+D.capacity+'</span>').join('');
  let html='';
  D.houses.forEach(h=>{
    const rs=D.regs.filter(r=>r.house===h.id&&(!q||r.name.includes(q)||r.phone.includes(q.replace(/\\D/g,''))));
    if(!rs.length&&q)return;
    html+='<h3>'+esc(h.name)+' ('+(D.capacity-h.remaining)+'/'+D.capacity+')</h3>';
    rs.forEach(r=>{
      html+='<div class="r" id="r_'+r.id+'"><div onclick="tog(\\''+r.id+'\\')" style="cursor:pointer"><div class="n">'+esc(r.name)+'</div><div class="m">'+r.phone+' · '+(r.slotName?'מביאה '+esc(r.slotName):'לא מביאה')+' · '+r.time+'</div></div>'+
      '<div class="ed"><select id="h_'+r.id+'">'+D.houses.map(x=>'<option value="'+x.id+'"'+(x.id===r.house?' selected':'')+'>'+esc(x.name)+' (נותרו '+x.remaining+')</option>').join('')+'</select>'+
      '<select id="s_'+r.id+'"><option value="">ללא משבצת אוכל</option>'+D.slots.map(s=>'<option value="'+s.id+'"'+(s.id===r.slot?' selected':'')+'>'+esc(s.name)+'</option>').join('')+'</select>'+
      '<div class="bt"><button onclick="save(\\''+r.id+'\\')">שמירה</button><button class="x" onclick="cxl(\\''+r.id+'\\',\\''+esc(r.name).replace(/'/g,'')+'\\')">ביטול רישום</button></div></div></div>';
    });
  });
  document.getElementById('list').innerHTML=html||'לא נמצאו נרשמות';
}
function tog(id){document.getElementById('r_'+id).classList.toggle('open')}
function done(res,okText,retry){
  if(res.ok){say(okText);load();return}
  if(res.error==='house_full'&&retry){ if(confirm('הבית מלא (15). להעביר בכל זאת מעבר למכסה?')) retry(); return }
  say(ERR[res.error]||'שגיאה'); if(res.error==='not_found')load();
}
function save(id){
  const h=document.getElementById('h_'+id).value,s=document.getElementById('s_'+id).value;
  const run=f=>google.script.run.withSuccessHandler(r=>done(r,'נשמר ✓',f?null:()=>run(true))).withFailureHandler(e=>say(e.message)).adminChange(id,h,s,f);
  run(false);
}
function cxl(id,name){ if(!confirm('לבטל את הרישום של '+name+'? (השורה תעבור ללשונית "בוטלו")'))return;
  google.script.run.withSuccessHandler(r=>done(r,'הרישום בוטל')).withFailureHandler(e=>say(e.message)).adminCancel(id) }
load();
</script></body></html>`;
