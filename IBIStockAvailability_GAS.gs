// IBI Stock Availability — GAS Backend v1.3
// v1.3: CEO PIN moved server-side — `importNames` now requires p.pin === CEO_PIN,
//       and `verifyCeo` lets the app validate the PIN without ever storing it in
//       the page's JavaScript. REDEPLOY required.
// v1.2: added `importNames` bulk action — seed the sheet from a "List of Product
//       Names" file. mode 'replace' wipes all data rows then writes fresh
//       name-only rows (blank stock); mode 'add' appends only names not already
//       present. Sent via POST (form-urlencoded) so large lists fit. REDEPLOY
//       required for the app's Import Product Names feature to work.
// v1.1: the add duplicate-guard fingerprint now includes p.nonce, so the app's
//       Undo (restore of a just-deleted row) is never wrongly suppressed.
//       Optional redeploy — v1.7 app works with the v1.0 backend too.
// India Business International — Stock & Inventory Availability
// Sheet ID: 1bp0OpZJKWB3-xEDKAgk5EErQ4JVg3uwmF2QAM35K6lA  (first tab, gid=0)
// All requests via GET (URL params) — avoids CORS/redirect issues.
//
// Deploy:  Extensions → Apps Script → paste this file → Deploy → New deployment →
//          Web app → Execute as: Me → Who has access: Anyone → Deploy.
//          Copy the /exec URL and paste it into the app (Menu → Backend Connection).
//
// The sheet keeps its existing layout (two header rows, data from row 3):
//   1 S.No | 2 Category | 3 Product | 4 HSN | 5 GST | 6 Image | 7 Packed |
//   8 Loose | 9 Damage | 10 Date of Updation | 11 Star Rating | 12 Keywords |
//   13-17 Amazon rate 1-5 | 18 Amazon Avg | 19-23 Flipkart rate 1-5 | 24 Flipkart Avg
// A hidden ID column (25) is added on the right so rows have a stable identity
// (existing rows are auto-assigned an ID the first time the app loads). The visible
// layout A–X is untouched. Amazon/Flipkart averages are written as plain numbers
// (computed here) so no #DIV/0! errors appear.

const CEO_PIN    = "8899";  // CEO password — gates importNames server-side (not exposed in the app's JS)
const SHEET_ID   = "1bp0OpZJKWB3-xEDKAgk5EErQ4JVg3uwmF2QAM35K6lA";
const SHEET_GID  = 0;      // first tab
const DATA_START = 3;      // rows 1 & 2 are headers
const ID_COL     = 25;     // appended ID column (Y)
const LAST_COL   = 25;
const TZ         = "Asia/Kolkata";

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheets().filter(s => s.getSheetId() === SHEET_GID)[0] || ss.getSheets()[0];
  // Ensure the ID column has a header so it reads cleanly and is obvious in the sheet.
  if (String(sh.getRange(1, ID_COL).getValue()).trim() === '') {
    sh.getRange(1, ID_COL).setValue('ID');
  }
  return sh;
}

function doGet(e) {
  const p      = e.parameter || {};
  const action = p.action || '';
  let result;
  try {
    switch (action) {
      case 'ping':        result = { status:'ok', message:'IBI Stock Availability GAS v1.3 is live!' }; break;
      case 'getAll':      result = getAllItems(); break;
      case 'add':         result = addItem(p); break;
      case 'update':      result = updateItem(p); break;
      case 'delete':      result = deleteItem(p.id); break;
      case 'verifyCeo':   result = { status:'ok', valid: String(p.pin || '') === CEO_PIN }; break;
      case 'importNames': result = importNames(p); break;
      default:       result = { status:'error', message:'Unknown action: ' + action };
    }
  } catch(err) {
    result = { status:'error', message: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) { return doGet(e); }

// ── READ ──────────────────────────────────────────────────────────────────
function getAllItems() {
  const sh = getSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < DATA_START) return { status:'ok', items: [] };

  const n = lastRow - DATA_START + 1;
  const data = sh.getRange(DATA_START, 1, n, LAST_COL).getValues();

  // Assign IDs to any rows missing one, then persist in a single write.
  const base = Date.now();
  const idWrites = [];
  let changed = false;
  for (let i = 0; i < data.length; i++) {
    let id = String(data[i][ID_COL - 1] || '').trim();
    const hasProduct = String(data[i][2] || '').trim() !== '';   // a real stock item has a product name
    if (!id && hasProduct) { id = 'SK' + base + '_' + (DATA_START + i); changed = true; }
    data[i][ID_COL - 1] = id;
    idWrites.push([id]);
  }
  if (changed) sh.getRange(DATA_START, ID_COL, idWrites.length, 1).setValues(idWrites);

  const items = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (String(r[2] || '').trim() === '') continue;   // no product name → not a stock item, skip
    const upd = parseUpdated_(r[9]);
    items.push({
      id:          String(r[ID_COL - 1] || ''),
      sno:         r[0] === '' || r[0] == null ? '' : String(r[0]).replace(/\.0$/, ''),
      category:    String(r[1] || ''),
      product:     String(r[2] || ''),
      hsn:         r[3] === '' || r[3] == null ? '' : String(r[3]).replace(/\.0$/, ''),
      gst:         numOrBlank_(r[4]),
      image:       String(r[5] || ''),
      packed:      numOrBlank_(r[6]),
      loose:       numOrBlank_(r[7]),
      damage:      numOrBlank_(r[8]),
      date:        upd.iso,
      time:        upd.time,
      dateRaw:     upd.iso ? '' : String(r[9] || ''),
      rating:      numOrBlank_(r[10]),
      keywords:    String(r[11] || ''),
      amazon:      [r[12], r[13], r[14], r[15], r[16]].map(numOrBlank_),
      amazonAvg:   numOrBlank_(r[17]),
      flipkart:    [r[18], r[19], r[20], r[21], r[22]].map(numOrBlank_),
      flipkartAvg: numOrBlank_(r[23]),
      row:         DATA_START + i
    });
  }
  return { status:'ok', items: items };
}

// ── ADD ───────────────────────────────────────────────────────────────────
function addItem(p) {
  const sh = getSheet();

  // Duplicate-submit guard: a cold start / double-tap can fire the same add twice.
  // p.nonce (sent by the app's Undo-restore) is part of the fingerprint so that
  // re-adding a just-deleted row is never wrongly suppressed as a duplicate.
  const cache = CacheService.getScriptCache();
  const fp = ['add', p.category, p.product, p.hsn, p.packed, p.loose, p.damage, p.nonce]
             .map(x => String(x == null ? '' : x).trim()).join('|');
  const key = 'stk_' + Utilities.base64EncodeWebSafe(
              Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, fp));
  const seen = cache.get(key);
  if (seen) return { status:'ok', id: seen, duplicate:true, message:'Duplicate suppressed.' };

  const id  = 'SK' + Date.now();
  const row = buildRow_(p, id);
  // Append after the current last data row (keeps everything below the two headers).
  const target = Math.max(sh.getLastRow() + 1, DATA_START);
  sh.getRange(target, 1, 1, LAST_COL).setValues([row]);

  cache.put(key, id, 90);
  return { status:'ok', id: id, message:'Added.' };
}

// ── BULK IMPORT PRODUCT NAMES ───────────────────────────────────────────────
// Seed the sheet from a "List of Product Names" file (name-only rows, blank
// stock, to be filled in manually). p.names = JSON array of names.
//   mode 'replace' → wipe all existing data rows, then write the fresh list.
//   mode 'add'     → append only names not already present (case-insensitive).
function importNames(p) {
  // Server-side CEO gate: the write is refused unless the correct PIN is sent,
  // so revealing the importer UI in dev tools is not enough to upload.
  if (String(p.pin || '') !== CEO_PIN) {
    return { status:'error', code:'auth', message:'Incorrect CEO password.' };
  }
  var names;
  try { names = JSON.parse(p.names || '[]'); }
  catch (e) { return { status:'error', message:'Bad names payload: ' + e }; }
  if (!Array.isArray(names)) return { status:'error', message:'names must be a JSON array' };

  // Clean + dedupe case-insensitively, preserving order.
  var seen = {}, clean = [];
  for (var i = 0; i < names.length; i++) {
    var nm = String(names[i] == null ? '' : names[i]).trim();
    if (!nm) continue;
    var k = nm.toLowerCase();
    if (seen[k]) continue;
    seen[k] = true; clean.push(nm);
  }

  var mode    = (p.mode || 'add').toLowerCase();
  var sh      = getSheet();
  var lastRow = sh.getLastRow();
  var base    = Date.now();

  if (mode === 'replace') {
    var removed = Math.max(0, lastRow - DATA_START + 1);
    if (lastRow >= DATA_START) {
      sh.getRange(DATA_START, 1, removed, LAST_COL).clearContent();
    }
    var rows = clean.map(function (nm, i) { return nameRow_(nm, i + 1, 'SK' + base + '_' + i); });
    if (rows.length) sh.getRange(DATA_START, 1, rows.length, LAST_COL).setValues(rows);
    return { status:'ok', mode:'replace', added: rows.length, removed: removed, total: rows.length };
  }

  // mode 'add' — append names not already in the sheet.
  var existing = {}, maxSno = 0;
  if (lastRow >= DATA_START) {
    var cur = sh.getRange(DATA_START, 1, lastRow - DATA_START + 1, 3).getValues();  // S.No + Cat + Product
    for (var j = 0; j < cur.length; j++) {
      var pv = String(cur[j][2] || '').trim().toLowerCase();
      if (pv) existing[pv] = true;
      var sv = parseInt(cur[j][0], 10);
      if (!isNaN(sv) && sv > maxSno) maxSno = sv;
    }
  }
  var toAdd   = clean.filter(function (nm) { return !existing[nm.toLowerCase()]; });
  var newRows = toAdd.map(function (nm, i) { return nameRow_(nm, maxSno + 1 + i, 'SK' + base + '_' + i); });
  if (newRows.length) {
    var target = Math.max(sh.getLastRow() + 1, DATA_START);
    sh.getRange(target, 1, newRows.length, LAST_COL).setValues(newRows);
  }
  return { status:'ok', mode:'add', added: newRows.length, skipped: clean.length - newRows.length, total: clean.length };
}

// A product-name-only row: only S.No, Product and ID set; everything else blank
// (no date stamp — these are catalog seeds to be filled in manually later).
function nameRow_(name, sno, id) {
  var row = [];
  for (var i = 0; i < LAST_COL; i++) row.push('');
  row[0]           = sno;    // 1  S.No
  row[2]           = name;   // 3  Product
  row[LAST_COL - 1] = id;    // 25 ID (col Y)
  return row;
}

// ── UPDATE ─────────────────────────────────────────────────────────────────
function updateItem(p) {
  if (!p.id) return { status:'error', message:'No ID provided.' };
  const sh = getSheet();
  const r  = findRowById_(sh, p.id);
  if (r < 0) return { status:'error', message:'ID not found: ' + p.id };
  const row = buildRow_(p, p.id);
  sh.getRange(r, 1, 1, LAST_COL).setValues([row]);
  return { status:'ok', message:'Updated: ' + p.id };
}

// ── DELETE ─────────────────────────────────────────────────────────────────
function deleteItem(id) {
  if (!id) return { status:'error', message:'No ID provided.' };
  const sh = getSheet();
  const r  = findRowById_(sh, id);
  if (r < 0) return { status:'error', message:'ID not found: ' + id };
  sh.deleteRow(r);
  return { status:'ok', message:'Deleted: ' + id };
}

// ── HELPERS ────────────────────────────────────────────────────────────────
function findRowById_(sh, id) {
  const lastRow = sh.getLastRow();
  if (lastRow < DATA_START) return -1;
  const ids = sh.getRange(DATA_START, ID_COL, lastRow - DATA_START + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) return DATA_START + i;
  }
  return -1;
}

// Build the full 25-column row from posted params.
function buildRow_(p, id) {
  const amz = [p.a1, p.a2, p.a3, p.a4, p.a5];
  const flp = [p.f1, p.f2, p.f3, p.f4, p.f5];
  const dateVal = fmtStamp_(new Date());   // auto "last updated" — current IST date + time, every add/update
  return [
    numOr_(p.sno, ''),         // 1  S.No
    p.category || '',          // 2  Category
    p.product  || '',          // 3  Product
    p.hsn || '',               // 4  HSN
    numOr_(p.gst, ''),         // 5  GST
    p.image || '',             // 6  Image
    numOr_(p.packed, ''),      // 7  Packed
    numOr_(p.loose, ''),       // 8  Loose
    numOr_(p.damage, ''),      // 9  Damage
    dateVal,                   // 10 Date of Updation
    numOr_(p.rating, ''),      // 11 Star Rating
    p.keywords || '',          // 12 Keywords
    numOr_(amz[0], ''), numOr_(amz[1], ''), numOr_(amz[2], ''), numOr_(amz[3], ''), numOr_(amz[4], ''), // 13-17
    avg_(amz),                 // 18 Amazon Avg
    numOr_(flp[0], ''), numOr_(flp[1], ''), numOr_(flp[2], ''), numOr_(flp[3], ''), numOr_(flp[4], ''), // 19-23
    avg_(flp),                 // 24 Flipkart Avg
    id                         // 25 ID
  ];
}

function numOr_(v, dflt) {
  if (v === '' || v == null) return dflt;
  const n = parseFloat(v);
  return isNaN(n) ? dflt : n;
}
function numOrBlank_(v) {
  if (v === '' || v == null) return '';
  const n = parseFloat(v);
  return isNaN(n) ? '' : n;
}
function avg_(arr) {
  const nums = arr.map(x => parseFloat(x)).filter(x => !isNaN(x) && x > 0);
  if (!nums.length) return '';
  return Math.round((nums.reduce((s, x) => s + x, 0) / nums.length) * 100) / 100;
}

function isoToDate_(iso) {
  const m = String(iso).match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], 12, 0, 0); // noon → no TZ off-by-one
  const d = new Date(iso);
  return isNaN(d.getTime()) ? new Date() : d;
}
function fmtLong_(d) {
  // Matches the user's existing style, e.g. "Tuesday, 7 July 2026".
  return Utilities.formatDate(d, TZ, "EEEE, d MMMM yyyy");
}

function fmtStamp_(d) {
  // Auto "last updated" stamp, e.g. "28 May 2026, Thu, 01:38:00 PM"
  return Utilities.formatDate(d, TZ, "dd MMM yyyy, EEE, hh:mm:ss a");
}

// Parse a "Date of Updation" cell into { iso, time }. Handles the new auto-stamp
// (with time), legacy Date objects, and older date-only text formats.
function parseUpdated_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return { iso: Utilities.formatDate(v, TZ, "yyyy-MM-dd"), time: "" };
  }
  var s = String(v || '').trim();
  if (!s) return { iso: '', time: '' };
  var m = s.match(/^(\d{1,2})\s+([A-Za-z]{3,})\s+(\d{4}),\s*[A-Za-z]{3,},\s*(\d{1,2}:\d{2}:\d{2}\s*[AaPp][Mm])$/);
  if (m) {
    var mo = _MONTHS[m[2].slice(0, 3).toLowerCase()];
    if (mo) return {
      iso:  m[3] + '-' + ('0' + mo).slice(-2) + '-' + ('0' + (+m[1])).slice(-2),
      time: m[4].toUpperCase()
    };
  }
  return { iso: parseDateToISO_(s), time: '' };
}

const _MONTHS = { jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12 };
function parseDateToISO_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
  let s = String(v || '').trim();
  if (!s) return '';
  const pad = n => ('0' + n).slice(-2);
  const iso = (y, mo, d) => y + '-' + pad(mo) + '-' + pad(d);

  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);              // yyyy-mm-dd
  if (m) return iso(m[1], +m[2], +m[3]);

  m = s.match(/^(\d{1,2})\s*[\/\-.]\s*(\d{1,2})\s*[\/\-.]\s*(\d{4})$/);  // dd/mm/yyyy (Indian, day-first; tolerates spaces like "26 - 5 - 2026")
  if (m) {
    let a = +m[1], b = +m[2]; const y = m[3];
    let day, mon;
    if (a > 12 && b <= 12)      { day = a; mon = b; }
    else if (b > 12 && a <= 12) { day = b; mon = a; }
    else                        { day = a; mon = b; }
    return iso(y, mon, day);
  }

  m = s.match(/^(\d{1,2})[\-\s]([A-Za-z]{3,})[\-\s](\d{4})$/);   // 7-Jul-2026 / 7 July 2026
  if (m) { const mo = _MONTHS[m[2].slice(0, 3).toLowerCase()]; if (mo) return iso(m[3], mo, +m[1]); }

  const s2 = s.replace(/^[A-Za-z]+,\s*/, '');                    // drop "Monday, " weekday prefix
  m = s2.match(/^(\d{1,2})\s+([A-Za-z]{3,})\s+(\d{4})$/);        // "2 March 2026"
  if (m) { const mo = _MONTHS[m[2].slice(0, 3).toLowerCase()]; if (mo) return iso(m[3], mo, +m[1]); }

  const d = new Date(s2);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, TZ, "yyyy-MM-dd");
  return '';
}
