/**
 * Sneha's Makeover - Booking backend (Google Apps Script web app)
 *
 * Sheets used (created by setup()):
 *   Bookings      : ID | Name | Phone | Date | Service | Location | Message | Status | Created At
 *   BlockedDates  : Date | Reason
 *
 * Admin password lives in Project Settings > Script properties > ADMIN_PASSWORD
 * (never write the password inside this file).
 */

const SHEET_BOOKINGS = 'Bookings';
const SHEET_BLOCKED = 'BlockedDates';
const HEADERS = ['ID', 'Name', 'Phone', 'Date', 'Service', 'Location', 'Message', 'Status', 'Created At'];
const BLOCKED_HEADERS = ['Date', 'Reason'];
const STATUSES = ['Pending', 'Confirmed', 'Cancelled'];
const SERVICES = ['Bridal HD Makeup', 'Reception / Engagement', 'Party Makeup', 'Saree Draping & Hair'];
const TZ = 'Asia/Kolkata';
const MAX_CONFIRMED_PER_DAY = 1;   // how many confirmed bookings Sneha can take on one date
const TOKEN_TTL = 21600;           // 6 hours (CacheService maximum)

/* ---------- one-time setup: run this once from the editor ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_BOOKINGS, HEADERS);
  ensureSheet_(ss, SHEET_BLOCKED, BLOCKED_HEADERS);
  // keep dates / phones as plain text so Sheets never converts them
  ss.getSheetByName(SHEET_BOOKINGS).getRange('C:D').setNumberFormat('@');
  ss.getSheetByName(SHEET_BLOCKED).getRange('A:A').setNumberFormat('@');
  if (!PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD')) {
    Logger.log('Now add ADMIN_PASSWORD in Project Settings > Script properties.');
  }
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- entry points ---------- */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'dates') return json_({ ok: true, unavailable: unavailableDates_() });
    return json_({ ok: true, service: "Sneha's Makeover booking API" });
  } catch (err) {
    return json_({ ok: false, error: 'server_error' });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (body.action) {
      case 'book':      return json_(createBooking_(body));
      case 'login':     return json_(login_(body));
      case 'list':      requireAuth_(body); return json_({ ok: true, bookings: listBookings_() });
      case 'setStatus': requireAuth_(body); return json_(setStatus_(body));
      case 'blocked':   requireAuth_(body); return json_({ ok: true, blocked: listBlocked_() });
      case 'block':     requireAuth_(body); return json_(blockDate_(body));
      case 'unblock':   requireAuth_(body); return json_(unblockDate_(body));
      default:          return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    if (msg === 'unauthorized') return json_({ ok: false, error: 'unauthorized' });
    return json_({ ok: false, error: 'server_error' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- helpers ---------- */
function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Run setup() first');
  return sh;
}

function ymd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}

function today_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function isValidDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const d = new Date(s + 'T00:00:00Z');
  return !isNaN(d.getTime()) && d.toISOString().slice(0, 10) === s;
}

// trim, limit length, and stop spreadsheet formula injection (=, +, -, @ at the start)
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function cleanPhone_(v) {
  let d = String(v == null ? '' : v).replace(/\D/g, '');
  if (d.length === 12 && d.indexOf('91') === 0) d = d.slice(2);
  if (d.length === 11 && d.charAt(0) === '0') d = d.slice(1);
  return /^[6-9]\d{9}$/.test(d) ? d : '';
}

function listBlocked_() {
  const rows = sheet_(SHEET_BLOCKED).getDataRange().getValues().slice(1);
  return rows
    .filter(function (r) { return ymd_(r[0]); })
    .map(function (r) { return { date: ymd_(r[0]), reason: String(r[1] || '') }; })
    .sort(function (a, b) { return a.date < b.date ? -1 : 1; });
}

function listBookings_() {
  const rows = sheet_(SHEET_BOOKINGS).getDataRange().getValues().slice(1);
  return rows
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        id: String(r[0]), name: String(r[1]), phone: String(r[2]), date: ymd_(r[3]),
        service: String(r[4]), location: String(r[5]), message: String(r[6]),
        status: String(r[7]), createdAt: String(r[8])
      };
    })
    .sort(function (a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
}

// dates customers cannot pick: blocked by Sneha, or already full with confirmed bookings
function unavailableDates_() {
  const out = {};
  listBlocked_().forEach(function (b) { out[b.date] = true; });
  const counts = {};
  listBookings_().forEach(function (b) {
    if (b.status === 'Confirmed') counts[b.date] = (counts[b.date] || 0) + 1;
  });
  Object.keys(counts).forEach(function (d) { if (counts[d] >= MAX_CONFIRMED_PER_DAY) out[d] = true; });
  return Object.keys(out).filter(function (d) { return d >= today_(); }).sort();
}

/* ---------- customer: create booking ---------- */
function createBooking_(d) {
  if (d.website) return { ok: true, id: 'SM-0000' };   // honeypot: bots fill this field

  const name = clean_(d.name, 80);
  const phone = cleanPhone_(d.phone);
  const date = String(d.date || '').trim();
  const service = String(d.service || '').trim();
  const location = clean_(d.location, 120);
  const message = clean_(d.message, 500);

  if (!name) return { ok: false, error: 'invalid_name' };
  if (!phone) return { ok: false, error: 'invalid_phone' };
  if (!isValidDate_(date) || date < today_()) return { ok: false, error: 'invalid_date' };
  if (SERVICES.indexOf(service) === -1) return { ok: false, error: 'invalid_service' };

  // simple spam brake: max 5 requests per phone per hour
  const cache = CacheService.getScriptCache();
  const rateKey = 'rate_' + phone;
  const used = Number(cache.get(rateKey) || 0);
  if (used >= 5) return { ok: false, error: 'too_many_requests' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (unavailableDates_().indexOf(date) !== -1) return { ok: false, error: 'date_unavailable' };
    const sh = sheet_(SHEET_BOOKINGS);
    const id = 'SM-' + ('0000' + sh.getLastRow()).slice(-4);   // first booking -> SM-0001
    const created = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    sh.appendRow([id, name, phone, date, service, location, message, 'Pending', created]);
    cache.put(rateKey, String(used + 1), 3600);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- owner: login + token ---------- */
function login_(d) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('login_fails') || 0);
  if (fails >= 5) return { ok: false, error: 'locked' };           // 15 min cool-down

  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real) return { ok: false, error: 'password_not_set' };

  if (String(d.password || '') !== real) {
    cache.put('login_fails', String(fails + 1), 900);
    return { ok: false, error: 'wrong_password' };
  }
  cache.remove('login_fails');
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('tok_' + token, '1', TOKEN_TTL);
  return { ok: true, token: token };
}

function requireAuth_(d) {
  const token = String(d.token || '');
  if (!token || !CacheService.getScriptCache().get('tok_' + token)) throw new Error('unauthorized');
}

/* ---------- owner: manage bookings and dates ---------- */
function setStatus_(d) {
  const id = String(d.id || '');
  const status = String(d.status || '');
  if (STATUSES.indexOf(status) === -1) return { ok: false, error: 'invalid_status' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(SHEET_BOOKINGS);
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== id) continue;
      if (status === 'Confirmed') {
        const date = ymd_(rows[i][3]);
        const others = rows.filter(function (r, j) {
          return j > 0 && j !== i && String(r[7]) === 'Confirmed' && ymd_(r[3]) === date;
        }).length;
        if (others >= MAX_CONFIRMED_PER_DAY) return { ok: false, error: 'date_full' };
      }
      sh.getRange(i + 1, 8).setValue(status);
      return { ok: true };
    }
    return { ok: false, error: 'not_found' };
  } finally {
    lock.releaseLock();
  }
}

function blockDate_(d) {
  const date = String(d.date || '').trim();
  if (!isValidDate_(date)) return { ok: false, error: 'invalid_date' };
  const sh = sheet_(SHEET_BLOCKED);
  const exists = listBlocked_().some(function (b) { return b.date === date; });
  if (!exists) sh.appendRow([date, clean_(d.reason, 100)]);
  return { ok: true };
}

function unblockDate_(d) {
  const date = String(d.date || '').trim();
  const sh = sheet_(SHEET_BLOCKED);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (ymd_(rows[i][0]) === date) sh.deleteRow(i + 1);
  }
  return { ok: true };
}
