/**
 * Goolee.my Attendance — Apps Script backend v2.0
 * ─────────────────────────────────────────────────
 * Bound to a Google Sheet with two tabs: Employees, Attendance.
 *
 * Deployment:
 *   Extensions → Apps Script → paste this → Save
 *   Deploy → New deployment → Web app
 *     Execute as:      Me
 *     Who has access:  Anyone      ← must be "Anyone"
 *   Copy the /exec URL into index.html → CONFIG.API_URL
 *
 * After every edit:
 *   Deploy → Manage deployments → pencil ✏️ → Version: New version → Deploy
 */

const EMPLOYEES_TAB       = 'Employees';
const ATTENDANCE_TAB      = 'Attendance';
const ACCURACY_MAX_M      = 50;
const PIN_LOCKOUT_TRIES   = 5;
const PIN_LOCKOUT_MINUTES = 5;

// ═══════════════════════════════════════════════════════════════════════
// HTTP entry points
// ═══════════════════════════════════════════════════════════════════════
function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action;

  if (action === 'ping') {
    return respond_(p.callback, { ok: true, message: 'Goolee attendance API is live.' });
  }
  if (action === 'employees') {
    return respond_(p.callback, { ok: true, employees: listEmployees_() });
  }
  if (action === 'punch') {
    return respond_(p.callback, handlePunch_(p));
  }
  return respond_(p.callback, { ok: false, message: 'Unknown action: ' + action });
}

function doPost(e) {
  // Kept for compatibility — the frontend uses GET/JSONP.
  const p = (e && e.parameter) || {};
  const body = p.action ? p : safeJson_(e.postData && e.postData.contents);
  if (!body) return respond_(null, { ok: false, message: 'Malformed request.' });
  if (body.action === 'punch') return respond_(null, handlePunch_(body));
  return respond_(null, { ok: false, message: 'Unknown action.' });
}

// ═══════════════════════════════════════════════════════════════════════
// PIN normalisation — fixes the leading-zero problem
// ═══════════════════════════════════════════════════════════════════════
function normalizePin_(raw) {
  if (raw === null || raw === undefined) return '';

  let s;
  if (typeof raw === 'number') {
    // Sheet stored "0093" as the number 93 — strip decimals
    s = String(Math.round(raw));
  } else {
    s = String(raw).trim();
  }

  // Pad 1–4 digit PINs to 4 characters so "93" becomes "0093"
  if (/^\d{1,4}$/.test(s)) return s.padStart(4, '0');
  return s;
}

// ═══════════════════════════════════════════════════════════════════════
// Sheet access
// ═══════════════════════════════════════════════════════════════════════
function sheet_(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('Missing sheet tab: ' + name);
  return s;
}

function listEmployees_() {
  const values = sheet_(EMPLOYEES_TAB).getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(h => String(h).trim().toLowerCase());
  const nameIdx   = header.indexOf('name');
  const activeIdx = header.indexOf('active');
  if (nameIdx < 0) return [];

  const out = [];
  for (let i = 1; i < values.length; i++) {
    const name   = values[i][nameIdx];
    const active = activeIdx >= 0 ? values[i][activeIdx] : true;
    if (!name) continue;
    if (active === false) continue;
    if (String(active).trim().toUpperCase() === 'FALSE') continue;
    out.push({ name: String(name).trim() });
  }
  return out;
}

function findEmployee_(name) {
  const values = sheet_(EMPLOYEES_TAB).getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0].map(h => String(h).trim().toLowerCase());
  const nameIdx   = header.indexOf('name');
  const pinIdx    = header.indexOf('pin');
  const activeIdx = header.indexOf('active');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][nameIdx]).trim() !== name) continue;
    return {
      name: String(values[i][nameIdx]).trim(),
      pin:  pinIdx >= 0 ? normalizePin_(values[i][pinIdx]) : '',
      active: activeIdx < 0 ? true
              : !(values[i][activeIdx] === false ||
                  String(values[i][activeIdx]).trim().toUpperCase() === 'FALSE')
    };
  }
  return null;
}

// ═══════════════════════════════════════════════════════════════════════
// Lockout — server-side, per employee
// ═══════════════════════════════════════════════════════════════════════
function lockKey_(name) { return 'lockout_' + name; }

function isLocked_(name) {
  const raw = PropertiesService.getScriptProperties().getProperty(lockKey_(name));
  if (!raw) return false;
  try { return JSON.parse(raw).until > Date.now(); }
  catch (e) { return false; }
}

function recordFailure_(name) {
  const props = PropertiesService.getScriptProperties();
  const raw   = props.getProperty(lockKey_(name));
  const state = raw ? JSON.parse(raw) : { tries: 0, until: 0 };

  state.tries = (state.tries || 0) + 1;
  if (state.tries >= PIN_LOCKOUT_TRIES) {
    state.until = Date.now() + PIN_LOCKOUT_MINUTES * 60 * 1000;
    state.tries = 0;
  }
  props.setProperty(lockKey_(name), JSON.stringify(state));
}

function clearFailure_(name) {
  PropertiesService.getScriptProperties().deleteProperty(lockKey_(name));
}

// ═══════════════════════════════════════════════════════════════════════
// Punch handler
// ═══════════════════════════════════════════════════════════════════════
function handlePunch_(body) {
  const name      = String(body.name || '').trim();
  const pin       = normalizePin_(body.pin);
  const pinHash   = String(body.pinHash || '').trim().toLowerCase();
  const punchType = body.punchType === 'Clock Out' ? 'Clock Out' : 'Clock In';
  const lat       = Number(body.lat);
  const lng       = Number(body.lng);
  const accuracy  = Number(body.accuracy);

  if (!name) {
    return { ok: false, message: 'Missing employee name.' };
  }

  if (isLocked_(name)) {
    return {
      ok: false,
      reason: 'locked_out',
      message: 'Too many incorrect PIN attempts. Try again in a few minutes.'
    };
  }

  // Mandatory GPS — accuracy ≤ 50 m
  if (isNaN(lat) || isNaN(lng) || isNaN(accuracy) || accuracy > ACCURACY_MAX_M) {
    return {
      ok: false,
      reason: 'bad_gps',
      message: 'Location was missing or not accurate enough. Move outdoors and try again.'
    };
  }

  const employee = findEmployee_(name);
  if (!employee || !employee.active) {
    return {
      ok: false,
      reason: 'unknown_employee',
      message: 'Employee not found or inactive.'
    };
  }

  // Accept plain PIN or its SHA-256 hash
  const pinOk     = /^\d{4}$/.test(pin) && pin === employee.pin;
  const pinHashOk = /^[a-f0-9]{64}$/.test(pinHash) && pinHash === sha256_(employee.pin);

  if (!pinOk && !pinHashOk) {
    recordFailure_(name);
    return { ok: false, reason: 'bad_pin', message: 'Incorrect PIN.' };
  }

  clearFailure_(name);

  const now = new Date();
  const tz  = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
  const mapsLink = 'https://www.google.com/maps?q=' + lat + ',' + lng;

  const sheet = sheet_(ATTENDANCE_TAB);
  const row   = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 7).setValues([[
    now,
    employee.name,
    punchType,
    '=HYPERLINK("' + mapsLink + '","Open in Maps")',
    accuracy,
    dateStr,
    timeStr
  ]]);

  return {
    ok: true,
    name: employee.name,
    punchType: punchType,
    date: dateStr,
    time: timeStr,
    mapsLink: mapsLink,
    accuracy: accuracy
  };
}

// ═══════════════════════════════════════════════════════════════════════
// Helpers
// ═══════════════════════════════════════════════════════════════════════
function respond_(callback, obj) {
  const json = JSON.stringify(obj);
  if (callback) {
    const safe = String(callback).replace(/[^\w$.]/g, '');
    return ContentService
      .createTextOutput(safe + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function safeJson_(text) {
  try { return JSON.parse(text); } catch (e) { return null; }
}

function sha256_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function (b) {
    const u = b < 0 ? b + 256 : b;
    return ('0' + u.toString(16)).slice(-2);
  }).join('');
}