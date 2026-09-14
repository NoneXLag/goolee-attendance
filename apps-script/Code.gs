/**
 * Goolee.my Attendance System — Apps Script backend
 * ─────────────────────────────────────────────────
 * Bound to a Google Sheet with two tabs:
 *
 *   Employees  | Name | PIN  | Active |
 *   Attendance | Timestamp | Name | Action | Location | Accuracy (m) | Date | Time |
 *
 * IMPORTANT: Format the PIN column as "Plain text" so leading zeros are kept.
 *
 * Deploy:
 *   1. Extensions → Apps Script, paste this file in as Code.gs
 *   2. Deploy → New deployment → Web app
 *        Execute as:      Me
 *        Who has access:  Anyone          ← MUST be "Anyone"
 *   3. Copy the /exec URL into CONFIG.API_URL in index.html
 *
 * Re-deploy every time you edit this file (Deploy → Manage deployments →
 * pencil icon → Version: New version → Deploy).
 */

const EMPLOYEES_SHEET      = 'Employees';
const ATTENDANCE_SHEET     = 'Attendance';
const ACCURACY_THRESHOLD_M = 50;
const PIN_MAX_ATTEMPTS     = 5;
const PIN_LOCKOUT_MINUTES  = 5;

// ═══════════════════════════════════════════════════════════════════════
// HTTP entry points — everything goes through doGet (JSONP)
// doPost is kept only for compatibility, but the frontend no longer uses it.
// ═══════════════════════════════════════════════════════════════════════
function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = params.action;

  if (action === 'employees') {
    return respondJson_(params.callback, { ok: true, employees: listActiveEmployees() });
  }

  if (action === 'punch') {
    return respondJson_(params.callback, handlePunch(params));
  }

  return respondJson_(params.callback, { ok: false, message: 'Unknown action.' });
}

function doPost(e) {
  const params = (e && e.parameter) || {};
  let body;

  if (params.action) {
    body = params;
  } else {
    try {
      body = JSON.parse(e.postData.contents);
    } catch (err) {
      return respondJson_(null, { ok: false, message: 'Malformed request.' });
    }
  }

  if (body.action === 'punch') {
    return respondJson_(null, handlePunch(body));
  }
  return respondJson_(null, { ok: false, message: 'Unknown action.' });
}

// ═══════════════════════════════════════════════════════════════════════
// Employees
// ═══════════════════════════════════════════════════════════════════════
function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet tab: ' + name);
  return sheet;
}

function listActiveEmployees() {
  const sheet  = getSheet_(EMPLOYEES_SHEET);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const header    = values[0].map(h => String(h).trim().toLowerCase());
  const nameCol   = header.indexOf('name');
  const activeCol = header.indexOf('active');
  if (nameCol < 0) return [];

  const out = [];
  for (let i = 1; i < values.length; i++) {
    const name   = values[i][nameCol];
    const active = activeCol >= 0 ? values[i][activeCol] : true;
    if (!name) continue;
    if (active === false) continue;
    if (String(active).trim().toUpperCase() === 'FALSE') continue;
    out.push({ name: String(name).trim() });
  }
  return out;
}

function findEmployee_(name) {
  const sheet  = getSheet_(EMPLOYEES_SHEET);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;

  const header    = values[0].map(h => String(h).trim().toLowerCase());
  const nameCol   = header.indexOf('name');
  const pinCol    = header.indexOf('pin');
  const activeCol = header.indexOf('active');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][nameCol]).trim() !== name) continue;
    return {
      name: String(values[i][nameCol]).trim(),
      pin:  pinCol >= 0 ? String(values[i][pinCol]).trim() : '',
      active: activeCol < 0 ? true
              : !(values[i][activeCol] === false ||
                  String(values[i][activeCol]).trim().toUpperCase() === 'FALSE')
    };
  }
  return null;
}

// ═══════════════════════════════════════════════════════════════════════
// Lockout — tracked per employee in script properties (server authoritative)
// ═══════════════════════════════════════════════════════════════════════
function lockoutKey_(name) { return 'lockout_' + name; }

function isLockedOut_(name) {
  const raw = PropertiesService.getScriptProperties().getProperty(lockoutKey_(name));
  if (!raw) return false;
  try {
    return JSON.parse(raw).lockedUntil > Date.now();
  } catch (e) {
    return false;
  }
}

function registerFailedPin_(name) {
  const props = PropertiesService.getScriptProperties();
  const raw   = props.getProperty(lockoutKey_(name));
  const state = raw ? JSON.parse(raw) : { attempts: 0, lockedUntil: 0 };

  state.attempts = (state.attempts || 0) + 1;
  if (state.attempts >= PIN_MAX_ATTEMPTS) {
    state.lockedUntil = Date.now() + PIN_LOCKOUT_MINUTES * 60 * 1000;
    state.attempts = 0;
  }
  props.setProperty(lockoutKey_(name), JSON.stringify(state));
}

function clearLockout_(name) {
  PropertiesService.getScriptProperties().deleteProperty(lockoutKey_(name));
}

// ═══════════════════════════════════════════════════════════════════════
// Punch handler
// ═══════════════════════════════════════════════════════════════════════
function handlePunch(body) {
  const name      = String(body.name || '').trim();
  const pin       = String(body.pin || '').trim();
  const pinHash   = String(body.pinHash || '').trim().toLowerCase();
  const punchType = body.punchType === 'Clock Out' ? 'Clock Out' : 'Clock In';
  const lat       = Number(body.lat);
  const lng       = Number(body.lng);
  const accuracy  = Number(body.accuracy);

  if (!name) {
    return { ok: false, message: 'Missing employee name.' };
  }

  if (isLockedOut_(name)) {
    return {
      ok: false,
      reason: 'locked_out',
      message: 'Too many incorrect PIN attempts. Try again in a few minutes.'
    };
  }

  // Mandatory GPS, accuracy ≤ 50 m — no fallback, no IP-only mode
  if (isNaN(lat) || isNaN(lng) || isNaN(accuracy) || accuracy > ACCURACY_THRESHOLD_M) {
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

  // Accept either the plain PIN (server-side testing) or its SHA-256 hash
  const hasValidPin     = /^\d{4}$/.test(pin) && pin === employee.pin;
  const hasValidPinHash = /^[a-f0-9]{64}$/.test(pinHash) &&
                          pinHash === sha256Hex_(employee.pin);

  if (!hasValidPin && !hasValidPinHash) {
    registerFailedPin_(name);
    return { ok: false, reason: 'bad_pin', message: 'Incorrect PIN.' };
  }

  clearLockout_(name);

  const now = new Date();
  const tz  = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');

  const mapsLink = 'https://www.google.com/maps?q=' + lat + ',' + lng;

  const sheet  = getSheet_(ATTENDANCE_SHEET);
  const newRow = sheet.getLastRow() + 1;

  // Write 7 columns. Location column gets a clickable HYPERLINK formula
  // so HR can open the map directly from the sheet.
  sheet.getRange(newRow, 1, 1, 7).setValues([[
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
    name:      employee.name,
    punchType: punchType,
    date:      dateStr,
    time:      timeStr,
    mapsLink:  mapsLink,
    accuracy:  accuracy
  };
}

// ═══════════════════════════════════════════════════════════════════════
// Response helper — JSON or JSONP
// ═══════════════════════════════════════════════════════════════════════
function respondJson_(callback, obj) {
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

// ═══════════════════════════════════════════════════════════════════════
// SHA-256 helper — must match the client-side hashPin() in index.html
// ═══════════════════════════════════════════════════════════════════════
function sha256Hex_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function (b) {
    const unsigned = b < 0 ? b + 256 : b;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
}