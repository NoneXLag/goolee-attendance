/**
 * Goolee.my Attendance System — Apps Script backend
 * -------------------------------------------------
 * Deploy this bound to a Google Sheet with two tabs:
 *
 *   Employees  |  Name | PIN  | Active |
 *   Attendance |  Timestamp | Name | Action | Location (Google Maps) | Accuracy (m) | Date | Time |
 *
 * Employees.PIN must be stored as TEXT (format the column as Plain text)
 * so that PINs like "0042" keep their leading zero.
 *
 * Setup:
 *   1. Extensions > Apps Script, paste this file in as Code.gs.
 *   2. Deploy > New deployment > Web app.
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   3. Copy the web app URL into CONFIG.API_URL in index.html.
 */

const EMPLOYEES_SHEET = 'Employees';
const ATTENDANCE_SHEET = 'Attendance';
const ACCURACY_THRESHOLD_M = 50;
const PIN_MAX_ATTEMPTS = 5;
const PIN_LOCKOUT_MINUTES = 5;

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'employees') {
    return respond({ ok: true, employees: listActiveEmployees() });
  }
  if (action === 'punch') {
    const result = handlePunch(e.parameter);
    if (e.parameter.callback) {
      return respondJsonp_(result, e.parameter.callback);
    }
    return respond(result);
  }
  return respond({ ok: false, message: 'Unknown action.' });
}

function doPost(e) {
  let body;
  try {
    if (e.parameter && e.parameter.action) {
      body = e.parameter;
    } else {
      body = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    return respond({ ok: false, message: 'Malformed request.' });
  }

  if (body.action === 'punch') {
    const result = handlePunch(body);
    if (body.responseMode === 'frame') {
      return respondFrame_(result, body.requestId);
    }
    return respond(result);
  }
  return respond({ ok: false, message: 'Unknown action.' });
}

// ---------------------------------------------------------------------
// Employees
// ---------------------------------------------------------------------

function getEmployeesSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EMPLOYEES_SHEET);
}

function listActiveEmployees() {
  const sheet = getEmployeesSheet_();
  const rows = sheet.getDataRange().getValues();
  const header = rows[0].map(h => String(h).trim().toLowerCase());
  const nameCol = header.indexOf('name');
  const activeCol = header.indexOf('active');

  const employees = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const name = row[nameCol];
    const active = row[activeCol];
    if (!name) continue;
    if (active === false || String(active).toUpperCase() === 'FALSE') continue;
    employees.push({ name: String(name).trim() });
  }
  return employees;
}

function findEmployeeRow_(name) {
  const sheet = getEmployeesSheet_();
  const rows = sheet.getDataRange().getValues();
  const header = rows[0].map(h => String(h).trim().toLowerCase());
  const nameCol = header.indexOf('name');
  const pinCol = header.indexOf('pin');
  const activeCol = header.indexOf('active');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][nameCol]).trim() === name) {
      return {
        name: String(rows[i][nameCol]).trim(),
        pin: String(rows[i][pinCol]).trim(),
        active: !(rows[i][activeCol] === false ||
          String(rows[i][activeCol]).toUpperCase() === 'FALSE')
      };
    }
  }
  return null;
}

// ---------------------------------------------------------------------
// PIN lockout — tracked per employee via PropertiesService.
// Apps Script web apps do not expose the caller's IP address, so this
// mirrors the client-side lockout in index.html but is authoritative:
// even if someone bypasses the browser lockout, the server still enforces it.
// ---------------------------------------------------------------------

function lockoutKey_(name) {
  return 'lockout_' + name;
}

function isLockedOut_(name) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(lockoutKey_(name));
  if (!raw) return false;
  const state = JSON.parse(raw);
  return state.lockedUntil > Date.now();
}

function registerFailedPin_(name) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(lockoutKey_(name));
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

// ---------------------------------------------------------------------
// Punch handling
// ---------------------------------------------------------------------

function handlePunch(body) {
  const name = String(body.name || '').trim();
  const pin = String(body.pin || '').trim();
  const pinHash = String(body.pinHash || '').trim().toLowerCase();
  const punchType = body.punchType === 'Clock Out' ? 'Clock Out' : 'Clock In';
  const lat = Number(body.lat);
  const lng = Number(body.lng);
  const accuracy = Number(body.accuracy);

  if (!name) return { ok: false, message: 'Missing employee name.' };

  if (isLockedOut_(name)) {
    return {
      ok: false,
      reason: 'locked_out',
      message: `Too many incorrect PIN attempts. Try again in a few minutes.`
    };
  }

  // Core rule: every punch requires a valid GPS reading, accuracy <= 50 m.
  // No exceptions, no fallbacks, no IP-only mode.
  if (isNaN(lat) || isNaN(lng) || isNaN(accuracy) || accuracy > ACCURACY_THRESHOLD_M) {
    return {
      ok: false,
      reason: 'bad_gps',
      message: 'Location was missing or not accurate enough. Move outdoors and try again.'
    };
  }

  const employee = findEmployeeRow_(name);
  if (!employee || !employee.active) {
    return { ok: false, reason: 'unknown_employee', message: 'Employee not found or inactive.' };
  }

  const hasValidPin = /^\d{4}$/.test(pin) && pin === employee.pin;
  const hasValidPinHash = /^[a-f0-9]{64}$/.test(pinHash) && pinHash === sha256Hex_(employee.pin);
  if (!hasValidPin && !hasValidPinHash) {
    registerFailedPin_(name);
    return { ok: false, reason: 'bad_pin', message: 'Incorrect PIN.' };
  }

  clearLockout_(name);

  const now = new Date();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');

  const mapsLink = 'https://www.google.com/maps?q=' + lat + ',' + lng;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET);
  // Written as a HYPERLINK formula (not appendRow) so the Maps link is
  // clickable straight from the sheet rather than showing as plain text.
  const newRow = sheet.getLastRow() + 1;
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
    name: employee.name,
    punchType,
    date: dateStr,
    time: timeStr,
    mapsLink,
    accuracy
  };
}

// ---------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function respondJsonp_(obj, callback) {
  const safeCallback = String(callback || '').replace(/[^\w$.]/g, '');
  const body = safeCallback + '(' + JSON.stringify(obj) + ');';
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function respondFrame_(obj, requestId) {
  const payload = JSON.stringify({
    source: 'goolee-attendance',
    requestId: String(requestId || ''),
    data: obj
  }).replace(/</g, '\\u003c');

  return HtmlService.createHtmlOutput(
    '<!doctype html><html><body><script>' +
    'window.parent.postMessage(' + payload + ', "*");' +
    '</script></body></html>'
  );
}

function sha256Hex_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );

  return bytes.map(function(byte) {
    const unsigned = byte < 0 ? byte + 256 : byte;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
}
