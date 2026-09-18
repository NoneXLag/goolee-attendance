/**
 * Goolee.my Attendance System — Apps Script backend
 * ─────────────────────────────────────────────────
 * • Reads all settings from the Settings sheet
 * • Exposes GET ?action=settings and ?action=employees
 * • Handles punches with GPS + PIN verification
 * • Prevents duplicate punches on the same day
 * • Computes early-leave / overtime on clock-out
 * • Auto-refreshes the dashboard ~5s after every punch
 *
 * Duplicate-punch detection now reads the REAL timestamp in column A
 * (never the date string in column F, which Sheets can silently
 * re-type as a Date and shift across timezones).
 *
 * Constants use the API_ prefix where they could clash with Dashboard.gs
 * (Apps Script merges all .gs files into one global scope).
 */

// ── Sheet tabs ──
const EMPLOYEES_TAB  = 'Employees';
const ATTENDANCE_TAB = 'Attendance';

// ── Rules ──
const ACCURACY_MAX_M      = 100;
const PIN_LOCKOUT_TRIES   = 5;
const PIN_LOCKOUT_MINUTES = 5;

// ── Cache keys (API-specific, prefixed to avoid Dashboard.gs clashes) ──
const CACHE_EMPLOYEES_KEY      = 'goolee_employees_cache_v2';
const CACHE_EMPLOYEES_FULL_KEY = 'goolee_employees_full_v2';
const API_CACHE_SETTINGS_KEY   = 'goolee_settings_api_v1';
const CACHE_EMPLOYEES_TTL      = 600;   // 10 min
const API_CACHE_SETTINGS_TTL   = 600;   // 10 min

// ── Dashboard auto-refresh timing ──
const DASHBOARD_REFRESH_DELAY_MS = 5000;   // 5 s after punch
const DASHBOARD_MIN_INTERVAL_MS  = 30000;  // min 30 s between refreshes

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
    return respond_(p.callback, {
      ok: true,
      employees: listEmployees_(p.refresh === '1')
    });
  }
  if (action === 'settings') {
    return respond_(p.callback, { ok: true, settings: getSettings_() });
  }
  if (action === 'punch') {
    return respond_(p.callback, handlePunch_(p));
  }
  return respond_(p.callback, { ok: false, message: 'Unknown action: ' + action });
}

function doPost(e) {
  const p = (e && e.parameter) || {};
  const body = p.action ? p : safeJson_(e.postData && e.postData.contents);
  if (!body) return respond_(null, { ok: false, message: 'Malformed request.' });
  if (body.action === 'punch') return respond_(null, handlePunch_(body));
  return respond_(null, { ok: false, message: 'Unknown action.' });
}

// ═══════════════════════════════════════════════════════════════════════
// SETTINGS — read from Settings sheet, cached
// ═══════════════════════════════════════════════════════════════════════
function getSettings_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(API_CACHE_SETTINGS_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const out = {
    workingDays:    'Mon,Tue,Wed,Thu,Fri',
    workStartHour:  10,
    workStartMin:   0,
    workEndHour:    19,
    workEndMin:     0,
    lateAfterMin:   15,
    earlyLeaveMin:  15,
    companyName:    'Goolee'
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Settings');
    if (sheet && sheet.getLastRow() >= 2) {
      const values = sheet.getRange(1, 1, sheet.getLastRow(), 3).getValues();
      for (let i = 1; i < values.length; i++) {
        const key = String(values[i][0] || '').trim().toLowerCase();
        const val = String(values[i][1] || '').trim();
        if (!key || !val) continue;

        if (key === 'working days') {
          out.workingDays = val;
        } else if (key === 'work start time') {
          const parts = val.split(':');
          out.workStartHour = parseInt(parts[0], 10);
          out.workStartMin  = parseInt(parts[1] || '0', 10);
        } else if (key === 'work end time') {
          const parts = val.split(':');
          out.workEndHour = parseInt(parts[0], 10);
          out.workEndMin  = parseInt(parts[1] || '0', 10);
        } else if (key === 'late after (min)') {
          out.lateAfterMin = parseInt(val, 10) || 15;
        } else if (key === 'early leave grace (min)') {
          out.earlyLeaveMin = parseInt(val, 10) || 15;
        } else if (key === 'company name') {
          out.companyName = val;
        }
      }
    }
  } catch (e) {
    console.error('Failed to read settings:', e);
  }

  if (isNaN(out.workStartHour)) out.workStartHour = 10;
  if (isNaN(out.workStartMin))  out.workStartMin  = 0;
  if (isNaN(out.workEndHour))   out.workEndHour   = 19;
  if (isNaN(out.workEndMin))    out.workEndMin    = 0;

  try { cache.put(API_CACHE_SETTINGS_KEY, JSON.stringify(out), API_CACHE_SETTINGS_TTL); } catch (e) {}
  return out;
}

function clearApiSettingsCache() {
  CacheService.getScriptCache().remove(API_CACHE_SETTINGS_KEY);
  SpreadsheetApp.getActiveSpreadsheet().toast('API settings cache cleared.', 'Goolee', 3);
}

// ═══════════════════════════════════════════════════════════════════════
// PIN normalisation
// ═══════════════════════════════════════════════════════════════════════
function normalizePin_(raw) {
  if (raw === null || raw === undefined) return '';
  let s;
  if (typeof raw === 'number') s = String(Math.round(raw));
  else s = String(raw).trim();
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

// ═══════════════════════════════════════════════════════════════════════
// Employees — cached
// ═══════════════════════════════════════════════════════════════════════
function listEmployees_(forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    const cached = cache.get(CACHE_EMPLOYEES_KEY);
    if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  }

  const values = sheet_(EMPLOYEES_TAB).getDataRange().getValues();
  if (values.length < 2) return [];
  const header = values[0].map(h => String(h).trim().toLowerCase());
  const nameIdx   = header.indexOf('name');
  const activeIdx = header.indexOf('active');
  if (nameIdx < 0) return [];

  const out = [];
  for (let i = 1; i < values.length; i++) {
    const name = values[i][nameIdx];
    const active = activeIdx >= 0 ? values[i][activeIdx] : true;
    if (!name) continue;
    if (active === false) continue;
    if (String(active).trim().toUpperCase() === 'FALSE') continue;
    out.push({ name: String(name).trim() });
  }
  try { cache.put(CACHE_EMPLOYEES_KEY, JSON.stringify(out), CACHE_EMPLOYEES_TTL); } catch (e) {}
  return out;
}

function findEmployee_(name) {
  const cache = CacheService.getScriptCache();
  let full = null;
  const cached = cache.get(CACHE_EMPLOYEES_FULL_KEY);
  if (cached) { try { full = JSON.parse(cached); } catch (e) { full = null; } }

  if (!full) {
    const values = sheet_(EMPLOYEES_TAB).getDataRange().getValues();
    if (values.length < 2) return null;
    const header = values[0].map(h => String(h).trim().toLowerCase());
    const nameIdx   = header.indexOf('name');
    const pinIdx    = header.indexOf('pin');
    const activeIdx = header.indexOf('active');
    full = [];
    for (let i = 1; i < values.length; i++) {
      const empName = String(values[i][nameIdx] || '').trim();
      if (!empName) continue;
      full.push({
        name: empName,
        pin: pinIdx >= 0 ? normalizePin_(values[i][pinIdx]) : '',
        active: activeIdx < 0 ? true
                : !(values[i][activeIdx] === false ||
                    String(values[i][activeIdx]).trim().toUpperCase() === 'FALSE')
      });
    }
    try { cache.put(CACHE_EMPLOYEES_FULL_KEY, JSON.stringify(full), CACHE_EMPLOYEES_TTL); } catch (e) {}
  }
  for (let i = 0; i < full.length; i++) {
    if (full[i].name === name) return full[i];
  }
  return null;
}

function clearApiEmployeeCache() {
  const cache = CacheService.getScriptCache();
  cache.remove(CACHE_EMPLOYEES_KEY);
  cache.remove(CACHE_EMPLOYEES_FULL_KEY);
}

// ═══════════════════════════════════════════════════════════════════════
// Lockout
// ═══════════════════════════════════════════════════════════════════════
function lockKey_(name) { return 'lockout_' + name; }

function isLocked_(name) {
  const raw = PropertiesService.getScriptProperties().getProperty(lockKey_(name));
  if (!raw) return false;
  try { return JSON.parse(raw).until > Date.now(); } catch (e) { return false; }
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
// Duplicate punch check — reads the REAL timestamp in column A.
// Do NOT trust column F (the date string). Sheets silently re-types it
// as a Date and the value can drift across timezone boundaries, which
// caused "already clocked in" to trigger on a fresh day.
// ═══════════════════════════════════════════════════════════════════════
function getTodayPunches_(employeeName, tz) {
  const sheet = sheet_(ATTENDANCE_TAB);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { hasIn: false, hasOut: false, inTime: null, outTime: null };

  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const lookback = Math.min(lastRow - 1, 1000);
  const startRow = lastRow - lookback + 1;
  const values = sheet.getRange(startRow, 1, lookback, 7).getValues();

  let hasIn = false, hasOut = false, inTime = null, outTime = null;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const ts      = row[0];                       // column A — the real Date
    const rowName = String(row[1] || '').trim();
    const action  = String(row[2] || '').trim();

    if (rowName !== employeeName) continue;
    if (!(ts instanceof Date)) continue;          // skip malformed rows

    // Format the row's own timestamp in the sheet TZ — never trust column F
    const rowDateStr = Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
    if (rowDateStr !== todayStr) continue;

    if (action === 'Clock In') {
      hasIn = true;
      if (!inTime) inTime = Utilities.formatDate(ts, tz, 'h:mm a');
    } else if (action === 'Clock Out') {
      hasOut = true;
      if (!outTime) outTime = Utilities.formatDate(ts, tz, 'h:mm a');
    }
  }
  return { hasIn, hasOut, inTime, outTime };
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

  if (!name) return { ok: false, message: 'Missing employee name.' };

  if (isLocked_(name)) {
    return { ok: false, reason: 'locked_out',
             message: 'Too many incorrect PIN attempts. Try again in a few minutes.' };
  }
  if (isNaN(lat) || isNaN(lng) || isNaN(accuracy) || accuracy > ACCURACY_MAX_M) {
    return { ok: false, reason: 'bad_gps',
             message: 'Location was missing or not accurate enough. Move outdoors and try again.' };
  }

  const employee = findEmployee_(name);
  if (!employee || !employee.active) {
    return { ok: false, reason: 'unknown_employee',
             message: 'Employee not found or inactive.' };
  }

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

  // ── Duplicate check now reads column A only (see getTodayPunches_) ──
  const existing = getTodayPunches_(employee.name, tz);

  if (punchType === 'Clock In' && existing.hasIn) {
    return {
      ok: false,
      reason: 'already_clocked_in',
      message: 'You have already clocked in today at ' + (existing.inTime || 'earlier') + '.',
      existingTime: existing.inTime
    };
  }
  if (punchType === 'Clock Out') {
    if (!existing.hasIn) {
      return { ok: false, reason: 'no_clock_in',
               message: "You haven't clocked in today yet. Please Clock In first." };
    }
    if (existing.hasOut) {
      return {
        ok: false,
        reason: 'already_clocked_out',
        message: 'You have already clocked out today at ' + (existing.outTime || 'earlier') + '.',
        existingTime: existing.outTime
      };
    }
  }

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

  // ── Early / overtime calculation for Clock Out ──
  const settings = getSettings_();
  let punchStatus = 'normal';
  let statusLabel = '';
  let overtimeMinutes = 0;
  let earlyMinutes = 0;

  if (punchType === 'Clock Out') {
    const nowH = Number(Utilities.formatDate(now, tz, 'H'));
    const nowM = Number(Utilities.formatDate(now, tz, 'm'));
    const nowMin = nowH * 60 + nowM;
    const endMin = settings.workEndHour * 60 + settings.workEndMin;
    const earlyThreshold = endMin - settings.earlyLeaveMin;

    if (nowMin > endMin) {
      overtimeMinutes = nowMin - endMin;
      punchStatus = 'overtime';
      statusLabel = formatDuration_(overtimeMinutes);
    } else if (nowMin < earlyThreshold) {
      earlyMinutes = endMin - nowMin;
      punchStatus = 'early';
      statusLabel = formatDuration_(earlyMinutes);
    }
  }

  SpreadsheetApp.flush();
  try { scheduleDashboardRefresh(); } catch (e) { console.error('schedule failed:', e); }

  return {
    ok: true,
    name: employee.name,
    punchType: punchType,
    date: dateStr,
    time: timeStr,
    mapsLink: mapsLink,
    accuracy: accuracy,
    punchStatus: punchStatus,
    statusLabel: statusLabel,
    overtimeMinutes: overtimeMinutes,
    earlyMinutes: earlyMinutes,
    workEndLabel: formatTimeLabel_(settings.workEndHour, settings.workEndMin)
  };
}

// ═══════════════════════════════════════════════════════════════════════
// Formatting helpers
// ═══════════════════════════════════════════════════════════════════════
function formatDuration_(mins) {
  mins = Math.round(mins);
  if (mins <= 0) return '0m';
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  if (h === 0) return m + 'm';
  if (m === 0) return h + 'h';
  return h + 'h ' + m + 'm';
}

function formatTimeLabel_(h, m) {
  const period = h >= 12 ? 'PM' : 'AM';
  const h12 = h % 12 || 12;
  return h12 + ':' + String(m).padStart(2, '0') + ' ' + period;
}

// ═══════════════════════════════════════════════════════════════════════
// Dashboard auto-refresh
// ═══════════════════════════════════════════════════════════════════════
function scheduleDashboardRefresh() {
  const props = PropertiesService.getScriptProperties();

  // Auto-reset stuck flag (older than 2 min)
  const pending = props.getProperty('dashboard_refresh_pending');
  if (pending === 'true') {
    const at = Number(props.getProperty('dashboard_refresh_pending_at') || 0);
    if (Date.now() - at < 120000) return;
    props.deleteProperty('dashboard_refresh_pending');
    props.deleteProperty('dashboard_refresh_pending_at');
  }

  // Rate limit
  const last = Number(props.getProperty('dashboard_last_refresh') || 0);
  if (Date.now() - last < DASHBOARD_MIN_INTERVAL_MS) return;

  props.setProperty('dashboard_refresh_pending', 'true');
  props.setProperty('dashboard_refresh_pending_at', String(Date.now()));

  // Clean up old triggers
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runScheduledDashboardRefresh') {
      ScriptApp.deleteTrigger(t);
    }
  });

  try {
    ScriptApp.newTrigger('runScheduledDashboardRefresh')
      .timeBased().after(DASHBOARD_REFRESH_DELAY_MS).create();
    console.log('Dashboard refresh scheduled');
  } catch (e) {
    console.error('Failed to create trigger:', e);
    props.deleteProperty('dashboard_refresh_pending');
    props.deleteProperty('dashboard_refresh_pending_at');
  }
}

function runScheduledDashboardRefresh() {
  const props = PropertiesService.getScriptProperties();

  try {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === 'runScheduledDashboardRefresh') {
        ScriptApp.deleteTrigger(t);
      }
    });
  } catch (e) { console.error('Trigger cleanup failed:', e); }

  props.deleteProperty('dashboard_refresh_pending');
  props.deleteProperty('dashboard_refresh_pending_at');
  props.setProperty('dashboard_last_refresh', String(Date.now()));

  try {
    if (typeof refreshDashboard === 'function') {
      refreshDashboard();
      console.log('✓ Dashboard auto-refreshed after punch');
    } else {
      console.error('refreshDashboard() not found — check Dashboard.gs is loaded');
    }
  } catch (e) {
    console.error('Auto-refresh dashboard failed:', e);
  }
}

function setupDashboardAutoRefresh() {
  const props = PropertiesService.getScriptProperties();

  props.deleteProperty('dashboard_refresh_pending');
  props.deleteProperty('dashboard_refresh_pending_at');
  props.deleteProperty('dashboard_last_refresh');

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runScheduledDashboardRefresh') {
      ScriptApp.deleteTrigger(t);
    }
  });

  try {
    ScriptApp.newTrigger('runScheduledDashboardRefresh')
      .timeBased().after(30000).create();
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Could not create trigger.\n\n' + e.message +
      '\n\nTry: Apps Script → Triggers → Add Trigger manually.'
    );
    return;
  }

  SpreadsheetApp.getUi().alert(
    '✅ Auto-refresh enabled.\n\n' +
    'A test refresh has been scheduled for 30 seconds from now.\n' +
    'After that, every punch triggers a dashboard update.'
  );
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

// ═══════════════════════════════════════════════════════════════════════
// DIAGNOSTIC + REPAIR TOOLS
// Run these manually from the Apps Script editor if you ever see
// "already clocked in" on a day with no prior punch.
// ═══════════════════════════════════════════════════════════════════════

/**
 * Diagnose the sheet timezone and the last few attendance rows.
 * Apps Script editor → pick diagnoseAttendanceDates → Run.
 */
function diagnoseAttendanceDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ATTENDANCE_TAB);
  const tz = ss.getSpreadsheetTimeZone();

  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const scriptTz = Session.getScriptTimeZone();

  const lastRow = sheet.getLastRow();
  const alertFn = (typeof SpreadsheetApp.getUi === 'function')
    ? function (title, msg) { SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK); }
    : function (title, msg) { console.log(title + '\n' + msg); };

  if (lastRow < 2) { alertFn('Attendance date diagnosis', 'Attendance sheet is empty.'); return; }

  const lookback = Math.min(lastRow - 1, 10);
  const startRow = lastRow - lookback + 1;
  const values = sheet.getRange(startRow, 1, lookback, 7).getValues();

  let msg = 'Sheet TZ: ' + tz + '\n' +
            'Script TZ: ' + scriptTz + '\n' +
            'Server "today": ' + todayStr + '\n\n' +
            'Last ' + values.length + ' rows:\n\n';

  values.forEach(function (row, i) {
    const ts       = row[0];
    const name     = String(row[1] || '').trim();
    const action   = String(row[2] || '').trim();
    const dateF    = row[5];
    const typeF    = (dateF instanceof Date) ? 'Date' : typeof dateF;
    const dateFmt  = (dateF instanceof Date)
      ? Utilities.formatDate(dateF, tz, 'yyyy-MM-dd')
      : String(dateF);
    const tsFmt    = (ts instanceof Date)
      ? Utilities.formatDate(ts, tz, 'yyyy-MM-dd HH:mm')
      : String(ts);

    msg += 'Row ' + (startRow + i) + ' · ' + name + ' · ' + action + '\n' +
           '  A (timestamp): ' + tsFmt + '\n' +
           '  F (date cell): ' + dateFmt + '  [' + typeF + ']\n\n';
  });

  alertFn('Attendance date diagnosis', msg);
}

/**
 * One-off: rebuild column F from column A using the current sheet TZ,
 * and lock column F to plain text so Sheets never re-types it again.
 */
function repairAttendanceDateColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ATTENDANCE_TAB);
  const tz = ss.getSpreadsheetTimeZone();
  const lastRow = sheet.getLastRow();

  const alertFn = (typeof SpreadsheetApp.getUi === 'function')
    ? function (title, msg) { SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK); }
    : function (title, msg) { console.log(title + '\n' + msg); };

  if (lastRow < 2) { alertFn('Repair skipped', 'Attendance sheet is empty.'); return; }

  const range = sheet.getRange(2, 6, lastRow - 1, 1);   // column F only
  const colA  = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const out   = colA.map(function (r) {
    const ts = r[0];
    return [(ts instanceof Date)
      ? Utilities.formatDate(ts, tz, 'yyyy-MM-dd')
      : ''];
  });

  range.setNumberFormat('@');   // force plain text so Sheets stops re-typing it
  range.setValues(out);

  alertFn(
    'Repair complete',
    '✅ Column F rebuilt from column A using TZ ' + tz + '.\n\n' +
    'If the sheet TZ is not Asia/Kuala_Lumpur, change it now:\n' +
    'File → Settings → Time zone → (GMT+08:00) Kuala Lumpur'
  );
}

/**
 * Clear every cached config so the next punch rereads Settings,
 * Employees, Holidays and Dashboard layout fresh.
 */
function clearAllCachesApi() {
  const cache = CacheService.getScriptCache();
  cache.remove(CACHE_EMPLOYEES_KEY);
  cache.remove(CACHE_EMPLOYEES_FULL_KEY);
  cache.remove(API_CACHE_SETTINGS_KEY);
  cache.remove('goolee_settings_v2');
  cache.remove('goolee_holidays_v2');
  SpreadsheetApp.getActiveSpreadsheet().toast('All caches cleared.', 'Goolee', 3);
}