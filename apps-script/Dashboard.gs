/**
 * ════════════════════════════════════════════════════════════════════
 *  Goolee.my — Attendance Dashboard
 * ════════════════════════════════════════════════════════════════════
 *  • Every section aligned to columns A:L
 *  • breakApart() at start so leftover merges never linger
 *  • Today's Performance + weekly hours monitoring + leave + holidays + daily log
 * ════════════════════════════════════════════════════════════════════
 */

// ── Sheet names ──
const SHEET_DASHBOARD  = 'Dashboard';
const SHEET_ATTENDANCE = 'Attendance';
const SHEET_EMPLOYEES  = 'Employees';
const SHEET_LEAVE      = 'Leave';
const SHEET_SETTINGS   = 'Settings';
const SHEET_HOLIDAYS   = 'Holidays';

// ── Defaults ──
const DEFAULT_WORKING_DAYS     = 'Mon,Tue,Wed,Thu,Fri';
const DEFAULT_WORK_START_HOUR  = null;
const DEFAULT_WORK_START_MIN   = null;
const DEFAULT_WORK_END_HOUR    = null;
const DEFAULT_WORK_END_MIN     = null;
const DEFAULT_LATE_AFTER_MIN   = 15;
const DEFAULT_EARLY_LEAVE_MIN  = 15;
const DEFAULT_WEEKLY_HOURS_TARGET = 45;

const LOG_DAYS_BACK     = 14;
const EMP_LOG_DAYS_BACK = 30;

// ── Cache keys ──
const CACHE_SETTINGS_KEY = 'goolee_settings_v4';
const CACHE_HOLIDAYS_KEY = 'goolee_holidays_v2';
const CACHE_CONFIG_TTL   = 1800;

// ── Leave types ──
const LEAVE_TYPES = ['MC', 'Emergency Leave', 'Annual Leave', 'Unpaid Leave'];

// ── Total columns used by the Dashboard ──
const DASH_TOTAL_COLS = 12;

// ── Palette ──
const C = {
  tealDark:   '#0F4038', teal:       '#176B5D', tealLight:  '#E5F2EE',
  green:      '#23865B', greenLight: '#E7F5EE',
  red:        '#B43D2A', redLight:   '#FCEAE6',
  amber:      '#9C6A12', amberLight: '#FFF4DF',
  blue:       '#1E4B8F', blueLight:  '#E4EEFB',
  purple:     '#6B3FA0', purpleLight:'#EFE5F7',
  gray:       '#62716D', grayLight:  '#EEF2F1',
  white:      '#FFFFFF', border:     '#D9E0DD', ink:        '#17211F'
};

const LEAVE_COLORS = {
  'MC':              { bg: C.redLight,    fg: C.red,    icon: '🏥' },
  'Emergency Leave': { bg: C.amberLight,  fg: C.amber,  icon: '🚨' },
  'Annual Leave':    { bg: C.tealLight,   fg: C.teal,   icon: '🌴' },
  'Unpaid Leave':    { bg: C.grayLight,   fg: C.gray,   icon: '📋' }
};

const HOLIDAY_COLOR = { bg: C.purpleLight, fg: C.purple, icon: '🎉' };

function leaveColor_(type) {
  return LEAVE_COLORS[type] || { bg: C.blueLight, fg: C.blue, icon: '📌' };
}

// ════════════════════════════════════════════════════════════════════
//  MENU
// ════════════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Goolee')
    .addItem('🔄  Refresh Dashboard',            'refreshDashboard')
    .addItem('🛠️  Rebuild Dashboard',            'setupDashboard')
    .addSeparator()
    .addItem('👤  Refresh Employee Tabs',        'refreshEmployeeTabs')
    .addItem('🔁  Force rebuild employee tabs',  'forceRebuildEmployeeTabs')
    .addSeparator()
    .addItem('📝  Open Leave Sheet',             'openLeaveSheet')
    .addItem('🔄  Refresh Leave Dropdown',       'refreshLeaveDropdown')
    .addItem('🎉  Open Holidays Sheet',          'openHolidaysSheet')
    .addItem('⚙️  Open Settings Sheet',          'openSettingsSheet')
    .addSeparator()
    .addItem('🧹  Clear all caches',             'clearAllCaches')
    .addItem('🎨  Refresh attendance formatting', 'highlightLateInAttendance')
    .addSeparator()
    .addItem('⏱️  Enable auto-refresh after punch', 'setupDashboardAutoRefresh')
    .addItem('🔍  Diagnose auto-refresh',        'diagnoseAutoRefresh')
    .addItem('⏱️  Enable hourly auto-refresh',   'enableAutoRefresh')
    .addItem('⏹️  Disable auto-refresh',         'disableAutoRefresh')
    .addToUi();
}

function openLeaveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_LEAVE);
  if (!sheet) sheet = setupLeaveSheet();
  ss.setActiveSheet(sheet);
}

function openHolidaysSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_HOLIDAYS);
  if (!sheet) sheet = setupHolidaysSheet();
  ss.setActiveSheet(sheet);
}

function openSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) sheet = setupSettingsSheet();
  ss.setActiveSheet(sheet);
}

// ════════════════════════════════════════════════════════════════════
//  SETUP
// ════════════════════════════════════════════════════════════════════
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  setupLeaveSheet();
  setupHolidaysSheet();
  setupSettingsSheet();

  let dash = ss.getSheetByName(SHEET_DASHBOARD);
  if (dash) {
    const ans = ui.alert(
      'Rebuild Dashboard?',
      'This will replace the existing Dashboard sheet.\n\nContinue?',
      ui.ButtonSet.YES_NO
    );
    if (ans !== ui.Button.YES) return;
    ss.deleteSheet(dash);
  }

  dash = ss.insertSheet(SHEET_DASHBOARD, 0);
  dash.setTabColor(C.teal);

  clearAllCachesInternal_();

  refreshDashboard();
  highlightLateInAttendanceInternal_();

  ss.setActiveSheet(dash);

  ui.alert(
    '✅ Dashboard created.\n\n' +
    'Sheets: Settings · Holidays · Leave · Dashboard · <employee tabs>'
  );
}

function setupSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SETTINGS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SETTINGS);
  } else if (sheet.getLastRow() >= 1 && sheet.getRange(1, 1).getValue() === 'Setting') {
    ensureSaturdaySettings_(sheet);
    return sheet;
  }

  const headers = ['Setting', 'Value', 'Notes'];
  sheet.getRange(1, 1, 1, 3).setValues([headers])
    .setBackground(C.tealDark).setFontColor(C.white)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 440);

  const rows = [
    ['Working Days',             DEFAULT_WORKING_DAYS, 'Use Mon,Tue,Wed,Thu,Fri or Mon-Fri.'],
    ['Saturday Working Day',     'TRUE',               'Saturday is a working day with its own hours. Set FALSE to disable.'],
    ['Saturday Work Start Time', '09:00',              '24-hour format. Saturday only.'],
    ['Saturday Work End Time',   '12:00',              '24-hour format. Saturday only.'],
    ['Work Start Time',          '10:00',              'Legacy only. Clock-in time is flexible.'],
    ['Work End Time',            '19:00',              'Legacy only. Clock-out time is flexible.'],
    ['Late After (min)',         '15',                 'Legacy only. Flexible punches are never marked late.'],
    ['Early Leave Grace (min)',  '15',                 'Legacy only. Flexible punches are never marked early.'],
    ['Minimum Weekly Hours',     '45',                 'Target hours per employee from Monday to Sunday.'],
    ['GPS Accuracy Limit (m)',   '200',                'Maximum GPS uncertainty accepted for a punch.'],
    ['Company Name',             'Goolee',             'Shown in the dashboard title.']
  ];
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  sheet.getRange(2, 1, rows.length, 1)
    .setFontWeight('bold').setFontColor(C.tealDark).setBackground(C.tealLight);
  sheet.getRange(2, 1, rows.length, 3).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(2, 3, rows.length, 1)
    .setFontSize(10).setFontColor(C.gray).setFontStyle('italic');

  sheet.setTabColor(C.blue);
  return sheet;
}

function ensureSaturdaySettings_(sheet) {
  const keys = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1)
    .getValues().map(function (row) { return String(row[0] || '').trim().toLowerCase(); });
  const missing = [
    ['Saturday Working Day', 'TRUE', 'Saturday is a working day with its own hours. Set FALSE to disable.'],
    ['Saturday Work Start Time', '09:00', '24-hour format. Saturday only.'],
    ['Saturday Work End Time', '12:00', '24-hour format. Saturday only.'],
    ['Minimum Weekly Hours', '45', 'Target hours per employee from Monday to Sunday.'],
    ['GPS Accuracy Limit (m)', '200', 'Maximum GPS uncertainty accepted for a punch.']
  ].filter(function (row) { return keys.indexOf(row[0].toLowerCase()) < 0; });
  if (missing.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
  }
}

function setupHolidaysSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_HOLIDAYS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_HOLIDAYS);
  } else if (sheet.getLastRow() >= 1 && sheet.getRange(1, 1).getValue() === 'Date') {
    return sheet;
  }

  const headers = ['Date', 'Holiday Name', 'Type'];
  sheet.getRange(1, 1, 1, 3).setValues([headers])
    .setBackground(C.tealDark).setFontColor(C.white)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 180);
  sheet.getRange('A2:A1000').setNumberFormat('yyyy-mm-dd');

  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Public Holiday', 'Company Holiday', 'Replacement Holiday'], true)
    .setAllowInvalid(true)
    .setHelpText('Pick a holiday type.')
    .build();
  sheet.getRange('C2:C1000').setDataValidation(typeRule);

  sheet.getRange('F1').setValue('💡  Add one row per holiday.')
    .setFontColor(C.gray).setFontStyle('italic').setFontSize(10);

  sheet.setTabColor(C.purple);
  return sheet;
}

function setupLeaveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_LEAVE);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LEAVE);
  } else if (sheet.getLastRow() >= 1 && sheet.getRange(1, 1).getValue() === 'Date') {
    return sheet;
  }

  const headers = ['Date', 'Name', 'Type', 'Notes'];
  sheet.getRange(1, 1, 1, 4).setValues([headers])
    .setBackground(C.tealDark).setFontColor(C.white)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 170);
  sheet.setColumnWidth(4, 340);
  sheet.getRange('A2:A1000').setNumberFormat('yyyy-mm-dd');

  const empSheet = ss.getSheetByName(SHEET_EMPLOYEES);
  if (empSheet) {
    const nameRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(empSheet.getRange('A2:A1000'), true)
      .setAllowInvalid(true)
      .setHelpText('Pick a name from the Employees sheet.')
      .build();
    sheet.getRange('B2:B1000').setDataValidation(nameRule);
  }

  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LEAVE_TYPES, true)
    .setAllowInvalid(false)
    .setHelpText('Pick a leave type.')
    .build();
  sheet.getRange('C2:C1000').setDataValidation(typeRule);

  sheet.getRange('F1').setValue('💡  Add one row per leave day.')
    .setFontColor(C.gray).setFontStyle('italic').setFontSize(10);

  sheet.setTabColor(C.amber);
  return sheet;
}

// ════════════════════════════════════════════════════════════════════
//  REFRESH — individual
// ════════════════════════════════════════════════════════════════════
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SHEET_DASHBOARD);
  if (!dash) {
    SpreadsheetApp.getUi().alert(
      'No Dashboard sheet yet.\nRun Goolee → Rebuild Dashboard first.'
    );
    return;
  }
  const ctx = buildContext_();
  if (!ctx) return;
  paintDashboard_(dash, ctx);
  // Remove legacy late/early rules from existing spreadsheets. Flexible
  // punch times are never formatted red; only inaccurate GPS is highlighted.
  try { highlightLateInAttendanceInternal_(); } catch (e) { console.error('Attendance formatting failed:', e); }
  refreshEmployeeTabs_(ctx);
}

function refreshEmployeeTabs() {
  PropertiesService.getScriptProperties().deleteProperty('employee_tabs_sig');
  refreshDashboard();
  SpreadsheetApp.getActiveSpreadsheet().toast('Employee tabs refreshed.', 'Goolee', 3);
}

function forceRebuildEmployeeTabs() {
  PropertiesService.getScriptProperties().deleteProperty('employee_tabs_sig');
  refreshDashboard();
  SpreadsheetApp.getActiveSpreadsheet().toast('Employee tabs rebuilt.', 'Goolee', 3);
}

function refreshLeaveDropdown() {
  try {
    refreshLeaveDropdownInternal_();
    SpreadsheetApp.getActiveSpreadsheet().toast('Leave dropdown refreshed.', 'Goolee', 3);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Leave dropdown failed:\n\n' + e.message);
  }
}

function highlightLateInAttendance() {
  try {
    highlightLateInAttendanceInternal_();
    SpreadsheetApp.getActiveSpreadsheet().toast('Attendance highlight applied.', 'Goolee', 3);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Highlight failed:\n\n' + e.message);
  }
}

function clearAllCaches() {
  try {
    clearAllCachesInternal_();
    SpreadsheetApp.getActiveSpreadsheet().toast('All caches cleared.', 'Goolee', 3);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Clear caches failed:\n\n' + e.message);
  }
}

// ════════════════════════════════════════════════════════════════════
//  INTERNAL HELPERS
// ════════════════════════════════════════════════════════════════════
function clearAllCachesInternal_() {
  const cache = CacheService.getScriptCache();
  cache.remove(CACHE_SETTINGS_KEY);
  cache.remove(CACHE_HOLIDAYS_KEY);
  cache.remove('goolee_employees_cache_v2');
  cache.remove('goolee_employees_full_v2');
  cache.remove('goolee_settings_api_v3');
}

function refreshLeaveDropdownInternal_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ss.getSheetByName(SHEET_EMPLOYEES);
  const leaveSheet = ss.getSheetByName(SHEET_LEAVE);

  if (!empSheet)   throw new Error('Missing Employees sheet');
  if (!leaveSheet) throw new Error('Missing Leave sheet');

  const nameRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(empSheet.getRange('A2:A1000'), true)
    .setAllowInvalid(true)
    .setHelpText('Pick a name from the Employees sheet.')
    .build();
  leaveSheet.getRange('B2:B1000').setDataValidation(nameRule);

  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LEAVE_TYPES, true)
    .setAllowInvalid(false)
    .setHelpText('Pick a leave type.')
    .build();
  leaveSheet.getRange('C2:C1000').setDataValidation(typeRule);
}

function highlightLateInAttendanceInternal_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) throw new Error('Attendance sheet not found');

  sheet.clearConditionalFormatRules();

  const lastRow = Math.max(sheet.getMaxRows(), 200);
  const range = sheet.getRange(2, 1, lastRow - 1, 7);

  const accuracyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($E2<>"", $E2 > 200)')
    .setBackground(C.amberLight).setFontColor(C.amber).setRanges([range]).build();

  // Clock times are flexible. Only an inaccurate GPS reading is highlighted.
  sheet.setConditionalFormatRules([accuracyRule]);
}

// ════════════════════════════════════════════════════════════════════
//  SETTINGS READER
// ════════════════════════════════════════════════════════════════════
function readSettings_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_SETTINGS_KEY);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      parsed.workingDays = new Set(parsed.workingDays);
      return parsed;
    } catch (e) {}
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SETTINGS);

  const out = {
    workingDays:    parseWorkingDays_(DEFAULT_WORKING_DAYS),
    saturdayWorkingDay: true,
    workStartHour:  DEFAULT_WORK_START_HOUR,
    workStartMin:   DEFAULT_WORK_START_MIN,
    workEndHour:    DEFAULT_WORK_END_HOUR,
    workEndMin:     DEFAULT_WORK_END_MIN,
    saturdayStartHour: 9,
    saturdayStartMin:  0,
    saturdayEndHour:   12,
    saturdayEndMin:    0,
    lateAfterMin:   DEFAULT_LATE_AFTER_MIN,
    earlyLeaveMin:  DEFAULT_EARLY_LEAVE_MIN,
    weeklyHoursTarget: DEFAULT_WEEKLY_HOURS_TARGET,
    companyName:    'Goolee'
  };

  if (sheet && sheet.getLastRow() >= 2) {
    const lastRow = sheet.getLastRow();
    const values = sheet.getRange(1, 1, lastRow, 3).getValues();
    for (let i = 1; i < values.length; i++) {
      const key = String(values[i][0] || '').trim().toLowerCase();
      const val = String(values[i][1] || '').trim();
      if (!key || !val) continue;

      if (key === 'working days') {
        out.workingDays = parseWorkingDays_(val);
      } else if (key === 'saturday working day') {
        out.saturdayWorkingDay = /^(true|yes|1)$/i.test(val);
      } else if (key === 'work start time') {
        const p = val.split(':');
        const h = Number(p[0]), m = Number(p[1] || 0);
        if (!isNaN(h)) out.workStartHour = h;
        if (!isNaN(m)) out.workStartMin = m;
      } else if (key === 'work end time') {
        const p = val.split(':');
        const h = Number(p[0]), m = Number(p[1] || 0);
        if (!isNaN(h)) out.workEndHour = h;
        if (!isNaN(m)) out.workEndMin = m;
      } else if (key === 'saturday work start time') {
        const p = val.split(':');
        const h = Number(p[0]), m = Number(p[1] || 0);
        if (!isNaN(h)) out.saturdayStartHour = h;
        if (!isNaN(m)) out.saturdayStartMin = m;
      } else if (key === 'saturday work end time') {
        const p = val.split(':');
        const h = Number(p[0]), m = Number(p[1] || 0);
        if (!isNaN(h)) out.saturdayEndHour = h;
        if (!isNaN(m)) out.saturdayEndMin = m;
      } else if (key === 'late after (min)') {
        const n = Number(val); if (!isNaN(n)) out.lateAfterMin = n;
      } else if (key === 'early leave grace (min)') {
        const n = Number(val); if (!isNaN(n)) out.earlyLeaveMin = n;
      } else if (key === 'minimum weekly hours' || key === 'weekly hours target') {
        const n = Number(val); if (!isNaN(n) && n > 0) out.weeklyHoursTarget = n;
      } else if (key === 'company name') {
        out.companyName = val;
      }
    }
  }

  if (out.saturdayWorkingDay) out.workingDays.add('Sat');
  if (!out.weeklyHoursTarget || out.weeklyHoursTarget <= 0) {
    out.weeklyHoursTarget = DEFAULT_WEEKLY_HOURS_TARGET;
  }

  const serializable = Object.assign({}, out, {
    workingDays: Array.from(out.workingDays)
  });
  try {
    cache.put(CACHE_SETTINGS_KEY, JSON.stringify(serializable), CACHE_CONFIG_TTL);
  } catch (e) {}
  return out;
}

function parseWorkingDays_(str) {
  const dayNames = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const set = new Set();
  const s = String(str || '').trim();
  if (!s) return new Set(['Mon','Tue','Wed','Thu','Fri']);

  const rangeMatch = s.match(/^(\w{3})\s*-\s*(\w{3})$/i);
  if (rangeMatch) {
    const startIdx = dayNames.findIndex(d => d.toLowerCase() === rangeMatch[1].toLowerCase());
    const endIdx   = dayNames.findIndex(d => d.toLowerCase() === rangeMatch[2].toLowerCase());
    if (startIdx >= 0 && endIdx >= 0) {
      let i = startIdx, guard = 0;
      while (guard < 14) {
        set.add(dayNames[i]);
        if (i === endIdx) break;
        i = (i + 1) % 7;
        guard++;
      }
      return set;
    }
  }

  s.split(',').forEach(part => {
    const p = part.trim();
    const idx = dayNames.findIndex(d => d.toLowerCase() === p.toLowerCase());
    if (idx >= 0) set.add(dayNames[idx]);
  });
  return set.size > 0 ? set : new Set(['Mon','Tue','Wed','Thu','Fri']);
}

// ════════════════════════════════════════════════════════════════════
//  HOLIDAYS READER
// ════════════════════════════════════════════════════════════════════
function readHolidays_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_HOLIDAYS_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_HOLIDAYS);
  if (!sheet || sheet.getLastRow() < 2) {
    try { cache.put(CACHE_HOLIDAYS_KEY, '[]', CACHE_CONFIG_TTL); } catch (e) {}
    return [];
  }

  const tz = ss.getSpreadsheetTimeZone();
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, 3).getValues();

  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const dateVal = row[0];
    const name = String(row[1] || '').trim();
    const type = String(row[2] || '').trim() || 'Public Holiday';
    if (!dateVal) continue;

    let dateStr;
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
    } else {
      const s = String(dateVal).trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) dateStr = s.substring(0, 10);
      else continue;
    }
    out.push({ date: dateStr, name: name || type, type: type });
  }
  try { cache.put(CACHE_HOLIDAYS_KEY, JSON.stringify(out), CACHE_CONFIG_TTL); } catch (e) {}
  return out;
}

// ════════════════════════════════════════════════════════════════════
//  CONTEXT BUILDER
// ════════════════════════════════════════════════════════════════════
function buildContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  const empSheet = ss.getSheetByName(SHEET_EMPLOYEES);
  if (!attSheet || !empSheet) {
    SpreadsheetApp.getUi().alert('Missing Attendance or Employees sheet.');
    return null;
  }

  const tz = ss.getSpreadsheetTimeZone();
  const settings = readSettings_();
  const holidays = readHolidays_();
  const holidaySet = new Set();
  holidays.forEach(h => holidaySet.add(h.date));

  // Active employees
  const empVals = empSheet.getDataRange().getValues();
  const activeEmployees = [];
  for (let i = 1; i < empVals.length; i++) {
    const name = String(empVals[i][0] || '').trim();
    if (!name) continue;
    const active = empVals[i][2];
    const isActive = active === true || String(active).trim().toUpperCase() === 'TRUE';
    if (isActive) activeEmployees.push(name);
  }

  // Punches
  const attLastRow = attSheet.getLastRow();
  const attLastCol = attSheet.getLastColumn();
  const punches = [];
  if (attLastRow >= 2) {
    const attVals = attSheet.getRange(1, 1, attLastRow, attLastCol).getValues();
    for (let i = 1; i < attVals.length; i++) {
      const row = attVals[i];
      const ts = row[0];
      const name = String(row[1] || '').trim();
      const action = String(row[2] || '').trim();
      if (!(ts instanceof Date) || !name || !action) continue;
      punches.push({ ts: ts, name: name, action: action });
    }
  }

  const leaves = readLeaves_();

  // Create daily records from punches and leave. Sessions are paired globally
  // by employee so a Clock In before midnight can close after midnight.
  const dailyMap = {};
  punches.forEach(function (p) {
    const dateStr = Utilities.formatDate(p.ts, tz, 'yyyy-MM-dd');
    const key = p.name + '|' + dateStr;
    if (!dailyMap[key]) dailyMap[key] = createDailyRecord_(p.name, dateStr);
  });

  leaves.forEach(function (lv) {
    const key = lv.name + '|' + lv.date;
    if (!dailyMap[key]) dailyMap[key] = createDailyRecord_(lv.name, lv.date);
    dailyMap[key].leave = lv.type;
    dailyMap[key].leaveNote = lv.notes || '';
  });

  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const sessionsByEmployee = buildSessionsByEmployee_(punches);

  Object.keys(sessionsByEmployee).forEach(function (employeeName) {
    sessionsByEmployee[employeeName].forEach(function (session) {
      const end = session.clockOut || now;
      if (!end || end <= session.clockIn) return;

      let cursor = session.clockIn;
      while (cursor < end) {
        const dateStr = Utilities.formatDate(cursor, tz, 'yyyy-MM-dd');
        const nextDay = startOfDateInTimeZone_(dateStringAddDays_(dateStr, 1), tz);
        const segmentEnd = end < nextDay ? end : nextDay;
        const key = employeeName + '|' + dateStr;
        if (!dailyMap[key]) dailyMap[key] = createDailyRecord_(employeeName, dateStr);

        const rec = dailyMap[key];
        const segment = {
          clockIn: session.clockIn,
          clockOut: session.clockOut,
          hours: (segmentEnd - cursor) / 3600000
        };
        rec.sessions.push(segment);
        rec.hours += segment.hours;
        if (dateStr === Utilities.formatDate(session.clockIn, tz, 'yyyy-MM-dd')) {
          rec.sessionCount++;
        }
        if (!rec.clockIn || session.clockIn < rec.clockIn) rec.clockIn = session.clockIn;
        if (session.clockOut && (!rec.clockOut || session.clockOut > rec.clockOut)) {
          rec.clockOut = session.clockOut;
        }
        if (!session.clockOut && segmentEnd >= end) rec.openSession = segment;

        cursor = segmentEnd;
      }
    });
  });

  const dailyList = [];
  for (const k in dailyMap) {
    const rec = dailyMap[k];
    rec.isHoliday = holidaySet.has(rec.date);
    dailyList.push({
      name: rec.name, date: rec.date,
      clockIn: rec.clockIn, clockOut: rec.clockOut,
      sessions: rec.sessions, sessionCount: rec.sessionCount,
      openSession: rec.openSession, hours: rec.hours,
      leave: rec.leave, leaveNote: rec.leaveNote,
      isHoliday: rec.isHoliday
    });
  }

  dailyList.sort(function (a, b) {
    if (a.date !== b.date) return b.date.localeCompare(a.date);
    return a.name.localeCompare(b.name);
  });

  // Today's KPIs
  const todayRecords = dailyList.filter(r => r.date === todayStr);

  const todayIsHoliday = holidaySet.has(todayStr);
  const onlineToday = todayRecords.filter(r => r.openSession).length;
  const leaveToday   = todayRecords.filter(r => r.leave && !r.clockIn).length;
  const offlineToday = Math.max(activeEmployees.length - onlineToday - leaveToday, 0);

  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const monthStartStr = Utilities.formatDate(monthStart, tz, 'yyyy-MM-dd');
  const monthEndStr = Utilities.formatDate(
    new Date(now.getFullYear(), now.getMonth() + 1, 0), tz, 'yyyy-MM-dd');
  const holidaysThisMonth = holidays.filter(
    h => h.date >= monthStartStr && h.date <= monthEndStr);

  // Weekly totals: Monday–Sunday, built from every completed session.
  const currentWeekStart = weekStartStr_(todayStr);
  const weeklyMap = {};
  dailyList.forEach(function (r) {
    const weekStart = weekStartStr_(r.date);
    const key = r.name + '|' + weekStart;
    if (!weeklyMap[key]) {
      weeklyMap[key] = {
        name: r.name, weekStart: weekStart,
        weekEnd: weekEndStr_(weekStart),
        totalHours: 0, daysPresent: 0, daysLeave: 0,
        sessionCount: 0, openSessions: 0
      };
    }
    const w = weeklyMap[key];
    w.totalHours += r.hours;
    w.sessionCount += r.sessionCount || 0;
    if (r.openSession) w.openSessions++;
    if (r.clockIn) w.daysPresent++;
    if (r.leave) w.daysLeave++;
  });

  const weeklyHistory = Object.keys(weeklyMap).map(function (key) {
    return weeklyMap[key];
  }).sort(function (a, b) {
    if (a.weekStart !== b.weekStart) return b.weekStart.localeCompare(a.weekStart);
    return a.name.localeCompare(b.name);
  });

  // Per-employee current-week monitoring
  const byEmp = {};
  activeEmployees.forEach(function (name) {
    byEmp[name] = {
      name: name, daysPresent: 0, daysLeave: 0,
      sessionCount: 0, openSessions: 0, totalHours: 0,
      weekStart: currentWeekStart,
      lastTs: null, lastAction: ''
    };
  });

  dailyList.filter(function (r) { return weekStartStr_(r.date) === currentWeekStart; })
  .forEach(function (r) {
    if (!byEmp[r.name]) {
      byEmp[r.name] = {
        name: r.name, daysPresent: 0, daysLeave: 0,
        sessionCount: 0, openSessions: 0, totalHours: 0,
        weekStart: currentWeekStart, lastTs: null, lastAction: ''
      };
    }
    const e = byEmp[r.name];
    if (r.leave) e.daysLeave++;
    if (r.clockIn) e.daysPresent++;
    e.sessionCount += r.sessionCount || 0;
    if (r.openSession) e.openSessions++;
    e.totalHours += r.hours;
  });

  punches.forEach(function (p) {
    const e = byEmp[p.name];
    if (!e) return;
    if (!e.lastTs || p.ts > e.lastTs) {
      e.lastTs = p.ts;
      e.lastAction = p.action;
    }
  });

  const employeeRows = activeEmployees.map(n => byEmp[n]).filter(Boolean);
  const weeklyTarget = Number(settings.weeklyHoursTarget) || DEFAULT_WEEKLY_HOURS_TARGET;
  const weeklyCompletedCount = employeeRows.filter(function (employee) {
    return (Number(employee.totalHours) || 0) >= weeklyTarget;
  }).length;
  const weeklyAchievementPct = employeeRows.length > 0
    ? Math.round((weeklyCompletedCount / employeeRows.length) * 100) : 0;

  // Today per employee
  const todayByEmp = {};
  activeEmployees.forEach(function (name) {
    todayByEmp[name] = {
      name: name,
      clockIn: null, clockOut: null, hours: 0,
      sessions: [], sessionCount: 0, openSession: null,
      leave: null, leaveNote: ''
    };
  });
  todayRecords.forEach(function (r) {
    if (todayByEmp[r.name]) {
      todayByEmp[r.name].clockIn     = r.clockIn;
      todayByEmp[r.name].clockOut    = r.clockOut;
      todayByEmp[r.name].hours       = r.hours;
      todayByEmp[r.name].sessions    = r.sessions || [];
      todayByEmp[r.name].sessionCount = r.sessionCount || 0;
      todayByEmp[r.name].openSession = r.openSession || null;
      todayByEmp[r.name].leave       = r.leave;
      todayByEmp[r.name].leaveNote   = r.leaveNote;
    }
  });
  const todayPerf = activeEmployees.map(n => todayByEmp[n]);
  const onLeaveToday = todayPerf.filter(t => t.leave && !t.clockIn);

  // Daily logs
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - LOG_DAYS_BACK);
  const cutoffStr = Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd');
  const recentDaily = dailyList.filter(r => r.date >= cutoffStr);

  const empCutoff = new Date();
  empCutoff.setDate(empCutoff.getDate() - EMP_LOG_DAYS_BACK);
  const empCutoffStr = Utilities.formatDate(empCutoff, tz, 'yyyy-MM-dd');
  const empDaily = dailyList.filter(r => r.date >= empCutoffStr);

  return {
    tz: tz, now: now, todayStr: todayStr,
    settings: settings, holidays: holidays,
    holidaysThisMonth: holidaysThisMonth,
    todayIsHoliday: todayIsHoliday,
    onlineToday: onlineToday,
    offlineToday: offlineToday,
    leaveToday: leaveToday,
    weeklyCompletedCount: weeklyCompletedCount,
    weeklyAchievementPct: weeklyAchievementPct,
    activeCount: activeEmployees.length,
    activeEmployees: activeEmployees,
    todayPerf: todayPerf,
    onLeaveToday: onLeaveToday,
    employeeRows: employeeRows,
    currentWeekStart: currentWeekStart,
    currentWeekEnd: weekEndStr_(currentWeekStart),
    weeklyHistory: weeklyHistory,
    recentDaily: recentDaily,
    empDaily: empDaily,
    byEmp: byEmp
  };
}

function createDailyRecord_(name, dateStr) {
  return {
    name: name,
    date: dateStr,
    sessions: [],
    sessionCount: 0,
    clockIn: null,
    clockOut: null,
    openSession: null,
    hours: 0,
    leave: null,
    leaveNote: '',
    isHoliday: false
  };
}

function buildSessionsByEmployee_(punches) {
  const grouped = {};
  punches.forEach(function (p) {
    if (!grouped[p.name]) grouped[p.name] = [];
    grouped[p.name].push(p);
  });

  const sessionsByEmployee = {};
  Object.keys(grouped).forEach(function (name) {
    const events = grouped[name].slice().sort(function (a, b) { return a.ts - b.ts; });
    const sessions = [];
    let openIn = null;

    events.forEach(function (p) {
      if (p.action === 'Clock In') {
        if (openIn) sessions.push({ clockIn: openIn, clockOut: null });
        openIn = p.ts;
      } else if (p.action === 'Clock Out' && openIn) {
        sessions.push({ clockIn: openIn, clockOut: p.ts });
        openIn = null;
      }
    });

    if (openIn) sessions.push({ clockIn: openIn, clockOut: null });
    sessionsByEmployee[name] = sessions;
  });

  return sessionsByEmployee;
}

function dateStringAddDays_(dateStr, days) {
  const p = String(dateStr).split('-').map(Number);
  const date = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  date.setUTCDate(date.getUTCDate() + days);
  return date.getUTCFullYear() + '-' + pad2_(date.getUTCMonth() + 1) + '-' + pad2_(date.getUTCDate());
}

function startOfDateInTimeZone_(dateStr, tz) {
  const p = String(dateStr).split('-').map(Number);
  const utcNoon = new Date(Date.UTC(p[0], p[1] - 1, p[2], 12, 0, 0));
  const zone = Utilities.formatDate(utcNoon, tz, 'Z');
  const match = String(zone).match(/^([+-])(\d{2})(\d{2})$/);
  const offsetMinutes = match
    ? (Number(match[2]) * 60 + Number(match[3])) * (match[1] === '-' ? -1 : 1)
    : 0;
  return new Date(Date.UTC(p[0], p[1] - 1, p[2]) - offsetMinutes * 60000);
}

function countWorkingDays_(year, month, workingDaysSet, holidaySet, tz, upToDate) {
  let count = 0;
  const lastDay = upToDate ? upToDate.getDate() : new Date(year, month + 1, 0).getDate();
  const dayNames = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  for (let d = 1; d <= lastDay; d++) {
    const date = new Date(year, month, d);
    const dow = dayNames[date.getDay()];
    const dateStr = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
    if (workingDaysSet.has(dow) && !holidaySet.has(dateStr)) count++;
  }
  return count;
}

function weekStartStr_(dateStr) {
  const p = String(dateStr).split('-').map(Number);
  const utc = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  const mondayOffset = (utc.getUTCDay() + 6) % 7;
  utc.setUTCDate(utc.getUTCDate() - mondayOffset);
  return utc.getUTCFullYear() + '-' + pad2_(utc.getUTCMonth() + 1) + '-' + pad2_(utc.getUTCDate());
}

function weekEndStr_(weekStartStr) {
  const p = String(weekStartStr).split('-').map(Number);
  const utc = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  utc.setUTCDate(utc.getUTCDate() + 6);
  return utc.getUTCFullYear() + '-' + pad2_(utc.getUTCMonth() + 1) + '-' + pad2_(utc.getUTCDate());
}

function weekLabel_(weekStart, weekEnd, tz) {
  // Noon avoids a date rollover when Apps Script and spreadsheet time zones differ.
  const start = new Date(weekStart + 'T12:00:00');
  const end = new Date(weekEnd + 'T12:00:00');
  return Utilities.formatDate(start, tz, 'd MMM') + ' – ' +
    Utilities.formatDate(end, tz, 'd MMM yyyy');
}

function readLeaves_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LEAVE);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const tz = ss.getSpreadsheetTimeZone();
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, 4).getValues();

  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const dateVal = row[0];
    const name    = String(row[1] || '').trim();
    const type    = String(row[2] || '').trim();
    const notes   = String(row[3] || '').trim();

    if (!dateVal || !name || !type) continue;
    if (LEAVE_TYPES.indexOf(type) === -1) continue;

    let dateStr;
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
    } else {
      const s = String(dateVal).trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) dateStr = s.substring(0, 10);
      else continue;
    }
    out.push({ date: dateStr, name: name, type: type, notes: notes });
  }
  return out;
}

// ════════════════════════════════════════════════════════════════════
//  FORMATTERS
// ════════════════════════════════════════════════════════════════════
function pad2_(n) { return (n < 10 ? '0' : '') + n; }

function formatHoursMinutes_(hours) {
  const total = Math.round(hours * 60);
  const h = Math.floor(total / 60);
  const m = total % 60;
  if (h === 0) return m + 'm';
  if (m === 0) return h + 'h';
  return h + 'h ' + m + 'm';
}

function formatSessionTimes_(sessions, tz, field) {
  if (!sessions || !sessions.length) return '—';
  const values = sessions.map(function (session) {
    return session[field] ? Utilities.formatDate(session[field], tz, 'h:mm a') : null;
  }).filter(Boolean);
  return values.length ? values.join(' · ') : '—';
}

// ════════════════════════════════════════════════════════════════════
//  TODAY'S STATUS BUILDER
// ════════════════════════════════════════════════════════════════════
function dailyStatusFor_(rec) {
  if (rec.isHoliday) return { text: HOLIDAY_COLOR.icon + ' Holiday', colors: HOLIDAY_COLOR };
  if (rec.leave) {
    const col = leaveColor_(rec.leave);
    return { text: col.icon + ' ' + rec.leave, colors: col };
  }
  if (rec.openSession) {
    return {
      text: '⏱️ Working · ' + formatHoursMinutes_(rec.hours),
      colors: { bg: C.greenLight, fg: C.green }
    };
  }
  if (rec.clockOut) {
    return {
      text: '✅ Completed · ' + formatHoursMinutes_(rec.hours) +
        (rec.sessionCount > 1 ? ' · ' + rec.sessionCount + ' sessions' : ''),
      colors: { bg: C.greenLight, fg: C.green }
    };
  }
  if (rec.clockIn) {
    return { text: '🟢 Punch recorded', colors: { bg: C.blueLight, fg: C.blue } };
  }
  return { text: '—', colors: { bg: C.grayLight, fg: C.gray } };
}

function todayStatusFor_(emp, ctx) {
  if (ctx.todayIsHoliday) {
    return { text: '🎉 Holiday', bg: HOLIDAY_COLOR.bg, fg: HOLIDAY_COLOR.fg };
  }
  if (emp.leave) {
    const col = leaveColor_(emp.leave);
    let text = col.icon + ' ' + emp.leave;
    if (emp.openSession) {
      text += ' · ⏱️ Working';
    } else if (emp.clockIn && emp.clockOut) {
      text += ' · 🏁 ' + formatHoursMinutes_(emp.hours) + ' worked';
    }
    return { text: text, bg: col.bg, fg: col.fg };
  }
  if (!emp.clockIn) {
    return { text: '⏳ Not yet in', bg: C.grayLight, fg: C.gray };
  }
  if (emp.openSession) {
    return {
      text: '⏱️ Working · ' + formatHoursMinutes_(emp.hours),
      bg: C.greenLight, fg: C.green
    };
  }
  if (!emp.clockOut) {
    return { text: '🟢 Punch recorded', bg: C.blueLight, fg: C.blue };
  }
  return {
    text: '🏁 Completed · ' + formatHoursMinutes_(emp.hours) +
      (emp.sessionCount > 1 ? ' · ' + emp.sessionCount + ' sessions' : ''),
    bg: C.greenLight, fg: C.green
  };
}

// ════════════════════════════════════════════════════════════════════
//  DASHBOARD RESET
// ════════════════════════════════════════════════════════════════════
function resetDashboard_(dash) {
  const maxR = Math.max(dash.getMaxRows(), 200);
  const maxC = Math.max(dash.getMaxColumns(), 30);

  try { dash.getRange(1, 1, maxR, maxC).breakApart(); } catch (e) {
    console.error('breakApart failed:', e);
  }

  dash.getRange(1, 1, maxR, maxC).clear();

  const widths = [190, 105, 100, 100, 100, 110, 120, 120, 120, 130, 130, 120];
  widths.forEach(function (w, i) { dash.setColumnWidth(i + 1, w); });

  try {
    dash.showColumns(1, DASH_TOTAL_COLS);
    if (dash.getMaxColumns() > DASH_TOTAL_COLS) {
      dash.hideColumns(DASH_TOTAL_COLS + 1, dash.getMaxColumns() - DASH_TOTAL_COLS);
    }
  } catch (e) {
    console.error('Column hide failed:', e);
  }
}

// ════════════════════════════════════════════════════════════════════
//  MAIN DASHBOARD PAINTER
// ════════════════════════════════════════════════════════════════════
function paintDashboard_(dash, ctx) {
  dash.getCharts().forEach(function (ch) { dash.removeChart(ch); });
  dash.clearConditionalFormatRules();

  const requiredRows = 250 + ctx.weeklyHistory.length;
  if (dash.getMaxRows() < requiredRows) {
    dash.insertRowsAfter(dash.getMaxRows(), requiredRows - dash.getMaxRows());
  }

  resetDashboard_(dash);

  const s = ctx.settings;

  let r = 1;

  // Title
  dash.setRowHeight(r, 72);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
    .setValue('🕐  ' + s.companyName + '.my — Attendance Dashboard')
    .setBackground(C.tealDark).setFontColor(C.white)
    .setFontSize(22).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  // Subtitle
  dash.setRowHeight(r, 30);
  const stamp = Utilities.formatDate(ctx.now, ctx.tz, 'EEEE, d MMM yyyy · h:mm a');
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
    .setValue('Last refreshed: ' + stamp + '  ·  Refreshes automatically every day')
    .setFontColor(C.gray).setFontSize(11).setFontStyle('italic')
    .setHorizontalAlignment('right').setVerticalAlignment('middle')
    .setBackground(C.grayLight);
  r++;

  dash.setRowHeight(r, 14);
  r++;

  // KPI header
  dash.setRowHeight(r, 40);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
    .setValue("📊  TODAY'S SNAPSHOT" + (ctx.todayIsHoliday ? '  ·  🎉 Holiday today' : ''))
    .setBackground(C.teal).setFontColor(C.white)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  const kpiDefs = [
    { label: '📅  DATE', value: Utilities.formatDate(ctx.now, ctx.tz, 'EEE, d MMM'),
      bg: C.grayLight, fg: C.gray },
    { label: '🟢  ONLINE', value: ctx.onlineToday + ' / ' + ctx.activeCount,
      bg: C.greenLight, fg: C.green },
    { label: '⚪  OFFLINE', value: ctx.offlineToday + ' / ' + ctx.activeCount,
      bg: C.grayLight, fg: C.gray },
    { label: '✅  PEOPLE COMPLETE', value: ctx.weeklyCompletedCount + ' / ' + ctx.activeCount,
      bg: ctx.weeklyCompletedCount > 0 ? C.greenLight : C.grayLight,
      fg: ctx.weeklyCompletedCount > 0 ? C.green : C.gray },
    { label: '🌴  ON LEAVE', value: String(ctx.leaveToday),
      bg: ctx.leaveToday > 0 ? C.tealLight : C.grayLight,
      fg: ctx.leaveToday > 0 ? C.teal : C.gray },
    { label: '🎯  ACHIEVED', value: ctx.weeklyAchievementPct + '%',
      bg: ctx.weeklyAchievementPct >= 100 ? C.greenLight : C.amberLight,
      fg: ctx.weeklyAchievementPct >= 100 ? C.green : C.amber }
  ];

  dash.setRowHeight(r, 26);
  const labelRow = r;
  kpiDefs.forEach(function (kpi, i) {
    const cs = i * 2 + 1;
    dash.getRange(labelRow, cs, 1, 2).merge()
      .setValue(kpi.label)
      .setBackground(kpi.bg).setFontColor(kpi.fg)
      .setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  r++;

  dash.setRowHeight(r, 62);
  const valueRow = r;
  kpiDefs.forEach(function (kpi, i) {
    const cs = i * 2 + 1;
    dash.getRange(valueRow, cs, 1, 2).merge()
      .setValue(kpi.value)
      .setBackground(kpi.bg).setFontColor(C.ink)
      .setFontSize(20).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    dash.getRange(labelRow, cs, 2, 2)
      .setBorder(true, true, true, true, true, true, C.border, SpreadsheetApp.BorderStyle.SOLID);
  });
  r++;

  dash.setRowHeight(r, 14);
  r++;

  // TODAY'S PERFORMANCE
  dash.setRowHeight(r, 40);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
    .setValue("🌟  TODAY'S PERFORMANCE — " +
              Utilities.formatDate(ctx.now, ctx.tz, 'EEEE, d MMM yyyy'))
    .setBackground(C.teal).setFontColor(C.white)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  dash.setRowHeight(r, 32);
  dash.getRange(r, 1, 1, 6).setValues([[
    'EMPLOYEE', 'CLOCK IN', 'CLOCK OUT', 'HOURS', 'LEAVE', 'STATUS'
  ]]);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
    .setBackground(C.grayLight)
    .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(r, 1, 1, 6)
    .setFontColor(C.tealDark).setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  r++;

  if (ctx.todayPerf.length === 0) {
    dash.setRowHeight(r, 34);
    dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
      .setValue('No active employees.')
      .setFontColor(C.gray).setFontStyle('italic')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground(C.grayLight)
      .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
    r++;
  } else {
    ctx.todayPerf.forEach(function (emp, i) {
      dash.setRowHeight(r, 34);

       const ci = formatSessionTimes_(emp.sessions, ctx.tz, 'clockIn');
       const co = formatSessionTimes_(emp.sessions, ctx.tz, 'clockOut');
      const hoursLabel = emp.hours > 0 ? formatHoursMinutes_(emp.hours) : '—';
      const leaveLabel = emp.leave
        ? leaveColor_(emp.leave).icon + ' ' + emp.leave : '—';
      const status = todayStatusFor_(emp, ctx);

      dash.getRange(r, 1, 1, 6).setValues([[
        emp.name, ci, co, hoursLabel, leaveLabel, status.text
      ]]);

      const altBg = (i % 2 === 0) ? C.white : C.grayLight;
      dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
        .setBackground(altBg).setFontSize(11)
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);

      dash.getRange(r, 1).setHorizontalAlignment('left')
        .setFontWeight('bold').setFontColor(C.tealDark);
      dash.getRange(r, 2, 1, 3).setHorizontalAlignment('center');

      if (emp.leave) {
        const col = leaveColor_(emp.leave);
        dash.getRange(r, 5)
          .setBackground(col.bg).setFontColor(col.fg)
          .setFontWeight('bold').setHorizontalAlignment('center');
      } else {
        dash.getRange(r, 5).setFontColor(C.gray).setHorizontalAlignment('center');
      }

      dash.getRange(r, 6)
        .setBackground(status.bg).setFontColor(status.fg)
        .setFontWeight('bold').setHorizontalAlignment('center');

      r++;
    });
  }

  dash.setRowHeight(r, 14);
  r++;

  // ON LEAVE TODAY
  if (ctx.onLeaveToday.length > 0) {
    dash.setRowHeight(r, 40);
    dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
      .setValue('🏖️  ON LEAVE TODAY — ' + ctx.onLeaveToday.length + ' employee' +
                (ctx.onLeaveToday.length === 1 ? '' : 's'))
      .setBackground(C.tealLight).setFontColor(C.tealDark)
      .setFontSize(14).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    r++;

    dash.setRowHeight(r, 32);
    dash.getRange(r, 1, 1, 3).setValues([['EMPLOYEE', 'LEAVE TYPE', 'NOTES']]);
    dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
      .setBackground(C.grayLight)
      .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(r, 1, 1, 3)
      .setFontColor(C.tealDark).setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    r++;

    ctx.onLeaveToday.forEach(function (emp, i) {
      dash.setRowHeight(r, 32);
      const col = leaveColor_(emp.leave);

      dash.getRange(r, 1, 1, 3).setValues([[
        emp.name,
        col.icon + ' ' + emp.leave,
        emp.leaveNote || ''
      ]]);

      dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
        .setBackground(col.bg).setFontSize(11)
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);

      dash.getRange(r, 1).setFontWeight('bold').setFontColor(C.tealDark)
        .setHorizontalAlignment('left');
      dash.getRange(r, 2).setFontColor(col.fg).setFontWeight('bold')
        .setHorizontalAlignment('center');
      dash.getRange(r, 3).setFontColor(C.gray).setFontSize(10)
        .setHorizontalAlignment('left');
      r++;
    });

    dash.setRowHeight(r, 14);
    r++;
  }

  // EMPLOYEE PERFORMANCE — THIS WEEK
  dash.setRowHeight(r, 40);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
    .setValue('🏢  EMPLOYEE WEEKLY MONITORING — ' +
      weekLabel_(ctx.currentWeekStart, ctx.currentWeekEnd, ctx.tz))
    .setBackground(C.teal).setFontColor(C.white)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  dash.setRowHeight(r, 34);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).setValues([[
    'EMPLOYEE', 'HOURS THIS WEEK', 'TARGET', 'REMAINING', 'SESSIONS',
    'DAYS WORKED', 'OPEN SESSIONS', 'STATUS', 'LAST ACTION', 'LAST PUNCH', '', ''
  ]]);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
    .setBackground(C.grayLight).setFontColor(C.tealDark)
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
  r++;

  if (ctx.employeeRows.length === 0) {
    dash.setRowHeight(r, 34);
    dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
      .setValue('No active employees.')
      .setFontColor(C.gray).setFontStyle('italic')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground(C.grayLight)
      .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
    r++;
  } else {
    ctx.employeeRows.forEach(function (emp, i) {
      dash.setRowHeight(r, 34);

      const totalHours = Number(emp.totalHours) || 0;
      const targetHours = Number(s.weeklyHoursTarget) || DEFAULT_WEEKLY_HOURS_TARGET;
      const remainingHours = Math.max(targetHours - totalHours, 0);
      const statusText = totalHours >= targetHours
        ? '✅ Target met'
        : '⏳ ' + formatHoursMinutes_(remainingHours) + ' remaining';
      const lastAction = emp.lastAction || '—';
      const lastPunch = emp.lastTs
        ? Utilities.formatDate(emp.lastTs, ctx.tz, 'd MMM · h:mm a')
        : '—';

      dash.getRange(r, 1, 1, DASH_TOTAL_COLS).setValues([[
        emp.name, formatHoursMinutes_(totalHours), targetHours + ' h',
        formatHoursMinutes_(remainingHours), emp.sessionCount || 0,
        emp.daysPresent || 0, emp.openSessions || 0, statusText,
        lastAction, lastPunch, '', ''
      ]]);

      const altBg = (i % 2 === 0) ? C.white : C.grayLight;
      dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
        .setBackground(altBg).setFontSize(11)
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);

      dash.getRange(r, 1).setHorizontalAlignment('left')
        .setFontWeight('bold').setFontColor(C.tealDark);
      dash.getRange(r, 2, 1, 7).setHorizontalAlignment('center');
      dash.getRange(r, 9, 1, 2).setFontColor(C.gray).setFontSize(10)
        .setHorizontalAlignment('center');

      dash.getRange(r, 8).setBackground(totalHours >= targetHours ? C.greenLight : C.amberLight)
        .setFontColor(totalHours >= targetHours ? C.green : C.amber)
        .setFontWeight('bold').setHorizontalAlignment('center');

      r++;
    });
  }

  dash.setRowHeight(r, 14);
  r++;

  // WEEKLY HISTORY
  dash.setRowHeight(r, 40);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
    .setValue('📈  WEEKLY HOURS HISTORY')
    .setBackground(C.blue).setFontColor(C.white)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  dash.setRowHeight(r, 34);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).setValues([[
    'WEEK', 'EMPLOYEE', 'TOTAL HOURS', 'TARGET', 'REMAINING', 'SESSIONS',
    'DAYS WORKED', 'OPEN SESSIONS', 'STATUS', '', '', ''
  ]]);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
    .setBackground(C.grayLight).setFontColor(C.tealDark)
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
  r++;

  ctx.weeklyHistory.forEach(function (week, i) {
    const target = Number(s.weeklyHoursTarget) || DEFAULT_WEEKLY_HOURS_TARGET;
    const remaining = Math.max(target - week.totalHours, 0);
    const met = week.totalHours >= target;
    dash.setRowHeight(r, 32);
    dash.getRange(r, 1, 1, DASH_TOTAL_COLS).setValues([[
      weekLabel_(week.weekStart, week.weekEnd, ctx.tz), week.name,
      formatHoursMinutes_(week.totalHours), target + ' h',
      formatHoursMinutes_(remaining), week.sessionCount, week.daysPresent,
      week.openSessions, met ? '✅ Target met' : '⏳ ' + formatHoursMinutes_(remaining) + ' remaining',
      '', '', ''
    ]]);
    dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
      .setBackground(i % 2 === 0 ? C.white : C.grayLight).setFontSize(11)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(r, 1).setFontColor(C.blue).setFontWeight('bold').setHorizontalAlignment('center');
    dash.getRange(r, 2).setFontColor(C.tealDark).setFontWeight('bold');
    dash.getRange(r, 3, 1, 6).setHorizontalAlignment('center');
    dash.getRange(r, 9).setBackground(met ? C.greenLight : C.amberLight)
      .setFontColor(met ? C.green : C.amber).setFontWeight('bold')
      .setHorizontalAlignment('center');
    r++;
  });

  dash.setRowHeight(r, 14);
  r++;

  // HOLIDAYS THIS MONTH
  if (ctx.holidaysThisMonth.length > 0) {
    dash.setRowHeight(r, 40);
    dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
      .setValue('🎉  HOLIDAYS THIS MONTH')
      .setBackground(C.purple).setFontColor(C.white)
      .setFontSize(14).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    r++;

    dash.setRowHeight(r, 32);
    dash.getRange(r, 1, 1, 3).setValues([['DATE', 'HOLIDAY NAME', 'TYPE']]);
    dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
      .setBackground(C.grayLight)
      .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(r, 1, 1, 3)
      .setFontColor(C.tealDark).setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    r++;

    ctx.holidaysThisMonth.forEach(function (h, i) {
      dash.setRowHeight(r, 30);
      const d = new Date(h.date + 'T00:00:00');
      dash.getRange(r, 1, 1, 3).setValues([[
        Utilities.formatDate(d, ctx.tz, 'EEE, d MMM yyyy'), h.name, h.type
      ]]);
      dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
        .setBackground(C.purpleLight).setFontSize(11)
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
      dash.getRange(r, 1).setFontColor(C.purple).setFontWeight('bold').setHorizontalAlignment('center');
      dash.getRange(r, 2).setFontColor(C.ink);
      dash.getRange(r, 3).setFontColor(C.gray).setFontSize(10).setHorizontalAlignment('center');
      r++;
    });

    dash.setRowHeight(r, 14);
    r++;
  }

  // DAILY LOG
  dash.setRowHeight(r, 40);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
    .setValue('📅  DAILY LOG — LAST ' + LOG_DAYS_BACK + ' DAYS')
    .setBackground(C.teal).setFontColor(C.white)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  dash.setRowHeight(r, 34);
  dash.getRange(r, 1, 1, 6).setValues([[
    'DATE', 'EMPLOYEE', 'CLOCK IN', 'CLOCK OUT', 'HOURS', 'STATUS'
  ]]);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
    .setBackground(C.grayLight)
    .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(r, 1, 1, 6)
    .setFontColor(C.tealDark).setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  r++;

  if (ctx.recentDaily.length === 0) {
    dash.setRowHeight(r, 34);
    dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
      .setValue('No punches or leave in the last ' + LOG_DAYS_BACK + ' days.')
      .setFontColor(C.gray).setFontStyle('italic')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground(C.grayLight)
      .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
    r++;
  } else {
    ctx.recentDaily.forEach(function (rec, i) {
      dash.setRowHeight(r, 32);

      const parts = rec.date.split('-');
      const d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
      const dateLabel = Utilities.formatDate(d, ctx.tz, 'EEE, d MMM');
       const ci = formatSessionTimes_(rec.sessions, ctx.tz, 'clockIn');
       const co = formatSessionTimes_(rec.sessions, ctx.tz, 'clockOut');
      let hoursLabel = '—';
      if (rec.hours > 0) hoursLabel = formatHoursMinutes_(rec.hours);

      const dailyStatus = dailyStatusFor_(rec);
      const statusText = dailyStatus.text;
      const statusColors = dailyStatus.colors;

      dash.getRange(r, 1, 1, 6).setValues([[
        dateLabel, rec.name, ci, co, hoursLabel, statusText
      ]]);

      const altBg = (i % 2 === 0) ? C.white : C.grayLight;
      dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
        .setBackground(altBg).setFontSize(11)
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);

      dash.getRange(r, 1).setFontColor(C.gray).setHorizontalAlignment('center');
      dash.getRange(r, 2).setFontWeight('bold').setFontColor(C.tealDark);
      dash.getRange(r, 3, 1, 3).setHorizontalAlignment('center');

      dash.getRange(r, 6).setBackground(statusColors.bg)
        .setFontColor(statusColors.fg).setFontWeight('bold')
        .setHorizontalAlignment('center');

      r++;
    });
  }

  dash.setRowHeight(r, 14);
  r++;

  // Rules note
  dash.setRowHeight(r, 30);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
    .setValue(
      '⚙️  Flexible clock-in/out · ' +
      'Weekly target: ' + (Number(s.weeklyHoursTarget) || DEFAULT_WEEKLY_HOURS_TARGET) + ' hours · ' +
      'Week runs Monday–Sunday · Holidays and approved leave are shown separately'
    )
    .setFontColor(C.gray).setFontSize(10).setFontStyle('italic')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  try { dash.setFrozenRows(3); } catch (e) {}

  SpreadsheetApp.flush();
}

// ════════════════════════════════════════════════════════════════════
//  PER-EMPLOYEE TABS
// ════════════════════════════════════════════════════════════════════
function refreshEmployeeTabs_(ctx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  const attSheet = ss.getSheetByName(SHEET_ATTENDANCE);
  const sig = attSheet.getLastRow() + '|' + ctx.activeEmployees.length;
  const lastSig = props.getProperty('employee_tabs_sig');

  if (lastSig === sig) {
    console.log('Employee tabs unchanged — skipped.');
    return;
  }

  ctx.activeEmployees.forEach(function (name) {
    const tabName = sanitizeTabName_(name);
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) sheet = ss.insertSheet(tabName);
    try {
      paintEmployeeTab_(sheet, name, ctx);
    } catch (err) {
      console.error('Failed to paint tab for ' + name + ':', err);
    }
  });

  ss.getSheets().forEach(function (sheet) {
    const n = sheet.getName();
    if (n === SHEET_DASHBOARD || n === SHEET_ATTENDANCE ||
        n === SHEET_EMPLOYEES || n === SHEET_LEAVE ||
        n === SHEET_SETTINGS || n === SHEET_HOLIDAYS) return;

    const isActive = ctx.activeEmployees.some(function (e) {
      return sanitizeTabName_(e) === n;
    });
    if (!isActive) {
      const a1 = sheet.getRange('A1').getValue();
      if (typeof a1 === 'string' &&
          a1.indexOf('Personal Attendance') !== -1 &&
          a1.indexOf('INACTIVE') === -1) {
        sheet.getRange('A1').setValue('⚠️  ' + a1 + '  —  INACTIVE');
        sheet.setTabColor(C.gray);
      }
    }
  });

  props.setProperty('employee_tabs_sig', sig);
}

function paintEmployeeTab_(sheet, empName, ctx) {
  sheet.getCharts().forEach(function (ch) { sheet.removeChart(ch); });

  const employeeHistoryCount = ctx.weeklyHistory.filter(function (w) { return w.name === empName; }).length;
  const employeeDailyCount = ctx.empDaily.filter(function (r) { return r.name === empName; }).length;
  const requiredRows = 80 + employeeHistoryCount + employeeDailyCount;
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  const maxR0 = Math.max(sheet.getMaxRows(), 200);
  const maxC0 = Math.max(sheet.getMaxColumns(), 12);
  try { sheet.getRange(1, 1, maxR0, maxC0).breakApart(); } catch (e) {}
  sheet.getRange(1, 1, maxR0, maxC0).clear();

  const emp = ctx.byEmp[empName] || {
    daysPresent: 0, daysLeave: 0, sessionCount: 0, openSessions: 0,
    totalHours: 0, lastTs: null, lastAction: ''
  };

  const weeklyTarget = Number(ctx.settings.weeklyHoursTarget) || DEFAULT_WEEKLY_HOURS_TARGET;
  const weeklyRemaining = Math.max(weeklyTarget - (Number(emp.totalHours) || 0), 0);
  const personalWeeks = ctx.weeklyHistory.filter(function (w) { return w.name === empName; });
  const personalDaily = ctx.empDaily.filter(function (r) { return r.name === empName; });

  const TOTAL_COLS = 6;
  [190, 145, 145, 145, 155, 155].forEach(function (w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  let r = 1;

  sheet.setRowHeight(r, 72);
  sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
    .setValue('👤  ' + empName + ' — Personal Attendance & Weekly Monitoring')
    .setBackground(C.tealDark).setFontColor(C.white)
    .setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  sheet.setRowHeight(r, 30);
  sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
    .setValue('Last refreshed: ' +
      Utilities.formatDate(ctx.now, ctx.tz, 'EEEE, d MMM yyyy · h:mm a'))
    .setFontColor(C.gray).setFontSize(11).setFontStyle('italic')
    .setHorizontalAlignment('right').setVerticalAlignment('middle')
    .setBackground(C.grayLight);
  r++;

  sheet.setRowHeight(r, 14);
  r++;

  sheet.setRowHeight(r, 40);
  sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
    .setValue('📊  THIS WEEK — ' + weekLabel_(ctx.currentWeekStart, ctx.currentWeekEnd, ctx.tz))
    .setBackground(C.teal).setFontColor(C.white)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  sheet.setRowHeight(r, 34);
  sheet.getRange(r, 1, 1, TOTAL_COLS).setValues([[
    'DAYS WORKED', 'DAYS LEAVE', 'SESSIONS', 'OPEN SESSIONS', 'TOTAL HOURS', 'REMAINING'
  ]]);
  sheet.getRange(r, 1, 1, TOTAL_COLS)
    .setBackground(C.grayLight).setFontColor(C.tealDark)
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
  r++;

  sheet.setRowHeight(r, 52);
  sheet.getRange(r, 1, 1, TOTAL_COLS).setValues([[
    emp.daysPresent, emp.daysLeave, emp.sessionCount || 0, emp.openSessions || 0,
    formatHoursMinutes_(Number(emp.totalHours) || 0), formatHoursMinutes_(weeklyRemaining)
  ]]);
  sheet.getRange(r, 1, 1, TOTAL_COLS)
    .setFontSize(18).setFontWeight('bold').setFontColor(C.ink)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(r, 1, 1, TOTAL_COLS).setBackgrounds([[
    C.greenLight, C.greenLight, C.greenLight, C.greenLight,
    Number(emp.totalHours) >= weeklyTarget ? C.greenLight : C.amberLight,
    Number(emp.totalHours) >= weeklyTarget ? C.greenLight : C.amberLight
  ]]);

  if (emp.daysLeave > 0) sheet.getRange(r, 4).setBackground(C.tealLight).setFontColor(C.teal);
  r++;

  sheet.setRowHeight(r, 30);
  sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
    .setValue(
      'Weekly target: ' + weeklyTarget + ' h' +
      '   ·   ' + (Number(emp.totalHours) >= weeklyTarget
        ? 'Target met' : formatHoursMinutes_(weeklyRemaining) + ' remaining') +
      '   ·   Sessions this week: ' + (emp.sessionCount || 0) +
      '   ·   Last action: ' + (emp.lastAction || '—') +
      (emp.lastTs
        ? '  (' + Utilities.formatDate(emp.lastTs, ctx.tz, 'd MMM · h:mm a') + ')'
        : '')
    )
    .setFontColor(C.gray).setFontSize(11).setFontStyle('italic')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(C.grayLight);
  r++;

  sheet.setRowHeight(r, 14);
  r++;

  // Weekly history for this employee
  sheet.setRowHeight(r, 40);
  sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
    .setValue('📈  WEEKLY HOURS HISTORY')
    .setBackground(C.blue).setFontColor(C.white)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  sheet.setRowHeight(r, 34);
  sheet.getRange(r, 1, 1, TOTAL_COLS).setValues([[
    'WEEK', 'TOTAL HOURS', 'TARGET', 'REMAINING', 'SESSIONS', 'STATUS'
  ]]);
  sheet.getRange(r, 1, 1, TOTAL_COLS)
    .setBackground(C.grayLight).setFontColor(C.tealDark)
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
  r++;

  personalWeeks.forEach(function (week, i) {
    const met = week.totalHours >= weeklyTarget;
    const remaining = Math.max(weeklyTarget - week.totalHours, 0);
    sheet.setRowHeight(r, 32);
    sheet.getRange(r, 1, 1, TOTAL_COLS).setValues([[
      weekLabel_(week.weekStart, week.weekEnd, ctx.tz),
      formatHoursMinutes_(week.totalHours), weeklyTarget + ' h',
      formatHoursMinutes_(remaining), week.sessionCount,
      met ? '✅ Target met' : '⏳ ' + formatHoursMinutes_(remaining) + ' remaining'
    ]]);
    sheet.getRange(r, 1, 1, TOTAL_COLS)
      .setBackground(i % 2 === 0 ? C.white : C.grayLight).setFontSize(11)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r, 1).setFontColor(C.blue).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(r, 2, 1, 4).setHorizontalAlignment('center');
    sheet.getRange(r, 6).setBackground(met ? C.greenLight : C.amberLight)
      .setFontColor(met ? C.green : C.amber).setFontWeight('bold').setHorizontalAlignment('center');
    r++;
  });

  sheet.setRowHeight(r, 14);
  r++;

  sheet.setRowHeight(r, 40);
  sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
    .setValue('📅  DAILY LOG — LAST ' + EMP_LOG_DAYS_BACK + ' DAYS')
    .setBackground(C.teal).setFontColor(C.white)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  sheet.setRowHeight(r, 34);
  sheet.getRange(r, 1, 1, TOTAL_COLS).setValues([[
    'DATE', 'CLOCK IN', 'CLOCK OUT', 'HOURS', 'STATUS', ''
  ]]);
  sheet.getRange(r, 1, 1, TOTAL_COLS)
    .setBackground(C.grayLight)
    .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(r, 1, 1, 5)
    .setFontColor(C.tealDark).setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  r++;

  if (personalDaily.length === 0) {
    sheet.setRowHeight(r, 34);
    sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
      .setValue('No punches or leave in the last ' + EMP_LOG_DAYS_BACK + ' days.')
      .setFontColor(C.gray).setFontStyle('italic')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground(C.grayLight)
      .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
    r++;
  } else {
    personalDaily.forEach(function (rec, i) {
      sheet.setRowHeight(r, 32);

      const parts = rec.date.split('-');
      const d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
      const dateLabel = Utilities.formatDate(d, ctx.tz, 'EEE, d MMM');
      const ci = formatSessionTimes_(rec.sessions, ctx.tz, 'clockIn');
      const co = formatSessionTimes_(rec.sessions, ctx.tz, 'clockOut');
      let hoursLabel = '—';
      if (rec.hours > 0) hoursLabel = formatHoursMinutes_(rec.hours);

      const dailyStatus = dailyStatusFor_(rec);
      const statusText = dailyStatus.text;
      const statusColors = dailyStatus.colors;

      sheet.getRange(r, 1, 1, 6).setValues([[
        dateLabel, ci, co, hoursLabel, statusText, ''
      ]]);

      const altBg = (i % 2 === 0) ? C.white : C.grayLight;
      sheet.getRange(r, 1, 1, TOTAL_COLS)
        .setBackground(altBg).setFontSize(11)
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);

      sheet.getRange(r, 1).setFontColor(C.gray).setHorizontalAlignment('center');
      sheet.getRange(r, 2, 1, 3).setHorizontalAlignment('center');
      sheet.getRange(r, 5).setBackground(statusColors.bg)
        .setFontColor(statusColors.fg).setFontWeight('bold').setHorizontalAlignment('center');

      r++;
    });
  }

  sheet.setRowHeight(r, 14);
  r++;

  sheet.setRowHeight(r, 30);
  sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
    .setValue(
      '⚙️  Flexible clock-in/out · Weekly target: ' +
      (Number(ctx.settings.weeklyHoursTarget) || DEFAULT_WEEKLY_HOURS_TARGET) +
      ' hours · Week runs Monday–Sunday'
    )
    .setFontColor(C.gray).setFontSize(10).setFontStyle('italic')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setTabColor(Number(emp.totalHours) >= weeklyTarget ? C.green : C.teal);
}

function sanitizeTabName_(name) {
  return String(name).replace(/[\[\]\*\?\/\\:]/g, '').trim().substring(0, 100);
}

// ════════════════════════════════════════════════════════════════════
//  AUTO-REFRESH — QUEUE / TICK PATTERN
// ════════════════════════════════════════════════════════════════════
function checkPendingRefresh() {
  const props = PropertiesService.getScriptProperties();
  const dirtyAt = props.getProperty('dashboard_dirty_at');

  if (!dirtyAt) return;

  const ts = Number(dirtyAt);
  const age = Date.now() - ts;

  if (age < 3000) return;

  if (age > 15 * 60 * 1000) {
    props.deleteProperty('dashboard_dirty_at');
    console.log('Discarded stale dirty flag');
    return;
  }

  props.deleteProperty('dashboard_dirty_at');

  try {
    if (typeof refreshDashboard === 'function') {
      refreshDashboard();
      console.log('✓ Auto-refreshed after ' + Math.round(age / 1000) + 's');
    }
  } catch (e) {
    console.error('Auto-refresh failed:', e);
  }
}

function setupDashboardAutoRefresh() {
  const props = PropertiesService.getScriptProperties();

  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'checkPendingRefresh' || fn === 'runScheduledDashboardRefresh') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  props.deleteProperty('dashboard_dirty_at');
  props.deleteProperty('dashboard_refresh_pending');
  props.deleteProperty('dashboard_refresh_pending_at');

  let ok = false;
  try {
    ScriptApp.newTrigger('checkPendingRefresh')
      .timeBased().everyMinutes(1).create();
    ok = true;
  } catch (e) {
    console.error('Trigger creation failed:', e);
  }

  if (!ok) {
    SpreadsheetApp.getUi().alert(
      '❌ Could not create the recurring trigger.\n\n' +
      'Please add it manually:\n' +
      'Apps Script → Triggers → Add Trigger\n' +
      '  Function: checkPendingRefresh\n' +
      '  Event source: Time-driven\n' +
      '  Type: Minutes timer · Every minute'
    );
    return;
  }

  SpreadsheetApp.getUi().alert(
    '✅ Auto-refresh enabled.\n\n' +
    'A recurring trigger runs every minute.\n' +
    'After any punch, the dashboard updates within 60 seconds.\n\n' +
    'Cleaned up ' + removed + ' old trigger' + (removed === 1 ? '' : 's') + '.'
  );
}

function diagnoseAutoRefresh() {
  const props = PropertiesService.getScriptProperties();
  const triggers = ScriptApp.getProjectTriggers();

  const dirtyAt = props.getProperty('dashboard_dirty_at');
  const dirtyStr = dirtyAt
    ? Utilities.formatDate(new Date(Number(dirtyAt)),
        Session.getScriptTimeZone(), 'd MMM yyyy · HH:mm:ss')
    : '(none)';

  let triggerInfo = '';
  triggers.forEach(function (t) {
    triggerInfo += '• ' + t.getHandlerFunction() +
                   '  (' + t.getTriggerSource() + ')\n';
  });
  if (!triggerInfo) triggerInfo = '• (none)\n';

  const hasRecurring = triggers.some(function (t) {
    return t.getHandlerFunction() === 'checkPendingRefresh';
  });

  const status = hasRecurring
    ? '✅ ACTIVE — Check runs every 1 minute'
    : '❌ INACTIVE — Run "⏱️ Enable auto-refresh after punch"';

  SpreadsheetApp.getUi().alert(
    '🔍 Auto-refresh diagnosis\n\n' +
    'Status: ' + status + '\n\n' +
    'Dirty flag set at: ' + dirtyStr + '\n' +
    'Total triggers: ' + triggers.length + ' / 20\n\n' +
    'All triggers:\n' + triggerInfo
  );
}

function enableAutoRefresh() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshDashboard') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('refreshDashboard')
    .timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('✅ Hourly auto-refresh enabled.');
}

function disableAutoRefresh() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'refreshDashboard' ||
        fn === 'checkPendingRefresh' ||
        fn === 'runScheduledDashboardRefresh') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('dashboard_dirty_at');
  props.deleteProperty('dashboard_refresh_pending');
  props.deleteProperty('dashboard_refresh_pending_at');

  SpreadsheetApp.getUi().alert(
    removed > 0
      ? '✅ Auto-refresh disabled (' + removed + ' trigger removed).'
      : 'No auto-refresh triggers were active.'
  );
}
