/**
 * ════════════════════════════════════════════════════════════════════
 *  Goolee.my — Attendance Dashboard
 * ════════════════════════════════════════════════════════════════════
 *  • Every section aligned to columns A:L
 *  • breakApart() at start so leftover merges never linger
 *  • Today's Performance + monthly table + leave + holidays + daily log
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
const DEFAULT_WORK_START_HOUR  = 10;
const DEFAULT_WORK_START_MIN   = 0;
const DEFAULT_WORK_END_HOUR    = 19;
const DEFAULT_WORK_END_MIN     = 0;
const DEFAULT_LATE_AFTER_MIN   = 15;
const DEFAULT_EARLY_LEAVE_MIN  = 15;

const LOG_DAYS_BACK     = 14;
const EMP_LOG_DAYS_BACK = 30;

// ── Cache keys ──
const CACHE_SETTINGS_KEY = 'goolee_settings_v2';
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
    .addItem('🎨  Reapply red highlight',        'highlightLateInAttendance')
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
  const sheet = setupSettingsSheet();
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
    const existing = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues()
      .map(r => String(r[0] || '').trim().toLowerCase());
    const additions = [
      ['Saturday Working Day', 'TRUE', 'Set FALSE to disable Saturday attendance.'],
      ['Saturday Work Start Time', '09:00', '24-hour format. Saturday schedule.'],
      ['Saturday Work End Time', '12:00', '24-hour format. Saturday schedule.']
    ].filter(r => existing.indexOf(r[0].toLowerCase()) === -1);
    if (additions.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, additions.length, 3).setValues(additions);
    }
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
    ['Work Start Time',          '10:00',              '24-hour format. Example: 09:00 or 10:00.'],
    ['Work End Time',            '19:00',              '24-hour format. Example: 18:00 or 19:00.'],
    ['Saturday Working Day',     'TRUE',               'Set FALSE to disable Saturday attendance.'],
    ['Saturday Work Start Time', '09:00',              '24-hour format. Saturday schedule.'],
    ['Saturday Work End Time',   '12:00',              '24-hour format. Saturday schedule.'],
    ['Late After (min)',         '15',                 'Minutes after Work Start before a clock-in is late.'],
    ['Early Leave Grace (min)',  '15',                 'Minutes before Work End that still count as on-time exit.'],
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
  cache.remove('goolee_settings_api_v1');
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

  // Settings are edited directly by HR. Always use the current sheet values
  // when rebuilding the dashboard instead of an older script-cache entry.
  const settings = readSettings_(true);
  const lateHour = settings.workStartHour;
  const lateMin  = settings.workStartMin + settings.lateAfterMin;
  const endHour  = settings.workEndHour;
  const endMin   = Math.max(settings.workEndMin - settings.earlyLeaveMin, 0);
  const satLateTotal = settings.saturdayStartHour * 60 +
    settings.saturdayStartMin + settings.lateAfterMin;
  const satLateHour = Math.floor(satLateTotal / 60);
  const satLateMin = satLateTotal % 60;
  const satEarlyTotal = settings.saturdayEndHour * 60 +
    settings.saturdayEndMin - settings.earlyLeaveMin;
  const satEarlyHour = Math.floor(Math.max(satEarlyTotal, 0) / 60);
  const satEarlyMin = Math.max(satEarlyTotal, 0) % 60;

  const lateRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      '=AND($C2="Clock In", IF(WEEKDAY($A2,2)=6,' +
      'TIMEVALUE($G2)>TIME(' + satLateHour + ',' + satLateMin + ',0),' +
      'TIMEVALUE($G2)>TIME(' + lateHour + ',' + lateMin + ',0)))')
    .setBackground(C.redLight).setFontColor(C.red).setRanges([range]).build();

  const earlyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      '=AND($C2="Clock Out", IF(WEEKDAY($A2,2)=6,' +
      'TIMEVALUE($G2)<TIME(' + satEarlyHour + ',' + satEarlyMin + ',0),' +
      'TIMEVALUE($G2)<TIME(' + endHour + ',' + endMin + ',0)))')
    .setBackground(C.redLight).setFontColor(C.red).setRanges([range]).build();

  const overtimeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(
      '=AND($C2="Clock Out", IF(WEEKDAY($A2,2)=6,' +
      'TIMEVALUE($G2)>TIME(' + settings.saturdayEndHour + ',' + settings.saturdayEndMin + ',0),' +
      'TIMEVALUE($G2)>TIME(' + settings.workEndHour + ',' + settings.workEndMin + ',0)))')
    .setBackground(C.greenLight).setFontColor(C.green).setRanges([range]).build();

  const accuracyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($E2<>"", $E2 > 100)')
    .setBackground(C.amberLight).setFontColor(C.amber).setRanges([range]).build();

  sheet.setConditionalFormatRules([accuracyRule, overtimeRule, earlyRule, lateRule]);
}

// ════════════════════════════════════════════════════════════════════
//  SETTINGS READER
// ════════════════════════════════════════════════════════════════════
function readSettings_(forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    const cached = cache.get(CACHE_SETTINGS_KEY);
    if (cached) {
      try {
        const parsed = JSON.parse(cached);
        parsed.workingDays = new Set(parsed.workingDays);
        return parsed;
      } catch (e) {}
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SETTINGS);

  const out = {
    workingDays:    parseWorkingDays_(DEFAULT_WORKING_DAYS),
    workStartHour:  DEFAULT_WORK_START_HOUR,
    workStartMin:   DEFAULT_WORK_START_MIN,
    workEndHour:    DEFAULT_WORK_END_HOUR,
    workEndMin:     DEFAULT_WORK_END_MIN,
    saturdayWorkingDay: false,
    saturdayStartHour: 9,
    saturdayStartMin:  0,
    saturdayEndHour:   12,
    saturdayEndMin:    0,
    lateAfterMin:   DEFAULT_LATE_AFTER_MIN,
    earlyLeaveMin:  DEFAULT_EARLY_LEAVE_MIN,
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
      } else if (key === 'company name') {
        out.companyName = val;
      }
    }
  }

  const serializable = Object.assign({}, out, {
    workingDays: Array.from(out.workingDays)
  });
  if (out.saturdayWorkingDay) out.workingDays.add('Sat');
  else out.workingDays.delete('Sat');
  serializable.workingDays = Array.from(out.workingDays);
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
  const settings = readSettings_(true);
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

  // Group by (name, date)
  const dailyMap = {};
  for (const p of punches) {
    const dateStr = Utilities.formatDate(p.ts, tz, 'yyyy-MM-dd');
    const key = p.name + '|' + dateStr;
    if (!dailyMap[key]) {
      dailyMap[key] = {
        name: p.name, date: dateStr,
        clockIn: null, clockOut: null,
        leave: null, leaveNote: '', isHoliday: false
      };
    }
    const rec = dailyMap[key];
    if (p.action === 'Clock In') {
      if (!rec.clockIn || p.ts < rec.clockIn) rec.clockIn = p.ts;
    } else if (p.action === 'Clock Out') {
      if (!rec.clockOut || p.ts > rec.clockOut) rec.clockOut = p.ts;
    }
  }

  for (const lv of leaves) {
    const key = lv.name + '|' + lv.date;
    if (!dailyMap[key]) {
      dailyMap[key] = {
        name: lv.name, date: lv.date,
        clockIn: null, clockOut: null,
        leave: null, leaveNote: '', isHoliday: false
      };
    }
    dailyMap[key].leave = lv.type;
    dailyMap[key].leaveNote = lv.notes || '';
  }

  // Per-record computation
  const dailyList = [];
  for (const k in dailyMap) {
    const rec = dailyMap[k];
    const schedule = scheduleForTimestamp_(rec.clockIn || rec.clockOut, tz, settings);
    let hours = 0;
    if (rec.clockIn && rec.clockOut && rec.clockOut > rec.clockIn) {
      hours = (rec.clockOut - rec.clockIn) / 3600000;
    }
    const isHol = holidaySet.has(rec.date);
    const late = (rec.clockIn && !rec.leave && !isHol)
      ? isLate_(rec.clockIn, tz, schedule) : false;

    let leftEarly = false;
    let earlyMin = 0;
    let overtimeMin = 0;
    if (rec.clockOut && !rec.leave && !isHol) {
      const outH = Number(Utilities.formatDate(rec.clockOut, tz, 'H'));
      const outM = Number(Utilities.formatDate(rec.clockOut, tz, 'm'));
      const outMin = outH * 60 + outM;
      const endMin = schedule.workEndHour * 60 + schedule.workEndMin;
      const earlyThreshold = endMin - schedule.earlyLeaveMin;

      if (outMin < earlyThreshold) {
        leftEarly = true;
        earlyMin = endMin - outMin;
      } else if (outMin > endMin) {
        overtimeMin = outMin - endMin;
      }
    }

    dailyList.push({
      name: rec.name, date: rec.date,
      clockIn: rec.clockIn, clockOut: rec.clockOut,
      hours: hours, late: late,
      leave: rec.leave, leaveNote: rec.leaveNote,
      isHoliday: isHol,
      leftEarly: leftEarly, earlyMin: earlyMin, overtimeMin: overtimeMin
    });
  }

  dailyList.sort(function (a, b) {
    if (a.date !== b.date) return b.date.localeCompare(a.date);
    return a.name.localeCompare(b.name);
  });

  // Today's KPIs
  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const todayRecords = dailyList.filter(r => r.date === todayStr);

  const todayIsHoliday = holidaySet.has(todayStr);
  const presentToday = todayRecords.filter(r => r.clockIn).length;
  const leaveToday   = todayRecords.filter(r => r.leave && !r.clockIn).length;
  const lateToday    = todayRecords.filter(r => r.late).length;
  const earlyToday   = todayRecords.filter(r => r.leftEarly).length;
  const onTimeToday  = presentToday - lateToday;
  const onTimePct    = presentToday > 0 ? (onTimeToday / presentToday) * 100 : 0;
  const hoursToday   = todayRecords.reduce((s, r) => s + r.hours, 0);

  const workingDaysElapsed = countWorkingDays_(
    now.getFullYear(), now.getMonth(),
    settings.workingDays, holidaySet, tz, now);
  const workingDaysTotal = countWorkingDays_(
    now.getFullYear(), now.getMonth(),
    settings.workingDays, holidaySet, tz, null);

  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const monthStartStr = Utilities.formatDate(monthStart, tz, 'yyyy-MM-dd');
  const monthEndStr = Utilities.formatDate(
    new Date(now.getFullYear(), now.getMonth() + 1, 0), tz, 'yyyy-MM-dd');
  const holidaysThisMonth = holidays.filter(
    h => h.date >= monthStartStr && h.date <= monthEndStr);

  const monthRecords = dailyList.filter(r => r.date >= monthStartStr);

  // Per-employee monthly
  const byEmp = {};
  activeEmployees.forEach(function (name) {
    byEmp[name] = {
      name: name,
      daysPresent: 0, daysLeave: 0, daysLate: 0, daysEarly: 0, daysAbsent: 0,
      totalHours: 0, totalOvertime: 0,
      lastTs: null, lastAction: ''
    };
  });

  monthRecords.forEach(function (r) {
    if (!byEmp[r.name]) {
      byEmp[r.name] = {
        name: r.name,
        daysPresent: 0, daysLeave: 0, daysLate: 0, daysEarly: 0, daysAbsent: 0,
        totalHours: 0, totalOvertime: 0, lastTs: null, lastAction: ''
      };
    }
    const e = byEmp[r.name];
    if (r.leave) e.daysLeave++;
    if (r.clockIn) e.daysPresent++;
    if (r.late) e.daysLate++;
    if (r.leftEarly) e.daysEarly++;
    if (r.overtimeMin > 0) e.totalOvertime += r.overtimeMin;
    e.totalHours += r.hours;
  });

  activeEmployees.forEach(function (name) {
    const e = byEmp[name];
    if (!e) return;
    e.daysAbsent = Math.max(workingDaysElapsed - e.daysPresent - e.daysLeave, 0);
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

  // Today per employee
  const todayByEmp = {};
  activeEmployees.forEach(function (name) {
    todayByEmp[name] = {
      name: name,
      clockIn: null, clockOut: null, hours: 0,
      late: false, leftEarly: false,
      earlyMin: 0, overtimeMin: 0,
      leave: null, leaveNote: ''
    };
  });
  todayRecords.forEach(function (r) {
    if (todayByEmp[r.name]) {
      todayByEmp[r.name].clockIn     = r.clockIn;
      todayByEmp[r.name].clockOut    = r.clockOut;
      todayByEmp[r.name].hours       = r.hours;
      todayByEmp[r.name].late        = r.late;
      todayByEmp[r.name].leftEarly   = r.leftEarly;
      todayByEmp[r.name].earlyMin    = r.earlyMin;
      todayByEmp[r.name].overtimeMin = r.overtimeMin;
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
    workingDaysElapsed: workingDaysElapsed,
    workingDaysTotal: workingDaysTotal,
    presentToday: presentToday, lateToday: lateToday,
    leaveToday: leaveToday, earlyToday: earlyToday,
    onTimePct: onTimePct, hoursToday: hoursToday,
    activeCount: activeEmployees.length,
    activeEmployees: activeEmployees,
    todayPerf: todayPerf,
    onLeaveToday: onLeaveToday,
    employeeRows: employeeRows,
    recentDaily: recentDaily,
    empDaily: empDaily,
    byEmp: byEmp
  };
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

function scheduleForTimestamp_(ts, tz, settings) {
  if (!ts) return settings;
  const day = Utilities.formatDate(ts, tz, 'EEE');
  if (day === 'Sat' && settings.saturdayWorkingDay) {
    return {
      workStartHour: settings.saturdayStartHour,
      workStartMin: settings.saturdayStartMin,
      workEndHour: settings.saturdayEndHour,
      workEndMin: settings.saturdayEndMin,
      lateAfterMin: settings.lateAfterMin,
      earlyLeaveMin: settings.earlyLeaveMin
    };
  }
  return settings;
}

function isLate_(ts, tz, settings) {
  const h = Number(Utilities.formatDate(ts, tz, 'H'));
  const m = Number(Utilities.formatDate(ts, tz, 'm'));
  const minutes = h * 60 + m;
  const threshold = settings.workStartHour * 60 + settings.workStartMin + settings.lateAfterMin;
  return minutes > threshold;
}

// ════════════════════════════════════════════════════════════════════
//  FORMATTERS
// ════════════════════════════════════════════════════════════════════
function pad2_(n) { return (n < 10 ? '0' : '') + n; }

function formatTimeLabel_(h, m) {
  const period = h >= 12 ? 'PM' : 'AM';
  const h12 = h % 12 || 12;
  return h12 + ':' + pad2_(m) + ' ' + period;
}

function formatDuration_(mins) {
  mins = Math.round(mins);
  if (mins <= 0) return '0m';
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  if (h === 0) return m + 'm';
  if (m === 0) return h + 'h';
  return h + 'h ' + m + 'm';
}

function formatHoursMinutes_(hours) {
  const total = Math.round(hours * 60);
  const h = Math.floor(total / 60);
  const m = total % 60;
  if (h === 0) return m + 'm';
  if (m === 0) return h + 'h';
  return h + 'h ' + m + 'm';
}

// ════════════════════════════════════════════════════════════════════
//  TODAY'S STATUS BUILDER
// ════════════════════════════════════════════════════════════════════
function todayStatusFor_(emp, ctx) {
  if (ctx.todayIsHoliday) {
    return { text: '🎉 Holiday', bg: HOLIDAY_COLOR.bg, fg: HOLIDAY_COLOR.fg };
  }
  if (emp.leave) {
    const col = leaveColor_(emp.leave);
    let text = col.icon + ' ' + emp.leave;
    if (emp.clockIn && !emp.clockOut) {
      text += ' · ⏱️ Working';
    } else if (emp.clockIn && emp.clockOut) {
      text += ' · 🏁 ' + formatHoursMinutes_(emp.hours) + ' worked';
    }
    return { text: text, bg: col.bg, fg: col.fg };
  }
  if (!emp.clockIn) {
    return { text: '⏳ Not yet in', bg: C.grayLight, fg: C.gray };
  }
  if (!emp.clockOut) {
    if (emp.late) return { text: '⏰ Late · Working', bg: C.redLight, fg: C.red };
    return { text: '✅ On Time · Working', bg: C.greenLight, fg: C.green };
  }
  if (emp.late && emp.leftEarly) {
    return {
      text: '⏰🚪 Late + Left Early · ' + formatDuration_(emp.earlyMin),
      bg: C.redLight, fg: C.red
    };
  }
  if (emp.late && emp.overtimeMin > 0) {
    return {
      text: '⏰➕ Late + Overtime · +' + formatDuration_(emp.overtimeMin),
      bg: C.amberLight, fg: C.amber
    };
  }
  if (emp.late) {
    return { text: '⏰ Late', bg: C.redLight, fg: C.red };
  }
  if (emp.leftEarly) {
    return {
      text: '🚪 Left Early · ' + formatDuration_(emp.earlyMin),
      bg: C.redLight, fg: C.red
    };
  }
  if (emp.overtimeMin > 0) {
    return {
      text: '➕ Overtime · +' + formatDuration_(emp.overtimeMin),
      bg: C.greenLight, fg: C.green
    };
  }
  return { text: '🏁 Completed · On Time', bg: C.greenLight, fg: C.green };
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

  resetDashboard_(dash);

  const s = ctx.settings;
  const startLabel = formatTimeLabel_(s.workStartHour, s.workStartMin);
  const endLabel   = formatTimeLabel_(s.workEndHour,   s.workEndMin);
  const lateLabel  = formatTimeLabel_(s.workStartHour, s.workStartMin + s.lateAfterMin);
  const earlyLabel = formatTimeLabel_(s.workEndHour,   Math.max(s.workEndMin - s.earlyLeaveMin, 0));

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
    { label: '✅  PRESENT', value: ctx.presentToday + ' / ' + ctx.activeCount,
      bg: C.greenLight, fg: C.green },
    { label: '⏰  LATE',
      value: ctx.lateToday + (ctx.presentToday > 0 ? '  (' + Math.round(100 - ctx.onTimePct) + '%)' : ''),
      bg: ctx.lateToday > 0 ? C.redLight : C.greenLight,
      fg: ctx.lateToday > 0 ? C.red : C.green },
    { label: '🚪  LEFT EARLY', value: String(ctx.earlyToday),
      bg: ctx.earlyToday > 0 ? C.redLight : C.greenLight,
      fg: ctx.earlyToday > 0 ? C.red : C.green },
    { label: '🌴  ON LEAVE', value: String(ctx.leaveToday),
      bg: ctx.leaveToday > 0 ? C.tealLight : C.grayLight,
      fg: ctx.leaveToday > 0 ? C.teal : C.gray },
    { label: '📆  WORKDAYS',
      value: ctx.workingDaysElapsed + ' / ' + ctx.workingDaysTotal,
      bg: C.blueLight, fg: C.blue }
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

      const ci = emp.clockIn ? Utilities.formatDate(emp.clockIn, ctx.tz, 'h:mm a') : '—';
      const co = emp.clockOut ? Utilities.formatDate(emp.clockOut, ctx.tz, 'h:mm a') : '—';
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

      if (emp.late)         dash.getRange(r, 2).setFontColor(C.red).setFontWeight('bold');
      if (emp.leftEarly)    dash.getRange(r, 3).setFontColor(C.red).setFontWeight('bold');
      if (emp.overtimeMin > 0) dash.getRange(r, 3).setFontColor(C.green).setFontWeight('bold');

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

  // EMPLOYEE PERFORMANCE — THIS MONTH
  dash.setRowHeight(r, 40);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
    .setValue('🏢  EMPLOYEE PERFORMANCE — THIS MONTH')
    .setBackground(C.teal).setFontColor(C.white)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  dash.setRowHeight(r, 34);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).setValues([[
    'EMPLOYEE', 'PRESENT', 'LEAVE', 'LATE', 'EARLY', 'ABSENT',
    'ON TIME %', 'TOTAL HOURS', 'OVERTIME', 'LAST ACTION', 'LAST PUNCH', ''
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

      const onTimePct = emp.daysPresent > 0
        ? ((emp.daysPresent - emp.daysLate) / emp.daysPresent) : 0;
      const totalHours = Number(emp.totalHours) || 0;
      const overtimeHours = (Number(emp.totalOvertime) || 0) / 60;
      const lastAction = emp.lastAction || '—';
      const lastPunch = emp.lastTs
        ? Utilities.formatDate(emp.lastTs, ctx.tz, 'd MMM · h:mm a')
        : '—';

      dash.getRange(r, 1, 1, DASH_TOTAL_COLS).setValues([[
        emp.name, emp.daysPresent, emp.daysLeave, emp.daysLate, emp.daysEarly,
        emp.daysAbsent, onTimePct, totalHours, overtimeHours,
        lastAction, lastPunch, ''
      ]]);

      const altBg = (i % 2 === 0) ? C.white : C.grayLight;
      dash.getRange(r, 1, 1, DASH_TOTAL_COLS)
        .setBackground(altBg).setFontSize(11)
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);

      dash.getRange(r, 7).setNumberFormat('0%');
      dash.getRange(r, 8).setNumberFormat('0.00');
      dash.getRange(r, 9).setNumberFormat('0.00');

      dash.getRange(r, 1).setHorizontalAlignment('left')
        .setFontWeight('bold').setFontColor(C.tealDark);
      dash.getRange(r, 2, 1, 8).setHorizontalAlignment('center');
      dash.getRange(r, 10, 1, 2).setFontColor(C.gray).setFontSize(10)
        .setHorizontalAlignment('center');

      if (emp.daysLeave > 0)  dash.getRange(r, 3).setFontColor(C.teal).setFontWeight('bold');
      if (emp.daysLate > 0)   dash.getRange(r, 4).setFontColor(C.red).setFontWeight('bold');
      if (emp.daysEarly > 0)  dash.getRange(r, 5).setFontColor(C.red).setFontWeight('bold');
      if (emp.daysAbsent > 0) dash.getRange(r, 6).setFontColor(C.red).setFontWeight('bold');
      else                    dash.getRange(r, 6).setFontColor(C.green).setFontWeight('bold');
      if (overtimeHours > 0)  dash.getRange(r, 9).setFontColor(C.green).setFontWeight('bold');

      const otCell = dash.getRange(r, 7);
      if (onTimePct >= 0.9) otCell.setFontColor(C.green).setFontWeight('bold');
      else if (onTimePct > 0 && onTimePct < 0.7) otCell.setFontColor(C.red).setFontWeight('bold');

      r++;
    });
  }

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
      const ci = rec.clockIn  ? Utilities.formatDate(rec.clockIn,  ctx.tz, 'h:mm a') : '—';
      const co = rec.clockOut ? Utilities.formatDate(rec.clockOut, ctx.tz, 'h:mm a') : '—';
      let hoursLabel = '—';
      if (rec.hours > 0) hoursLabel = formatHoursMinutes_(rec.hours);

      let statusText, statusColors;
      if (rec.isHoliday) {
        statusText = HOLIDAY_COLOR.icon + ' Holiday';
        statusColors = HOLIDAY_COLOR;
      } else if (rec.leave) {
        const col = leaveColor_(rec.leave);
        statusText = col.icon + ' ' + rec.leave;
        statusColors = col;
      } else if (rec.late && rec.leftEarly) {
        statusText = '⏰🚪 Late + Left Early · ' + formatDuration_(rec.earlyMin);
        statusColors = { bg: C.redLight, fg: C.red };
      } else if (rec.late && rec.overtimeMin > 0) {
        statusText = '⏰➕ Late + Overtime · +' + formatDuration_(rec.overtimeMin);
        statusColors = { bg: C.amberLight, fg: C.amber };
      } else if (rec.late) {
        statusText = '⏰ Late';
        statusColors = { bg: C.redLight, fg: C.red };
      } else if (rec.leftEarly) {
        statusText = '🚪 Left Early · ' + formatDuration_(rec.earlyMin);
        statusColors = { bg: C.redLight, fg: C.red };
      } else if (rec.overtimeMin > 0) {
        statusText = '➕ Overtime · +' + formatDuration_(rec.overtimeMin);
        statusColors = { bg: C.greenLight, fg: C.green };
      } else if (rec.clockIn) {
        statusText = '✅ On Time';
        statusColors = { bg: C.greenLight, fg: C.green };
      } else {
        statusText = '—';
        statusColors = { bg: C.grayLight, fg: C.gray };
      }

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

      if (rec.late)          dash.getRange(r, 3).setFontColor(C.red).setFontWeight('bold');
      if (rec.leftEarly)     dash.getRange(r, 4).setFontColor(C.red).setFontWeight('bold');
      if (rec.overtimeMin > 0) dash.getRange(r, 4).setFontColor(C.green).setFontWeight('bold');

      r++;
    });
  }

  dash.setRowHeight(r, 14);
  r++;

  // Rules note
  dash.setRowHeight(r, 30);
  dash.getRange(r, 1, 1, DASH_TOTAL_COLS).merge()
    .setValue(
      '⚙️  Working days: ' + Array.from(s.workingDays).join(', ') +
      '  ·  Work hours: ' + startLabel + ' – ' + endLabel +
      '  ·  Late after ' + lateLabel +
      '  ·  Early before ' + earlyLabel +
      '  ·  Holidays and approved leave are excluded'
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

  const maxR0 = Math.max(sheet.getMaxRows(), 200);
  const maxC0 = Math.max(sheet.getMaxColumns(), 12);
  try { sheet.getRange(1, 1, maxR0, maxC0).breakApart(); } catch (e) {}
  sheet.getRange(1, 1, maxR0, maxC0).clear();

  const emp = ctx.byEmp[empName] || {
    daysPresent: 0, daysLeave: 0, daysLate: 0, daysEarly: 0, daysAbsent: 0,
    totalHours: 0, totalOvertime: 0, lastTs: null, lastAction: ''
  };

  const onTimePct = emp.daysPresent > 0
    ? ((emp.daysPresent - emp.daysLate) / emp.daysPresent) : 0;
  const avgHours = emp.daysPresent > 0 ? emp.totalHours / emp.daysPresent : 0;
  const overtimeHours = (Number(emp.totalOvertime) || 0) / 60;
  const personalDaily = ctx.empDaily.filter(function (r) { return r.name === empName; });

  const TOTAL_COLS = 6;
  [190, 145, 145, 145, 155, 155].forEach(function (w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  let r = 1;

  sheet.setRowHeight(r, 72);
  sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
    .setValue('👤  ' + empName + ' — Personal Attendance')
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
    .setValue('📊  THIS MONTH')
    .setBackground(C.teal).setFontColor(C.white)
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  sheet.setRowHeight(r, 34);
  sheet.getRange(r, 1, 1, TOTAL_COLS).setValues([[
    'DAYS PRESENT', 'DAYS LATE', 'DAYS EARLY', 'DAYS LEAVE', 'TOTAL HOURS', 'OVERTIME'
  ]]);
  sheet.getRange(r, 1, 1, TOTAL_COLS)
    .setBackground(C.grayLight).setFontColor(C.tealDark)
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);
  r++;

  sheet.setRowHeight(r, 52);
  sheet.getRange(r, 1, 1, TOTAL_COLS).setValues([[
    emp.daysPresent, emp.daysLate, emp.daysEarly, emp.daysLeave,
    Number(emp.totalHours) || 0, overtimeHours
  ]]);
  sheet.getRange(r, 1, 1, TOTAL_COLS)
    .setFontSize(18).setFontWeight('bold').setFontColor(C.ink)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, C.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(r, 5, 1, 2).setNumberFormat('0.00');
  sheet.getRange(r, 1, 1, TOTAL_COLS).setBackgrounds([[
    C.greenLight, C.greenLight, C.greenLight, C.greenLight, C.amberLight, C.greenLight
  ]]);

  if (emp.daysLate > 0)  sheet.getRange(r, 2).setBackground(C.redLight).setFontColor(C.red);
  if (emp.daysEarly > 0) sheet.getRange(r, 3).setBackground(C.redLight).setFontColor(C.red);
  if (emp.daysLeave > 0) sheet.getRange(r, 4).setBackground(C.tealLight).setFontColor(C.teal);
  if (overtimeHours > 0) sheet.getRange(r, 6).setBackground(C.greenLight).setFontColor(C.green);
  r++;

  sheet.setRowHeight(r, 30);
  sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
    .setValue(
      'On-time rate: ' + Math.round(onTimePct * 100) + '%' +
      '   ·   Avg hours/day: ' + avgHours.toFixed(1) + ' h' +
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
      const ci = rec.clockIn  ? Utilities.formatDate(rec.clockIn,  ctx.tz, 'h:mm a') : '—';
      const co = rec.clockOut ? Utilities.formatDate(rec.clockOut, ctx.tz, 'h:mm a') : '—';
      let hoursLabel = '—';
      if (rec.hours > 0) hoursLabel = formatHoursMinutes_(rec.hours);

      let statusText, statusColors;
      if (rec.isHoliday) {
        statusText = HOLIDAY_COLOR.icon + ' Holiday';
        statusColors = HOLIDAY_COLOR;
      } else if (rec.leave) {
        const col = leaveColor_(rec.leave);
        statusText = col.icon + ' ' + rec.leave;
        statusColors = col;
      } else if (rec.late && rec.leftEarly) {
        statusText = '⏰🚪 Late + Left Early · ' + formatDuration_(rec.earlyMin);
        statusColors = { bg: C.redLight, fg: C.red };
      } else if (rec.late && rec.overtimeMin > 0) {
        statusText = '⏰➕ Late + Overtime · +' + formatDuration_(rec.overtimeMin);
        statusColors = { bg: C.amberLight, fg: C.amber };
      } else if (rec.late) {
        statusText = '⏰ Late';
        statusColors = { bg: C.redLight, fg: C.red };
      } else if (rec.leftEarly) {
        statusText = '🚪 Left Early · ' + formatDuration_(rec.earlyMin);
        statusColors = { bg: C.redLight, fg: C.red };
      } else if (rec.overtimeMin > 0) {
        statusText = '➕ Overtime · +' + formatDuration_(rec.overtimeMin);
        statusColors = { bg: C.greenLight, fg: C.green };
      } else if (rec.clockIn) {
        statusText = '✅ On Time';
        statusColors = { bg: C.greenLight, fg: C.green };
      } else {
        statusText = '—';
        statusColors = { bg: C.grayLight, fg: C.gray };
      }

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

      if (rec.late)          sheet.getRange(r, 2).setFontColor(C.red).setFontWeight('bold');
      if (rec.leftEarly)     sheet.getRange(r, 3).setFontColor(C.red).setFontWeight('bold');
      if (rec.overtimeMin > 0) sheet.getRange(r, 3).setFontColor(C.green).setFontWeight('bold');

      r++;
    });
  }

  sheet.setRowHeight(r, 14);
  r++;

  sheet.setRowHeight(r, 30);
  sheet.getRange(r, 1, 1, TOTAL_COLS).merge()
    .setValue(
      '⚙️  Work hours: ' +
      formatTimeLabel_(ctx.settings.workStartHour, ctx.settings.workStartMin) + ' – ' +
      formatTimeLabel_(ctx.settings.workEndHour, ctx.settings.workEndMin) +
      '   ·   Late after ' +
      formatTimeLabel_(ctx.settings.workStartHour,
        ctx.settings.workStartMin + ctx.settings.lateAfterMin) +
      '   ·   Early before ' +
      formatTimeLabel_(ctx.settings.workEndHour,
        Math.max(ctx.settings.workEndMin - ctx.settings.earlyLeaveMin, 0))
    )
    .setFontColor(C.gray).setFontSize(10).setFontStyle('italic')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setTabColor((emp.daysLate > 0 || emp.daysEarly > 0) ? C.red : C.teal);
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
