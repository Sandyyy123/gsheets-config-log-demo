/**
 * Device Configuration Log - Google Apps Script
 * ------------------------------------------------
 * Maintains two sheets in one spreadsheet:
 *
 *   "Config Log"   - an append-only history of every device configuration
 *                    change (one row per change), with data validation on the
 *                    structured columns so logging stays clean.
 *
 *   "Current State" - an auto-derived sheet showing the MOST CURRENT
 *                     configuration for each device (the latest row per
 *                     Device ID by timestamp). Rebuilt automatically whenever
 *                     the log changes, and on demand from the menu.
 *
 * The log is the single source of truth. "Current State" is always computed
 * from it, so it can never drift out of sync.
 */

// ----- Configuration ---------------------------------------------------------

var LOG_SHEET = 'Config Log';
var CURRENT_SHEET = 'Current State';

// Column order for the log (also used to build headers + validation).
var LOG_HEADERS = [
  'Timestamp',      // when the change was recorded
  'Device ID',      // unique device key, e.g. RTR-001
  'Device Type',    // dropdown
  'Location',       // free text
  'Firmware',       // e.g. 2.4.1
  'IP Address',     // e.g. 10.0.0.12
  'Status',         // dropdown
  'Changed By',     // person / system
  'Notes'           // free text
];

// Dropdown vocabularies for validated columns.
var DEVICE_TYPES = ['Router', 'Switch', 'Firewall', 'Access Point', 'Server', 'Sensor'];
var STATUSES     = ['Active', 'Standby', 'Maintenance', 'Retired'];

// ----- Menu ------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Config Tools')
    .addItem('Rebuild Current State', 'rebuildCurrentState')
    .addItem('Add Sample Data', 'addSampleData')
    .addSeparator()
    .addItem('Set Up / Repair Template', 'setupTemplate')
    .addToUi();
}

// ----- One-time / repair setup ----------------------------------------------

/**
 * Creates both sheets, writes headers, applies formatting and data validation.
 * Safe to re-run: it repairs an existing template without deleting log rows.
 */
function setupTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);

  // Header row
  log.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
     .setFontWeight('bold').setBackground('#6c5ce7').setFontColor('#ffffff');
  log.setFrozenRows(1);
  log.autoResizeColumns(1, LOG_HEADERS.length);

  applyValidation_(log);

  // Current State sheet
  var cur = ss.getSheetByName(CURRENT_SHEET) || ss.insertSheet(CURRENT_SHEET);
  cur.clear();
  cur.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
     .setFontWeight('bold').setBackground('#2d3436').setFontColor('#ffffff');
  cur.setFrozenRows(1);

  rebuildCurrentState();
  SpreadsheetApp.getActiveSpreadsheet().toast('Template ready.', 'Config Tools', 5);
}

/**
 * Applies dropdown validation to the Device Type and Status columns for a
 * generous number of rows, so new entries are constrained as you type.
 */
function applyValidation_(log) {
  var typeCol = LOG_HEADERS.indexOf('Device Type') + 1;
  var statusCol = LOG_HEADERS.indexOf('Status') + 1;
  var nRows = 1000;

  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEVICE_TYPES, true)
    .setAllowInvalid(false)
    .setHelpText('Pick a device type from the list.')
    .build();
  log.getRange(2, typeCol, nRows, 1).setDataValidation(typeRule);

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUSES, true)
    .setAllowInvalid(false)
    .setHelpText('Pick a status from the list.')
    .build();
  log.getRange(2, statusCol, nRows, 1).setDataValidation(statusRule);
}

// ----- Auto current-state derivation ----------------------------------------

/**
 * Rebuilds the "Current State" sheet from the log: for each Device ID, keeps
 * the row with the latest Timestamp. This is the core "most current config"
 * requirement.
 */
function rebuildCurrentState() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(LOG_SHEET);
  var cur = ss.getSheetByName(CURRENT_SHEET);
  if (!log || !cur) { return; }

  var values = log.getDataRange().getValues();
  if (values.length < 2) {
    clearBody_(cur);
    return;
  }

  var header = values[0];
  var tsIdx = header.indexOf('Timestamp');
  var idIdx = header.indexOf('Device ID');

  // latest[deviceId] = row with the newest timestamp
  var latest = {};
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var id = row[idIdx];
    if (id === '' || id === null) { continue; }
    var ts = toTime_(row[tsIdx]);
    if (!(id in latest) || ts >= latest[id]._ts) {
      row._ts = ts;
      latest[id] = row;
    }
  }

  // Sort by Device ID for a stable, readable view.
  var ids = Object.keys(latest).sort();
  var out = ids.map(function (id) {
    return latest[id].slice(0, header.length);
  });

  clearBody_(cur);
  if (out.length) {
    cur.getRange(2, 1, out.length, header.length).setValues(out);
    cur.autoResizeColumns(1, header.length);
  }
  cur.getRange(1, 1).setNote('Auto-generated from "' + LOG_SHEET +
    '". Last rebuilt: ' + new Date());
}

function clearBody_(sheet) {
  var last = sheet.getLastRow();
  if (last > 1) {
    sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).clearContent().clearNote();
  }
}

function toTime_(v) {
  if (v instanceof Date) { return v.getTime(); }
  var d = new Date(v);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

// ----- Trigger: keep Current State live -------------------------------------

/**
 * Simple installable-style edit handler. When a cell in the log changes, the
 * Current State sheet is rebuilt so it is always in sync.
 *
 * Note: onEdit is a simple trigger. For multi-user / programmatic edits, also
 * install rebuildCurrentState on an installable onChange trigger via
 * installTriggers().
 */
function onEdit(e) {
  if (!e || !e.range) { return; }
  if (e.range.getSheet().getName() !== LOG_SHEET) { return; }
  if (e.range.getRow() === 1) { return; } // header edit, ignore
  rebuildCurrentState();
}

/** Installs an onChange trigger so structural changes also refresh state. */
function installTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'rebuildCurrentState') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('rebuildCurrentState')
    .forSpreadsheet(ss)
    .onChange()
    .create();
}
