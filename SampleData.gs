/**
 * SampleData.gs
 * Populates the Config Log with a realistic device-change history so you can
 * see the "Current State" sheet derive the latest config per device.
 *
 * Run "Add Sample Data" from the Config Tools menu. It appends rows to the log;
 * the onEdit trigger (or a manual "Rebuild Current State") refreshes the view.
 */

function addSampleData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(LOG_SHEET);
  if (!log) { setupTemplate(); log = ss.getSheetByName(LOG_SHEET); }

  // [Timestamp, Device ID, Type, Location, Firmware, IP, Status, ChangedBy, Notes]
  // Two devices appear more than once - Current State must keep only the latest.
  var rows = [
    [new Date('2026-05-01T09:00:00'), 'RTR-001', 'Router', 'HQ Rack A', '2.3.0', '10.0.0.1',  'Active',      'a.lopez', 'Initial provisioning'],
    [new Date('2026-05-10T14:30:00'), 'RTR-001', 'Router', 'HQ Rack A', '2.4.1', '10.0.0.1',  'Active',      'a.lopez', 'Firmware upgrade 2.3.0 -> 2.4.1'],
    [new Date('2026-05-02T11:15:00'), 'SW-014',  'Switch', 'HQ Rack B', '1.9.2', '10.0.0.14', 'Active',      'm.chen',  'Deployed'],
    [new Date('2026-06-01T08:05:00'), 'FW-002',  'Firewall', 'DMZ',     '5.1.0', '10.0.1.2',  'Active',      'm.chen',  'New ruleset applied'],
    [new Date('2026-06-15T16:45:00'), 'FW-002',  'Firewall', 'DMZ',     '5.1.0', '10.0.1.2',  'Maintenance', 'a.lopez', 'Scheduled patch window'],
    [new Date('2026-05-20T10:00:00'), 'AP-031',  'Access Point', 'Floor 3', '3.0.7', '10.0.2.31', 'Active',  'k.singh', 'Mounted, SSID configured'],
    [new Date('2026-06-18T12:00:00'), 'SRV-009', 'Server', 'HQ Rack C', '20.04', '10.0.0.9',  'Standby',     'k.singh', 'Failover node']
  ];

  var startRow = log.getLastRow() + 1;
  log.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  rebuildCurrentState();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Added ' + rows.length + ' sample rows. Current State refreshed.',
    'Config Tools', 5);
}
