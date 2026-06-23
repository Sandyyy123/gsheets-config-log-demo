# Device Configuration Log - Google Sheets + Apps Script

A two-sheet Google Sheets template for tracking device configurations over time,
with the **most current configuration per device** derived automatically.

This is a working demo of the pattern: an append-only log is the single source
of truth, and a clean "current state" view is always computed from it, so the
two can never drift out of sync.

## What it does

| Sheet | Purpose |
|-------|---------|
| **Config Log** | Append-only history. One row per configuration change. Structured columns (Device Type, Status) are constrained by dropdown validation so the log stays clean. |
| **Current State** | Auto-generated. Shows the latest row per `Device ID` (newest `Timestamp` wins). Rebuilt automatically on every edit, and on demand from the menu. |

### Columns
`Timestamp, Device ID, Device Type, Location, Firmware, IP Address, Status, Changed By, Notes`

- **Device Type** dropdown: Router, Switch, Firewall, Access Point, Server, Sensor
- **Status** dropdown: Active, Standby, Maintenance, Retired

## How "most current config" works

`rebuildCurrentState()` reads the whole log, groups rows by `Device ID`, and
keeps the row with the latest `Timestamp` for each device. The result is sorted
by `Device ID` and written to the **Current State** sheet. Because it is always
recomputed from the log, deleting or correcting a log row updates the current
view automatically.

It stays live through:
- a simple `onEdit` trigger (refreshes after manual edits), and
- an optional installable `onChange` trigger (`installTriggers()`) for
  programmatic / multi-user edits.

## Setup

1. Create a new Google Sheet.
2. **Extensions -> Apps Script**.
3. Add the three files in this repo:
   - `Code.gs` (core logic, menu, triggers)
   - `SampleData.gs` (optional demo data)
   - `appsscript.json` (manifest - paste into the project's `appsscript.json`)
4. Save, then reload the spreadsheet.
5. From the **Config Tools** menu choose **Set Up / Repair Template**, then
   **Add Sample Data**.
6. (Optional) run `installTriggers()` once from the editor to enable the
   onChange trigger.

## Files

- `Code.gs` - configuration, menu, validation, `rebuildCurrentState()`, triggers.
- `SampleData.gs` - realistic sample history (two devices change twice, so you
  can verify Current State keeps only the latest).
- `appsscript.json` - V8 runtime manifest.

## Notes

- The log is the source of truth; never edit Current State by hand (it is
  overwritten on each rebuild).
- Validation is applied to the first 1000 data rows; re-run **Set Up / Repair
  Template** to extend or repair it.

---
Demo by Dr. Sandeep Grover. Intended as an illustrative reference implementation.
