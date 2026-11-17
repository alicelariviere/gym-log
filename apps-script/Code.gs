/**
 * Gym log: the bridge between the phone app and your Google Sheets.
 * Deployed as a web app (Execute as: me, Who has access: anyone).
 * Only requests carrying KEY are answered.
 */
const KEY = 'PUT-YOUR-OWN-KEY-HERE'; // the live copy in Apps Script holds the real key
const DEFAULT_SHEET = '1BkMAMVD4iy4LWJHNNxCtmPntHSHCX0Y9_JR1ESHiZ0E';
const HISTORY = 'History';

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return out_({ ok: true, app: 'gym-log' });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_json' }); }
  if (!b || b.key !== KEY) return out_({ ok: false, error: 'bad_key' });
  let ss;
  try { ss = SpreadsheetApp.openById(b.sheetId || DEFAULT_SHEET); }
  catch (err) { return out_({ ok: false, error: 'no_sheet', message: String(err) }); }
  // Only gym sheets: never let the app touch any other spreadsheet.
  if (ss.getId() !== DEFAULT_SHEET && !/^gym/i.test(ss.getName())) return out_({ ok: false, error: 'no_sheet' });

  if (b.action === 'load') {
    try {
      const tabs = ss.getSheets().map(function (s) {
        return { name: s.getName(), hidden: s.isSheetHidden(), grid: s.getDataRange().getDisplayValues() };
      });
      return out_({ ok: true, title: ss.getName(), tabs: tabs });
    } catch (err) { return out_({ ok: false, error: 'load_failed', message: String(err) }); }
  }

  if (b.action !== 'write') return out_({ ok: false, error: 'bad_action' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName(b.tab);
    if (!sh) return out_({ ok: false, error: 'no_tab' });
    const cell = sh.getRange(b.a1);
    // Guard: if rows were added or moved in the sheet, don't write to the wrong exercise.
    if (b.expect) {
      const found = String(sh.getRange(cell.getRow(), b.expect.col).getDisplayValue()).trim();
      if (found !== String(b.expect.name).trim()) return out_({ ok: false, error: 'row_moved', found: found });
    }
    cell.setValue(asText_(b.value));
    let histRow = null;
    if (b.history && b.value) histRow = upsertHistory_(ss, b.history, b.value);
    return out_({ ok: true, histRow: histRow });
  } catch (err) {
    return out_({ ok: false, error: 'write_failed', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Keep values exactly as typed: stop Sheets turning "63" into a number or "10/8" into a date.
function asText_(v) {
  v = String(v == null ? '' : v);
  if (!v) return '';
  return /^[=+\-@]/.test(v) || /^[\d\s.,\/:-]+$/.test(v) ? "'" + v : v;
}

function normDate_(s) {
  const m = String(s).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : String(s).trim();
}

// One row per exercise, person and day: update today's row if it exists, otherwise add one.
function upsertHistory_(ss, h, value) {
  let sh = ss.getSheetByName(HISTORY);
  if (!sh) {
    sh = ss.insertSheet(HISTORY);
    sh.getRange(1, 1, 1, 5).setValues([['Date', 'Workout', 'Exercise', 'Person', 'Weight']]);
    sh.setFrozenRows(1);
  }
  if (!String(sh.getRange(1, 1).getDisplayValue()).trim()) {
    sh.getRange(1, 1, 1, 5).setValues([['Date', 'Workout', 'Exercise', 'Person', 'Weight']]);
  }
  const last = sh.getLastRow();
  if (last >= 2) {
    const vals = sh.getRange(2, 1, last - 1, 4).getDisplayValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (normDate_(vals[i][0]) === h.date && vals[i][2] === h.exercise && vals[i][3] === h.person) {
        sh.getRange(i + 2, 2).setValue(h.workout);
        sh.getRange(i + 2, 5).setValue(asText_(value));
        return i + 2;
      }
    }
  }
  const row = Math.max(last + 1, 2);
  sh.getRange(row, 1, 1, 5).setValues([["'" + h.date, h.workout, h.exercise, h.person, asText_(value)]]);
  return row;
}
