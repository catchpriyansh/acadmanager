/**
 * AcadManager — Google Apps Script Backend
 * Dr. Priyansh Singh | IIT Indore
 *
 * SETUP INSTRUCTIONS (do this once):
 * ─────────────────────────────────────────────────────────────────
 * 1. Go to https://sheets.google.com and create a new spreadsheet
 *    named "AcadManager Database"
 *
 * 2. In that spreadsheet, open Extensions → Apps Script
 *
 * 3. Delete any existing code and paste THIS entire file
 *
 * 4. Click Save (Ctrl+S), then click Deploy → New Deployment
 *    - Type: Web App
 *    - Execute as: Me
 *    - Who has access: Anyone  ← important (the app uses a password for security)
 *
 * 5. Click Deploy → copy the Web App URL
 *
 * 6. Open your AcadManager app → click "⬆ Sync" button → paste the URL
 *
 * That's it! Your data will now sync to Google Sheets.
 * ─────────────────────────────────────────────────────────────────
 */

const SHEET_NAMES = {
  students:   'Students',
  lectures:   'Lectures',
  attendance: 'Attendance',
  tasks:      'Tasks',
  meta:       'SyncLog',
};

// ── GET handler (load data) ──────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action || 'getAll';
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'getAll') {
      const result = {
        students:   readSheet(ss, SHEET_NAMES.students),
        lectures:   readSheet(ss, SHEET_NAMES.lectures),
        attendance: readSheet(ss, SHEET_NAMES.attendance),
        tasks:      readSheet(ss, SHEET_NAMES.tasks),
      };
      return jsonResponse(result);
    }

    return jsonResponse({ error: 'Unknown action' }, 400);
  } catch(err) {
    return jsonResponse({ success: false, error: err.message }, 500);
  }
}

// ── POST handler (save data) ─────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'saveAll') {
      writeSheet(ss, SHEET_NAMES.students,   data.students   || []);
      writeSheet(ss, SHEET_NAMES.lectures,   data.lectures   || []);
      writeSheet(ss, SHEET_NAMES.attendance, data.attendance || []);
      writeSheet(ss, SHEET_NAMES.tasks,      data.tasks      || []);
      logSync(ss, data.meta);
      return jsonResponse({ success: true, timestamp: new Date().toISOString() });
    }

    return jsonResponse({ error: 'Unknown action' }, 400);
  } catch(err) {
    return jsonResponse({ success: false, error: err.message }, 500);
  }
}

// ── Helpers ──────────────────────────────────────────────────────

function readSheet(ss, name) {
  const sh = getOrCreateSheet(ss, name);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0];
  return vals.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function writeSheet(ss, name, records) {
  const sh = getOrCreateSheet(ss, name);
  sh.clearContents();
  if (!records || records.length === 0) return;

  const headers = Object.keys(records[0]);
  const rows = records.map(r => headers.map(h => r[h] !== undefined ? r[h] : ''));

  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1B3A6B')
    .setFontColor('#FFFFFF');

  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // Auto-resize columns
  headers.forEach((_, i) => sh.autoResizeColumn(i + 1));
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function logSync(ss, meta) {
  const sh = getOrCreateSheet(ss, SHEET_NAMES.meta);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,4).setValues([['SyncedAt','Semester','Year','Notes']]).setFontWeight('bold');
  }
  sh.appendRow([
    new Date().toISOString(),
    meta && meta.semester ? meta.semester : '',
    meta && meta.year ? meta.year : '',
    'Auto sync from AcadManager'
  ]);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
