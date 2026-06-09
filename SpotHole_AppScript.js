// ── Spot Hole — Google Apps Script ───────────────────────────────────────
// Deploy as a Web App:
//   Execute as: Me
//   Who has access: Anyone
//
// Fill in the two constants below, then deploy.

const SHEET_ID = '1BXHayOuxIZd7H8Oqq5NRRxiFq90teX0k2pdfX1N_wmU';
const SECRET   = 'spotholekey';         // must match SECRET in Spothole_web_app.html

// ─────────────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    // Reject wrong secret
    if (payload.secret !== SECRET) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss        = SpreadsheetApp.openById(SHEET_ID);
    const sheetName = payload.sheetName || getTodayName();

    // Lock so two simultaneous first-writes don't both create the tab + header
    const lock = LockService.getScriptLock();
    lock.tryLock(10000);
    let sheet;
    try {
      sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.appendRow([
          'session_id','time','latitude','longitude',
          'speed_mph','heading_deg','heading_compass','x_accel_g','y_accel_g','z_accel_g','type','device','bump_id'
        ]);
      }
    } finally {
      lock.releaseLock();
    }

    // Append rows (appendRow is safe for concurrent callers — no overwrites)
    const rows = payload.rows || [];
    rows.forEach(row => sheet.appendRow(row));

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, sheet: sheetName, appended: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getDateKey() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
}

// ── GET endpoint — returns real row count for a sheet tab ─────────────────
// URL: <SCRIPT_URL>?date=YYYYMMDD&type=events   (type = events|waveforms|highres)
// No secret required — read-only, row counts only.
function doGet(e) {
  try {
    const date  = (e.parameter && e.parameter.date)  || getDateKey();
    const type  = (e.parameter && e.parameter.type)  || 'events';
    const sheetName = date + '_' + type;
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    const rows  = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0; // subtract header
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, sheet: sheetName, rows }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Test this script manually ─────────────────────────────────────────────
function testPost() {
  const result = doPost({
    postData: {
      contents: JSON.stringify({
        secret: SECRET,
        sheetName: getDateKey() + '_events',
        rows: [
          ['test_session','2024-01-01T00:00:00Z',40.7128,-74.0060,
           25.5,180,'S',0.05,0.02,0.87,'session_start','Test / Script','']
        ]
      })
    }
  });
  Logger.log(result.getContent());
}
