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

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheets()[0];

    // Write header if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'session_id','time','latitude','longitude',
        'speed_mph','heading_deg','heading_compass','z_accel_g','type'
      ]);
    }

    // Append each row
    const rows = payload.rows || [];
    rows.forEach(row => sheet.appendRow(row));

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, appended: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Test this script manually ─────────────────────────────────────────────
// Run testPost() from the editor to verify it can write to your sheet
// before deploying.
function testPost() {
  const result = doPost({
    postData: {
      contents: JSON.stringify({
        secret: SECRET,
        rows: [
          ['test_session','2024-01-01T00:00:00Z',40.7128,-74.0060,
           25.5,180,'S',0.05,'spot_check']
        ]
      })
    }
  });
  Logger.log(result.getContent());
}
