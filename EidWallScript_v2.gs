// ── EID WISH WALL — Google Apps Script v2 ───────────────────────────
// IMPORTANT: Replace your existing script with this version.
// Steps:
//   1. Open script.google.com → your "Eid Wish Wall" project
//   2. Delete everything, paste this entire file, Save
//   3. Deploy → New Deployment (do a NEW deployment, don't just save)
//      - Type: Web App
//      - Execute as: Me
//      - Who has access: Anyone
//   4. Copy the new Web App URL and update API in eid.html
// ─────────────────────────────────────────────────────────────────────

const SHEET_NAME = 'Wishes';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Name', 'City', 'Message', 'ID']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ALL requests come in as GET to avoid CORS preflight.
// action=post  → save a wish
// (no action)  → fetch all wishes
function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();

  if (action === 'post') {
    // ── SAVE WISH ──
    try {
      const name = (e.parameter.name || '').toString().trim().slice(0, 30);
      const city = (e.parameter.city || '').toString().trim().slice(0, 30);
      const msg  = (e.parameter.msg  || '').toString().trim().slice(0, 180);
      if (!name || !msg) {
        return jsonResponse({ error: 'Name and message are required' });
      }
      const id    = Utilities.getUuid();
      const sheet = getSheet();
      sheet.appendRow([new Date(), name, city, msg, id]);
      return jsonResponse({ ok: true, id });
    } catch(err) {
      return jsonResponse({ error: err.message });
    }

  } else {
    // ── FETCH WISHES ──
    try {
      const sheet = getSheet();
      const data  = sheet.getDataRange().getValues();
      if (data.length <= 1) {
        return jsonResponse({ wishes: [] });
      }
      const wishes = data.slice(1).reverse().map(row => ({
        ts:   new Date(row[0]).getTime(),
        name: row[1],
        city: row[2],
        msg:  row[3],
        id:   row[4]
      }));
      return jsonResponse({ wishes });
    } catch(err) {
      return jsonResponse({ error: err.message });
    }
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
