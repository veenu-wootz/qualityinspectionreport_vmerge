/**
 * googleSheets.js
 * Appends a row to the 'Checkin' subsheet of the configured Google Sheet.
 * Uses the same service account as googleDrive.js.
 *
 * Sheet column mapping (A=0 index):
 *   A  (0)  — CheckIn ID   → data.report_no
 *   B  (1)  — ID           → data.report_no
 *   C  (2)  — ''
 *   D  (3)  — ''
 *   E  (4)  — Project name → '' (for now)
 *   F  (5)  — Part number  → data.part_number
 *   G  (6)  — Assembly drawing → data.part_drawing || ''
 *   H  (7)  — Status       → 'Update'
 *   I  (8)  — Description  → data.remarks || ''
 *   J  (9)  — Inspection Image → data.insp_image || ''
 *   K  (10) — ''
 *   L  (11) — Created by   → data.created_by || ''
 *   M  (12) — Timestamp    → ISO timestamp
 *   N–Z (13–25) — '' (13 empty cols)
 *   AA (26) — Files        → driveUrl
 *
 * Env vars required:
 *   GOOGLE_SERVICE_ACCOUNT_JSON — full service account JSON key (as string)
 *   GOOGLE_SHEET_ID             — sheet ID from the Google Sheets URL
 */

const { google } = require('googleapis');

function getAuthClient() {
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!raw) throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON env var not set');

  let credentials;
  try {
    credentials = JSON.parse(raw);
  } catch (e) {
    throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON: ' + e.message);
  }

  return new google.auth.GoogleAuth({
    credentials,
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
    ],
  });
}

/**
 * Append one row to the Checkin subsheet.
 * @param {object} data     — parsed QIR payload
 * @param {string} driveUrl — uploaded PDF Drive URL
 */
async function appendToSheet(data, driveUrl) {
  const sheetId = process.env.GOOGLE_SHEET_ID;
  if (!sheetId) throw new Error('GOOGLE_SHEET_ID env var not set');

  const auth   = getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth });

  const timestamp = new Date().toISOString();

  // Build the row — 27 elements (A through AA)
  // Empty strings for unused columns to preserve alignment
  const row = [
    data.report_no        || '',  // A  — CheckIn ID
    data.report_no        || '',  // B  — ID
    '',                           // C  — (unused)
    '',                           // D  — (unused)
    '',                           // E  — Project name (later)
    data.part_number      || '',  // F  — Part number
    data.part_drawing     || '',  // G  — Assembly drawing
    'Update',                     // H  — Status
    data.remarks          || '',  // I  — Description
    data.insp_image       || '',  // J  — Inspection Image
    '',                           // K  — (unused)
    data.created_by       || '',  // L  — Created by
    timestamp,                    // M  — Timestamp
    '', '', '', '', '',           // N, O, P, Q, R
    '', '', '', '', '',           // S, T, U, V, W
    '', '', '',                   // X, Y, Z
    driveUrl              || '',  // AA — Files (PDF Drive link)
  ];

  console.log(`  Appending row to Checkin sheet (${row.length} columns)...`);

  await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range:         'Checkin!A:AA',   // targets the Checkin subsheet
    valueInputOption: 'USER_ENTERED', // preserves URLs as clickable links
    insertDataOption: 'INSERT_ROWS',  // always adds a new row, never overwrites
    requestBody: {
      values: [row],
    },
  });

  console.log(`  Row appended to Checkin sheet for report: ${data.report_no}`);
}

module.exports = { appendToSheet };
