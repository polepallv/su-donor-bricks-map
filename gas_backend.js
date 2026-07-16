/**
 * ══════════════════════════════════════════════════════════════
 *  Donor Brick Monument — Google Apps Script Backend
 *  Paste this code into a new Google Apps Script project,
 *  then deploy as a Web App (see setup steps below).
 * ══════════════════════════════════════════════════════════════
 *
 *  SETUP STEPS:
 *  1. Go to https://script.google.com and create a new project.
 *  2. Delete any existing code and paste this entire file.
 *  3. Click Deploy → New deployment.
 *  4. Type: Web app
 *  5. Execute as: Me (your Google account)
 *  6. Who has access: Anyone
 *  7. Click Deploy and copy the URL (ends in /exec).
 *  8. Open section1.html, section2.html, and section3.html.
 *     Find the line:  const GAS_URL = '...';
 *     Replace the existing value with your new URL (in single quotes):
 *       const GAS_URL = 'https://script.google.com/macros/s/ABC.../exec';
 *  9. Save and re-host all three HTML files.
 *     From now on, every admin edit is saved to Google Sheets
 *     and visible to all visitors immediately.
 *
 *  RE-DEPLOYING AFTER CODE CHANGES:
 *  • If you change this script later, go to Deploy → Manage deployments,
 *    click the pencil icon, change Version to "New version", then Update.
 *    The URL stays the same.
 *
 * ══════════════════════════════════════════════════════════════
 *  HOW IT WORKS
 * ══════════════════════════════════════════════════════════════
 *
 *  All requests come in as GET (JSONP) to avoid CORS issues:
 *
 *  Fetch edits:
 *    ?action=get&section=1&callback=fn
 *    → Returns fn({"r,c": {lines:[…], size:1, updated:"2026-...Z"}, …})
 *    The `updated` timestamp lets the frontend prefer whichever of a
 *    browser's local edit or the shared sheet's edit is actually newer,
 *    instead of a local edit permanently overriding the shared one.
 *
 *  Save an edit:
 *    ?action=set&section=1&key=5,3&l1=NAME&l2=LINE2&l3=LINE3&l4=&l5=&l6=&l7=&l8=&size=1&callback=fn
 *    → Upserts a row in the BrickEdits sheet, returns fn({ok:true})
 *
 *  The Sheet ("BrickEdits") has columns:
 *    section | key | line1 | line2 | line3 | line4 | line5 | line6 | line7 | line8 | size | updated
 */

// ── Sheet name ──────────────────────────────────────────────
const SHEET_NAME = 'BrickEdits';

// ── Column indices (0-based) ─────────────────────────────────
const COL_SECTION = 0;
const COL_KEY     = 1;
const COL_L1      = 2;
const COL_L2      = 3;
const COL_L3      = 4;
const COL_L4      = 5;
const COL_L5      = 6;
const COL_L6      = 7;
const COL_L7      = 8;
const COL_L8      = 9;
const COL_SIZE    = 10;
const COL_UPDATED = 11;

/**
 * Handle all incoming requests.
 * Apps Script only calls doGet for web-app requests.
 */
function doGet(e) {
  const p      = e.parameter || {};
  const action = p.action || 'get';
  const cb     = p.callback || null;

  let result;

  try {
    if (action === 'set') {
      result = handleSet(p);
    } else {
      result = handleGet(p);
    }
  } catch (err) {
    result = { error: err.message };
  }

  const json = JSON.stringify(result);
  const output = cb
    ? cb + '(' + json + ')'
    : json;

  const mime = cb
    ? ContentService.MimeType.JAVASCRIPT
    : ContentService.MimeType.JSON;

  return ContentService.createTextOutput(output).setMimeType(mime);
}

/**
 * GET all edits for a section.
 * Returns { "row,col": { lines: [...], size: N }, ... }
 *
 * Uses getValues() (raw stored values), not getDisplayValues(): Sheets'
 * display layer silently hides a leading apostrophe on any cell (its
 * universal "treat as text" convention), which would otherwise strip
 * a real leading apostrophe from inscription text like "'88" or "'13".
 * The write side (handleSet) already forces plain-text number format
 * before writing, which is what actually prevents date/number
 * auto-conversion — getDisplayValues() here was redundant and harmful.
 */
function handleGet(p) {
  const section = String(p.section || '');
  const sheet   = getSheet();
  const data    = sheet.getDataRange().getValues();

  const result  = {};
  // Start at row 1 to skip header
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[COL_SECTION]) !== section) continue;

    const key   = String(row[COL_KEY]);
    const lines = [row[COL_L1], row[COL_L2], row[COL_L3], row[COL_L4],
                   row[COL_L5], row[COL_L6], row[COL_L7], row[COL_L8]]
                    .map(v => String(v || ''))
                    .filter(Boolean);
    const size    = parseInt(row[COL_SIZE]) || 1;
    const updated = String(row[COL_UPDATED] || '');

    result[key] = { lines, size, updated };
  }

  return result;
}

/**
 * SET (upsert) one brick edit.
 * Finds existing row for section+key; updates it or appends a new row.
 *
 * The target range's number format is forced to plain text ('@') before
 * writing, so Sheets doesn't auto-convert date-like inscription text
 * (e.g. "AUGUST 16, 2008") into an actual Date value.
 */
function handleSet(p) {
  const section = String(p.section || '');
  const key     = String(p.key     || '');
  const l1      = String(p.l1      || '');
  const l2      = String(p.l2      || '');
  const l3      = String(p.l3      || '');
  const l4      = String(p.l4      || '');
  const l5      = String(p.l5      || '');
  const l6      = String(p.l6      || '');
  const l7      = String(p.l7      || '');
  const l8      = String(p.l8      || '');
  const size    = parseInt(p.size) || 1;

  if (!section || !key) throw new Error('Missing section or key');

  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();

  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_SECTION]) === section &&
        String(data[i][COL_KEY])     === key) {
      foundRow = i + 1; // 1-indexed for Sheet.getRange
      break;
    }
  }

  const newRow  = [section, key, l1, l2, l3, l4, l5, l6, l7, l8, size, new Date().toISOString()];
  const rowIdx  = foundRow > 0 ? foundRow : sheet.getLastRow() + 1;
  const range   = sheet.getRange(rowIdx, 1, 1, newRow.length);

  range.setNumberFormat('@'); // plain text, prevents Sheets auto-converting dates/numbers
  range.setValues([newRow]);

  return { ok: true };
}

/**
 * Get (or create) the BrickEdits sheet.
 */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Write header row
    sheet.appendRow(['section', 'key', 'line1', 'line2', 'line3', 'line4',
                      'line5', 'line6', 'line7', 'line8', 'size', 'updated']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
  }

  return sheet;
}

/**
 * (Optional) Run this function once from the Apps Script editor to
 * verify the setup is working. Check the execution log for output.
 */
function testSetup() {
  const sheet = getSheet();
  Logger.log('Sheet found: ' + sheet.getName());
  Logger.log('Row count: ' + sheet.getLastRow());

  // Insert a test edit
  handleSet({ section:'1', key:'0,0', l1:'TEST BRICK', l2:'SETUP WORKS', l3:'', l4:'', l5:'', l6:'', l7:'', l8:'', size:'1' });
  Logger.log('Test write OK');

  // Read it back
  const result = handleGet({ section:'1' });
  Logger.log('Test read: ' + JSON.stringify(result));

  // Clean up test row
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SHEET_NAME);
  const d  = s.getDataRange().getValues();
  for (let i = d.length - 1; i >= 1; i--) {
    if (d[i][COL_L1] === 'TEST BRICK') { s.deleteRow(i + 1); break; }
  }
  Logger.log('Test cleanup OK — setup verified!');
}
