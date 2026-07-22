/**
 * ══════════════════════════════════════════════════════════════
 *  Donor Brick Garden — Google Apps Script Backend
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
 *  8. Open mccombs-pergola-bricks.html, lois-perkins-west.html, and lois-perkins-east.html.
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
 *    `lines` is returned exactly as stored — including blank entries in
 *    the middle of the array — so a deliberately blank line stays put
 *    instead of later lines shifting up to fill the gap.
 *
 *  Save an edit (requires the ADMIN_TOKEN below, or the write is rejected):
 *    ?action=set&section=1&key=5,3&l1=NAME&l2=&l3=LINE3&…&l12=&size=1&token=…&callback=fn
 *    → Upserts a row in the BrickEdits sheet, returns fn({ok:true})
 *
 *  The Sheet ("BrickEdits") has columns:
 *    section | key | line1..line12 | size | updated
 */

// ── Sheet name ──────────────────────────────────────────────
const SHEET_NAME = 'BrickEdits';

// ── Write-protection token ─────────────────────────────────
// The site's admin login prompt and "Edit plaques" UI are purely
// client-side — the GAS_URL is public (every visitor's browser needs
// it to read the map), so without this check anyone who found the URL
// could write directly to the Sheet, bypassing the site entirely. Every
// save must include this exact token; reads (action=get) stay open, no
// token needed, since the map itself is meant to be publicly viewable.
// Keep this in sync with the ADMIN_PASS constant in every page's HTML.
const ADMIN_TOKEN = 'SWU_Monument';

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
const COL_L9      = 10;
const COL_L10     = 11;
const COL_L11     = 12;
const COL_L12     = 13;
const COL_SIZE    = 14;
const COL_UPDATED = 15;
const LINE_COUNT  = 12;

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
 *
 * `lines` is NOT filtered down to just the non-blank entries: a blank
 * entry in the middle of the array (e.g. a deliberately skipped line
 * between two inscribed ones) is kept in place, so the frontend can
 * render it as an intentional gap instead of everything after it
 * shifting up to fill the hole.
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
    const lines = [
      row[COL_L1],  row[COL_L2],  row[COL_L3],  row[COL_L4],
      row[COL_L5],  row[COL_L6],  row[COL_L7],  row[COL_L8],
      row[COL_L9],  row[COL_L10], row[COL_L11], row[COL_L12],
    ].map(v => String(v || ''));
    // Drop only trailing blanks (nothing meaningful after the last
    // inscribed line) — interior blanks stay exactly where they are.
    while (lines.length && !lines[lines.length - 1]) lines.pop();

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
 *
 * Accepts up to 12 lines (l1..l12). Blank ones are stored as empty
 * strings in place — the caller is expected to have already trimmed
 * any meaningless trailing blanks, but interior blanks (a deliberately
 * skipped line) are written through as-is.
 *
 * Sheets treats a leading apostrophe on a written value as its own
 * "force text" input marker and consumes it — even via setValues() on an
 * already-plain-text-formatted cell, and even though a mid-string
 * apostrophe (e.g. "CLASS OF '13") is untouched. A line that itself
 * starts with an apostrophe (e.g. "'88 REUNION") needs an extra leading
 * apostrophe on write so the one Sheets consumes isn't the real one.
 *
 * Rejects the write outright if the request doesn't carry the correct
 * ADMIN_TOKEN — see the constant's comment above for why this exists.
 */
function handleSet(p) {
  if (String(p.token || '') !== ADMIN_TOKEN) throw new Error('Unauthorized');

  const section = String(p.section || '');
  const key     = String(p.key     || '');
  const lines   = [];
  for (let i = 1; i <= LINE_COUNT; i++) {
    let line = String(p['l' + i] || '');
    if (line.charAt(0) === "'") line = "'" + line;
    lines.push(line);
  }
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

  const newRow  = [section, key].concat(lines, [size, new Date().toISOString()]);
  const rowIdx  = foundRow > 0 ? foundRow : sheet.getLastRow() + 1;
  const range   = sheet.getRange(rowIdx, 1, 1, newRow.length);

  range.setNumberFormat('@'); // plain text, prevents Sheets auto-converting dates/numbers
  range.setValues([newRow]);

  return { ok: true };
}

/**
 * Get (or create) the BrickEdits sheet. If the sheet already exists from
 * before the 12-line schema (8 line columns instead of 12), its header
 * row is extended in place with the 4 missing line columns so existing
 * data isn't disturbed.
 */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  const headers = ['section', 'key',
    'line1', 'line2', 'line3', 'line4', 'line5', 'line6',
    'line7', 'line8', 'line9', 'line10', 'line11', 'line12',
    'size', 'updated'];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return sheet;
  }

  // Upgrade an older 8-line sheet in place: if there are fewer header
  // columns than expected, insert 4 blank columns before the old
  // size/updated columns and relabel the whole header row.
  const existingCols = sheet.getLastColumn();
  if (existingCols > 0 && existingCols < headers.length) {
    sheet.insertColumnsBefore(existingCols - 1, headers.length - existingCols);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
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

  // Insert a test edit, with a deliberately blank line 2 to check that
  // interior blanks round-trip correctly.
  handleSet({ section: '1', key: '0,0', l1: 'TEST BRICK', l2: '', l3: 'SETUP WORKS', size: '1', token: ADMIN_TOKEN });
  Logger.log('Test write OK');

  // Read it back
  const result = handleGet({ section: '1' });
  Logger.log('Test read: ' + JSON.stringify(result));
  const readBack = result['0,0'];
  if (!readBack || readBack.lines[1] !== '') {
    Logger.log('WARNING: blank interior line did not round-trip as expected!');
  } else {
    Logger.log('Interior blank line preserved correctly.');
  }

  // Clean up test row
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SHEET_NAME);
  const d  = s.getDataRange().getValues();
  for (let i = d.length - 1; i >= 1; i--) {
    if (d[i][COL_L1] === 'TEST BRICK') { s.deleteRow(i + 1); break; }
  }
  Logger.log('Test cleanup OK — setup verified!');
}
