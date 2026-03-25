function syncRequestsToControl() {
  const SOURCE_SPREADSHEET_ID = 'YOUR_REQUESTS_SPREADSHEET_ID';
  const TARGET_SPREADSHEET_ID = 'YOUR_CONTROL_SPREADSHEET_ID';
  const SOURCE_SHEET_NAME     = 'Página1';
  const TARGET_SHEET_NAME     = 'Imagens Yunique';

  // ── Open spreadsheets ─────────────────────────────────────────────────────
  const ssSource = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  const ssTarget = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);

  const sourceSheet = ssSource.getSheetByName(SOURCE_SHEET_NAME);
  const targetSheet = ssTarget.getSheetByName(TARGET_SHEET_NAME);

  if (!sourceSheet) throw new Error(`Sheet "${SOURCE_SHEET_NAME}" not found in Requests spreadsheet.`);
  if (!targetSheet) throw new Error(`Sheet "${TARGET_SHEET_NAME}" not found in Control spreadsheet.`);

  // ── Date formatting helpers ───────────────────────────────────────────────
  function pad(n) { return String(n).padStart(2, '0'); }

  // Format: dd-MM-yyyy HH:mm  (used for Start-Date and End-Date)
  function formatDateTime(val) {
    if (!val) return '';
    const d = new Date(val);
    if (isNaN(d)) return String(val);
    return `${pad(d.getDate())}-${pad(d.getMonth()+1)}-${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }

  // Format: dd/MM/yyyy  (used for Request-Date)
  function formatDate(val) {
    if (!val) return '';
    const d = new Date(val);
    if (isNaN(d)) return String(val);
    return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()}`;
  }

  // Format: HH:mm  (used for Request-Time)
  function formatTime(val) {
    if (!val) return '';
    if (typeof val === 'string') return val.substring(0, 5);
    const d = new Date(val);
    if (isNaN(d)) return String(val);
    return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }

  // ── Read all source data (batch) ──────────────────────────────────────────
  const totalSourceRows = sourceSheet.getLastRow();
  if (totalSourceRows < 2) {
    Logger.log('No data found in the Requests sheet.');
    return;
  }

  const sourceData = sourceSheet
    .getRange(2, 1, totalSourceRows - 1, 14)
    .getValues();

  // ── Build Set of existing Client-Protocols in target (col B = col 2) ──────
  const lastTargetRow = targetSheet.getLastRow();
  const existingProtocols = new Set();

  if (lastTargetRow >= 2) {
    targetSheet
      .getRange(2, 2, lastTargetRow - 1, 1)
      .getValues()
      .forEach(([p]) => { if (p) existingProtocols.add(String(p).trim()); });
  }

  Logger.log(`Protocols already in Control: ${existingProtocols.size}`);

  // ── Filter and build new rows ─────────────────────────────────────────────
  const newRows = [];

  sourceData.forEach(row => {
    // Source column mapping (0-based index):
    // 0  = Protocol-Yunique   (full ID string)
    // 1  = Client-Protocol    (numeric — dedup key)
    // 2  = Requester
    // 3  = Condominium
    // 4  = Partner
    // 5  = Email
    // 6  = Occurrence-Type
    // 7  = SLA
    // 8  = Syndic
    // 9  = Requested-Camera-Location
    // 10 = Start-Date
    // 11 = End-Date
    // 12 = Request-Date
    // 13 = Request-Time

    const protocolYunique = String(row[0] || '').trim();
    const clientProtocol  = String(row[1] || '').trim();

    if (!clientProtocol) return;                        // skip empty rows
    if (existingProtocols.has(clientProtocol)) return;  // skip duplicates

    existingProtocols.add(clientProtocol); // prevent in-loop duplicates

    newRows.push([
      protocolYunique,          // A – Protocol-Yunique
      clientProtocol,           // B – Client-Protocol
      row[2],                   // C – Requester
      row[3],                   // D – Condominium
      row[4],                   // E – Partner
      row[5],                   // F – Email
      row[6],                   // G – Occurrence-Type
      row[7],                   // H – SLA
      row[8],                   // I – Syndic
      row[9],                   // J – Requested-Camera-Location
      formatDateTime(row[10]),  // K – Start-Date      → dd-MM-yyyy HH:mm
      formatDateTime(row[11]),  // L – End-Date        → dd-MM-yyyy HH:mm
      formatDate(row[12]),      // M – Request-Date    → dd/MM/yyyy
      formatTime(row[13]),      // N – Request-Time    → HH:mm
      '',                       // O – STATUS          (intentionally blank)
      '',                       // P – Send-Date       (blank)
      '',                       // Q – Notes           (blank)
      '', '', '',               // R, S, T             (extra columns, blank)
    ]);
  });

  // ── Write all new rows at once (batch write) ──────────────────────────────
  if (newRows.length === 0) {
    Logger.log('✅ Nothing to sync — all entries are already up to date.');
    return;
  }

  const firstEmptyRow = targetSheet.getLastRow() + 1;
  targetSheet
    .getRange(firstEmptyRow, 1, newRows.length, 20)
    .setValues(newRows);

  Logger.log(`✅ ${newRows.length} new row(s) inserted successfully.`);
}
