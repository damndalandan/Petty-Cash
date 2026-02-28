// =============================================
// PETTY CASH SYSTEM — Google Apps Script
// v4.0 — BIR-compliant, date-based IDs, full audit trail
// =============================================

const SPREADSHEET_ID = '12RGOYbXlHz70wtVskB_SNRjO1pRuhN9pdwauqzBEI-E';

const SHEETS = {
  ENTRIES      : 'PettyCash_Entries',
  DENOMINATIONS: 'PettyCash_Denominations',
  SUMMARY      : 'PettyCash_Summary',
  RECEIPTS     : 'PettyCash_Receipts',
  NO_RECEIPTS  : 'PettyCash_NoReceipts',
  ACCESS       : 'PettyCash_Access'
};

// ─────────────────────────────────────────────
// WEB APP ENTRY POINT
// ─────────────────────────────────────────────
function isAuthorized(email) {
  if (!email || email === 'unknown') return false;

  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ACCESS);

    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const rowEmail  = String(data[i][0] || '').trim().toLowerCase();
      const rowStatus = String(data[i][3] || '').trim().toUpperCase();

      if (rowEmail === email.toLowerCase() && rowStatus === 'ACTIVE') {
        return true;
      }
    }

    return false;
  } catch(e) {
    console.error('isAuthorized error:', e);
    return false;
  }
}

function getUserEmail() {
  try { return Session.getActiveUser().getEmail() || 'unknown'; }
  catch(e) { return 'unknown'; }
}

function doGet() {
  const email = getUserEmail();

  if (!isAuthorized(email)) {
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <title>Access Denied — Petty Cash</title>
          <style>
            * { box-sizing: border-box; margin: 0; padding: 0; }
            body {
              font-family: 'Segoe UI', sans-serif;
              background: #f3f4f6;
              display: flex;
              align-items: center;
              justify-content: center;
              min-height: 100vh;
              padding: 1rem;
            }
            .card {
              background: white;
              border-radius: 1rem;
              padding: 2.5rem;
              max-width: 420px;
              width: 100%;
              text-align: center;
              box-shadow: 0 4px 24px rgba(0,0,0,0.08);
              border-top: 4px solid #ef4444;
            }
            .icon { font-size: 3rem; margin-bottom: 1rem; }
            h1 { font-size: 1.25rem; font-weight: 700; color: #1f2937; margin-bottom: 0.5rem; }
            p { font-size: 0.875rem; color: #6b7280; line-height: 1.6; margin-bottom: 0.5rem; }
            .email {
              display: inline-block;
              margin-top: 0.75rem;
              padding: 0.4rem 0.75rem;
              background: #fef2f2;
              border: 1px solid #fecaca;
              border-radius: 0.5rem;
              font-family: monospace;
              font-size: 0.8rem;
              color: #dc2626;
            }
          </style>
        </head>
        <body>
          <div class="card">
            <div class="icon">🚫</div>
            <h1>Access Denied</h1>
            <p>You don't have permission to access the Petty Cash System.</p>
            <p>Please contact your administrator to request access.</p>
            ${email !== 'unknown' ? `<div class="email">${email}</div>` : ''}
          </div>
        </body>
      </html>
    `)
    .setTitle('Access Denied — Petty Cash')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  initializeSheets();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Petty Cash System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────

function normalizeDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return new Date(val.getTime() - (val.getTimezoneOffset() * 60000))
      .toISOString().split('T')[0];
  }
  return String(val).split('T')[0];
}

/**
 * Generates a human-readable, date-based ID with daily incrementing counter.
 *
 * Format: PREFIX-MMMDDYY-NN
 * Example: EXP-JAN0226-01, RCP-FEB1526-03
 *
 * @param {string}  prefix     - ID prefix e.g. 'EXP', 'RCP', 'NRC', 'DEN-OC'
 * @param {string}  date       - ISO date string 'YYYY-MM-DD'
 * @param {Sheet}   sheet      - The Apps Script Sheet object to scan for existing IDs
 * @returns {string}
 */
function generateId(prefix, date, sheet) {
  const MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN',
                  'JUL','AUG','SEP','OCT','NOV','DEC'];

  // Parse the date
  const dateObj = new Date(date + 'T00:00:00');
  const mon     = MONTHS[dateObj.getMonth()];
  const dd      = String(dateObj.getDate()).padStart(2, '0');
  const yy      = String(dateObj.getFullYear()).slice(-2);

  // Build the date segment: MMMDDYY  e.g. JAN0226
  const dateSeg = `${mon}${dd}${yy}`;
  // The prefix portion to match: e.g. "EXP-JAN0226-"
  const matchPrefix = `${prefix}-${dateSeg}-`;

  // Scan the sheet's first column for existing IDs on this date with this prefix
  let maxIncrement = 0;

  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      ids.forEach(([id]) => {
        if (typeof id === 'string' && id.startsWith(matchPrefix)) {
          const suffix = id.slice(matchPrefix.length);
          const num    = parseInt(suffix, 10);
          if (!isNaN(num) && num > maxIncrement) maxIncrement = num;
        }
      });
    }
  }

  const nextNum = String(maxIncrement + 1).padStart(2, '0');
  return `${prefix}-${dateSeg}-${nextNum}`;
}

/**
 * Compute Philippine VAT breakdown from a gross (VAT-inclusive) amount.
 * VAT rate: 12%
 */
function computeVAT(grossAmount) {
  const gross        = parseFloat(grossAmount) || 0;
  const vatableSales = gross / 1.12;
  const vatAmount    = gross - vatableSales;
  return {
    grossAmount  : parseFloat(gross.toFixed(2)),
    vatableSales : parseFloat(vatableSales.toFixed(2)),
    vatAmount    : parseFloat(vatAmount.toFixed(2))
  };
}

function calculateDenomTotal(denominations) {
  return [1000, 500, 200, 100, 50, 20, 10, 5, 1, 0.25].reduce((sum, bill) => {
    return sum + (parseFloat(denominations[String(bill)]) || 0) * bill;
  }, 0);
}

function formatHeaderRow(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  range.setBackground('#1a1a2e');
  range.setFontColor('#ffffff');
  range.setFontWeight('bold');
}

// ─────────────────────────────────────────────
// SHEET INITIALIZATION
// ─────────────────────────────────────────────
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  if (!ss.getSheetByName(SHEETS.ENTRIES)) {
    const s = ss.insertSheet(SHEETS.ENTRIES);
    s.appendRow([
      'Entry_ID','Date','Type','Category','Description',
      'Amount','Has_Receipt','Reference_No','Requested_By','Approved_By',
      'Status','Created_At','Updated_At','Deleted_At'
    ]);
    s.setFrozenRows(1);
    formatHeaderRow(s);
    s.getRange('F2:F').setNumberFormat('₱#,##0.00');
    s.setColumnWidths(1, 14, 120);
    s.setColumnWidth(5, 220);
  }

  if (!ss.getSheetByName(SHEETS.DENOMINATIONS)) {
    const s = ss.insertSheet(SHEETS.DENOMINATIONS);
    s.appendRow([
      'Record_ID','Date','Type',
      'D_1000','D_500','D_200','D_100','D_50',
      'D_20','D_10','D_5','D_1','D_025',
      'Total','Notes','Created_At'
    ]);
    s.setFrozenRows(1);
    formatHeaderRow(s);
    s.getRange('N2:N').setNumberFormat('₱#,##0.00');
  }

  if (!ss.getSheetByName(SHEETS.SUMMARY)) {
    const s = ss.insertSheet(SHEETS.SUMMARY);
    s.appendRow([
      'Summary_ID','Date','Opening_Cash','Cash_Advance',
      'Total_Exp_With_Receipt','Total_Exp_No_Receipt','Total_Expenses',
      'Total_Cash_Over','Total_Replenishment',
      'Closing_Cash','Variance','Status','Closed_By','Updated_At'
    ]);
    s.setFrozenRows(1);
    formatHeaderRow(s);
    s.getRange('C2:K').setNumberFormat('₱#,##0.00');
  }

  // ── PettyCash_Receipts (BIR Purchases Journal) ──────
  if (!ss.getSheetByName(SHEETS.RECEIPTS)) {
    const s = ss.insertSheet(SHEETS.RECEIPTS);
    s.appendRow([
      'Receipt_ID','Entry_ID','Date','Supplier_Name',
      'Address','TIN','Receipt_No',
      'Gross_Amount','Vatable_Sales','VAT_Amount',
      'Created_By','Created_At'
    ]);
    s.setFrozenRows(1);
    formatHeaderRow(s);
    s.getRange('H2:J').setNumberFormat('₱#,##0.00');
    s.setColumnWidth(4, 200);
    s.setColumnWidth(5, 200);
    s.setColumnWidth(6, 140);
  }

  // ── PettyCash_NoReceipts ──────────────────────────────
  if (!ss.getSheetByName(SHEETS.NO_RECEIPTS)) {
    const s = ss.insertSheet(SHEETS.NO_RECEIPTS);
    s.appendRow([
      'NR_ID','Entry_ID','Date','Description',
      'Amount','Requested_By','Created_By','Created_At'
    ]);
    s.setFrozenRows(1);
    formatHeaderRow(s);
    s.getRange('E2:E').setNumberFormat('₱#,##0.00');
    s.setColumnWidth(1, 160);
    s.setColumnWidth(2, 160);
    s.setColumnWidth(3, 120);
    s.setColumnWidth(4, 260);
    s.setColumnWidth(5, 120);
    s.setColumnWidth(6, 160);
    s.setColumnWidth(7, 200);
    s.setColumnWidth(8, 180);
  }

  if (!ss.getSheetByName(SHEETS.ACCESS)) {
    const s = ss.insertSheet(SHEETS.ACCESS);
    s.appendRow(['Email', 'Name', 'Role', 'Status', 'Added_At', 'Notes']);
    s.setFrozenRows(1);
    formatHeaderRow(s);
    s.setColumnWidth(1, 260);
    s.setColumnWidth(2, 180);
    s.setColumnWidth(3, 120);
    s.setColumnWidth(4, 100);
    s.setColumnWidth(5, 180);
    s.setColumnWidth(6, 240);

    s.getRange('A1').setNote('Required. Full Google account email e.g. juan@gmail.com');
    s.getRange('C1').setNote('Role label e.g. Admin, Cashier, Viewer — for reference only');
    s.getRange('D1').setNote('ACTIVE or INACTIVE — only ACTIVE accounts can log in');

    s.appendRow([
      'example@gmail.com',
      'Juan Dela Cruz',
      'Cashier',
      'INACTIVE',
      new Date().toISOString(),
      'Sample row — set Status to ACTIVE to enable access'
    ]);

    s.getRange(2, 1, 1, 6).setFontColor('#9ca3af').setFontStyle('italic');
    s.getRange('D2:D').setHorizontalAlignment('center');
  }

  return { success: true };
}

// ─────────────────────────────────────────────
// SUMMARY RECALCULATION
// ─────────────────────────────────────────────
function recalculateDailySummary(date) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const entryData = ss.getSheetByName(SHEETS.ENTRIES).getDataRange().getValues();
    let totalExp = 0, totalReceipt = 0, totalNoReceipt = 0,
        cashAdvance = 0, totalCashOver = 0, totalReplenishment = 0;

    for (let i = 1; i < entryData.length; i++) {
      const row   = entryData[i];
      const rDate = normalizeDate(row[1]);
      if (rDate !== date || row[10] === 'DELETED') continue;

      const amt  = parseFloat(row[5]) || 0;
      const type = row[2];

      if (type === 'CASH_ADVANCE') {
        cashAdvance += amt;
      } else if (type === 'CASH_OVER') {
        totalCashOver += amt;
      } else if (type === 'REPLENISHMENT') {
        totalReplenishment += amt;
      } else {
        totalExp += amt;
        if (row[6] === 'YES') totalReceipt   += amt;
        else                  totalNoReceipt += amt;
      }
    }

    const denomData = ss.getSheetByName(SHEETS.DENOMINATIONS).getDataRange().getValues();
    let openingCash = 0, closingCash = 0, hasClosing = false;

    for (let i = 1; i < denomData.length; i++) {
      if (normalizeDate(denomData[i][1]) !== date) continue;
      const type  = denomData[i][2];
      const total = parseFloat(denomData[i][12]) || 0;
      if (type === 'START') openingCash = total;
      if (type === 'END')   { closingCash = total; hasClosing = true; }
    }

    const expected = openingCash - (totalExp + cashAdvance);
    const variance = closingCash - expected;
    const status   = hasClosing ? 'CLOSED' : 'OPEN';

    const sumSheet = ss.getSheetByName(SHEETS.SUMMARY);
    const sumData  = sumSheet.getDataRange().getValues();
    let targetRow  = -1, existingClosedBy = '';

    for (let i = 1; i < sumData.length; i++) {
      if (normalizeDate(sumData[i][1]) === date) {
        targetRow = i + 1;
        existingClosedBy = sumData[i][12] || '';
        break;
      }
    }

    const summaryRow = [
      openingCash, cashAdvance,
      totalReceipt, totalNoReceipt, totalExp,
      totalCashOver, totalReplenishment,
      closingCash, variance, status,
      existingClosedBy,
      new Date().toISOString()
    ];

    if (targetRow === -1) {
      const sumId = generateId('SUM', date, sumSheet);
      sumSheet.appendRow([sumId, date, ...summaryRow]);
    } else {
      sumSheet.getRange(targetRow, 3, 1, 12).setValues([summaryRow]);
    }

    return { success: true };
  } catch(e) {
    console.error('recalculateDailySummary:', e);
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
// EXPENSE ENTRIES CRUD
// ─────────────────────────────────────────────
function saveExpenseEntry(data) {
  try {
    const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
    const entrySheet = ss.getSheetByName(SHEETS.ENTRIES);
    const now        = new Date().toISOString();
    const id         = generateId('EXP', data.date, entrySheet);

    entrySheet.appendRow([
      id,
      data.date,
      data.type        || 'EXPENSE',
      data.category    || 'Miscellaneous',
      data.description || '',
      parseFloat(data.amount) || 0,
      data.hasReceipt ? 'YES' : 'NO',
      data.referenceNo || '',
      data.requestedBy || '',
      data.approvedBy  || '',
      'ACTIVE',
      now,
      now,
      ''
    ]);

    // ── Auto-log to PettyCash_NoReceipts if no receipt and type is EXPENSE ──
    if (!data.hasReceipt && (data.type === 'EXPENSE' || !data.type)) {
      saveNoReceiptRecord({
        entryId    : id,
        date       : data.date,
        description: data.description || '',
        amount     : parseFloat(data.amount) || 0,
        requestedBy: data.requestedBy || ''
      });
    }

    recalculateDailySummary(data.date);
    return { success: true, id };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function saveCashOverEntry(data) {
  return saveExpenseEntry({ ...data, type: 'CASH_OVER', hasReceipt: false });
}

function saveReplenishmentEntry(data) {
  return saveExpenseEntry({ ...data, type: 'REPLENISHMENT', hasReceipt: false });
}

function updateExpenseEntry(payload) {
  try {
    const id = payload.id;
    if (!id) return { success: false, message: 'No Entry_ID provided' };

    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet     = ss.getSheetByName(SHEETS.ENTRIES);
    const dataRange = sheet.getDataRange().getValues();

    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] !== id) continue;

      const oldDate = normalizeDate(dataRange[i][1]);
      const row     = i + 1;
      const now     = new Date().toISOString();

      sheet.getRange(row, 2, 1, 10).setValues([[
        payload.date,
        payload.type        || 'EXPENSE',
        payload.category    || '',
        payload.description || '',
        parseFloat(payload.amount) || 0,
        payload.hasReceipt ? 'YES' : 'NO',
        payload.referenceNo || '',
        payload.requestedBy || '',
        payload.approvedBy  || '',
        'ACTIVE'
      ]]);
      sheet.getRange(row, 13).setValue(now);

      // ── Sync NoReceipts sheet on update ──
      // If receipt status changed TO no-receipt, add a record if not already there
      // If changed TO has-receipt, remove from NoReceipts
      syncNoReceiptOnUpdate(payload, dataRange[i]);

      recalculateDailySummary(payload.date);
      if (oldDate && oldDate !== payload.date) recalculateDailySummary(oldDate);

      return { success: true };
    }
    return { success: false, message: 'Entry not found' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Keeps PettyCash_NoReceipts in sync when an entry is edited.
 * - Was receipted → now no receipt:  add a NRC row
 * - Was no receipt → now receipted:  soft-delete the NRC row (mark RECEIPTED)
 * - Still no receipt + fields changed: update the NRC row
 */
function syncNoReceiptOnUpdate(newPayload, oldRow) {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const nrSheet = ss.getSheetByName(SHEETS.NO_RECEIPTS);
    if (!nrSheet) return;

    const wasReceipted = oldRow[6] === 'YES';
    const nowReceipted = !!newPayload.hasReceipt;
    const entryId      = newPayload.id;
    const isExpense    = (newPayload.type === 'EXPENSE' || !newPayload.type);

    if (!isExpense) return; // CASH_OVER / REPLENISHMENT / CASH_ADVANCE don't go in NoReceipts

    const nrData  = nrSheet.getDataRange().getValues();

    // Find existing NRC row for this entry
    let existingRow = -1;
    for (let i = 1; i < nrData.length; i++) {
      if (nrData[i][1] === entryId) { existingRow = i + 1; break; }
    }

    if (!wasReceipted && nowReceipted) {
      // Receipt was added — mark NRC row as RECEIPTED (update description col as note)
      if (existingRow > -1) {
        nrSheet.getRange(existingRow, 4).setValue('[RECEIPTED] ' + (nrData[existingRow - 1][3] || ''));
      }
    } else if (wasReceipted && !nowReceipted) {
      // Receipt was removed — create a new NRC row
      saveNoReceiptRecord({
        entryId    : entryId,
        date       : newPayload.date,
        description: newPayload.description || '',
        amount     : parseFloat(newPayload.amount) || 0,
        requestedBy: newPayload.requestedBy || ''
      });
    } else if (!wasReceipted && !nowReceipted && existingRow > -1) {
      // Still no receipt but fields may have changed — update the row
      nrSheet.getRange(existingRow, 4, 1, 3).setValues([[
        newPayload.description || '',
        parseFloat(newPayload.amount) || 0,
        newPayload.requestedBy || ''
      ]]);
    }
  } catch(e) {
    console.error('syncNoReceiptOnUpdate error:', e);
  }
}

function deleteExpenseEntry(entryId) {
  try {
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet     = ss.getSheetByName(SHEETS.ENTRIES);
    const dataRange = sheet.getDataRange().getValues();
    const userEmail = getUserEmail();
    const now       = new Date().toISOString();

    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] !== entryId) continue;
      const date = normalizeDate(dataRange[i][1]);
      const row  = i + 1;

      sheet.getRange(row, 11).setValue('DELETED');
      sheet.getRange(row, 13).setValue(now);
      sheet.getRange(row, 14).setValue(`${userEmail} @ ${now}`);

      // Mark corresponding NRC row as DELETED
      markNoReceiptDeleted(entryId);

      recalculateDailySummary(date);
      return { success: true };
    }
    return { success: false, message: 'Entry not found' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getExpenseEntries(date) {
  try {
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet     = ss.getSheetByName(SHEETS.ENTRIES);
    const dataRange = sheet.getDataRange().getValues();
    const entries   = [];

    for (let i = 1; i < dataRange.length; i++) {
      const row     = dataRange[i];
      const rowDate = normalizeDate(row[1]);
      if (rowDate !== date || row[10] === 'DELETED') continue;

      entries.push({
        id         : row[0],
        date       : rowDate,
        type       : row[2],
        category   : row[3],
        description: row[4],
        amount     : row[5],
        hasReceipt : row[6] === 'YES',
        referenceNo: row[7],
        requestedBy: row[8],
        approvedBy : row[9],
        status     : row[10],
        createdAt  : row[11],
        updatedAt  : row[12]
      });
    }
    return { success: true, data: entries };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// ─────────────────────────────────────────────
// NO RECEIPTS — PettyCash_NoReceipts
// Columns: NR_ID | Entry_ID | Date | Description | Amount | Requested_By | Created_By | Created_At
// ─────────────────────────────────────────────

function saveNoReceiptRecord(data) {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const nrSheet = ss.getSheetByName(SHEETS.NO_RECEIPTS);
    const user    = getUserEmail();
    const now     = new Date().toISOString();
    const nrId    = generateId('NRC', data.date, nrSheet);

    nrSheet.appendRow([
      nrId,                    // NR_ID
      data.entryId  || '',     // Entry_ID (FK → PettyCash_Entries)
      data.date,               // Date
      data.description || '',  // Description
      parseFloat(data.amount) || 0, // Amount
      data.requestedBy || '',  // Requested_By
      user,                    // Created_By
      now                      // Created_At
    ]);

    return { success: true, id: nrId };
  } catch(e) {
    console.error('saveNoReceiptRecord error:', e);
    return { success: false, message: e.toString() };
  }
}

function markNoReceiptDeleted(entryId) {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const nrSheet = ss.getSheetByName(SHEETS.NO_RECEIPTS);
    if (!nrSheet) return;

    const data = nrSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === entryId) {
        // Prefix description with [DELETED] as a soft marker
        const existing = data[i][3] || '';
        if (!existing.startsWith('[DELETED]')) {
          nrSheet.getRange(i + 1, 4).setValue('[DELETED] ' + existing);
        }
        break;
      }
    }
  } catch(e) {
    console.error('markNoReceiptDeleted error:', e);
  }
}

function getNoReceiptsByDate(date) {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const nrSheet = ss.getSheetByName(SHEETS.NO_RECEIPTS);
    if (!nrSheet) return { success: true, data: [] };

    const data    = nrSheet.getDataRange().getValues();
    const records = [];

    for (let i = 1; i < data.length; i++) {
      const row     = data[i];
      const rowDate = normalizeDate(row[2]);
      if (rowDate !== date) continue;

      // Exclude soft-deleted rows
      const desc = String(row[3] || '');
      if (desc.startsWith('[DELETED]')) continue;

      records.push({
        nrId       : row[0],
        entryId    : row[1],
        date       : rowDate,
        description: desc,
        amount     : row[4],
        requestedBy: row[5],
        createdBy  : row[6],
        createdAt  : row[7]
      });
    }
    return { success: true, data: records };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// ─────────────────────────────────────────────
// BIR RECEIPTS — Purchases Journal
// ─────────────────────────────────────────────
function saveReceiptRecord(data) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.RECEIPTS);
    const vat   = computeVAT(data.grossAmount);
    const id    = generateId('RCP', data.date, sheet);
    const user  = getUserEmail();
    const now   = new Date().toISOString();

    sheet.appendRow([
      id,                    // Receipt_ID
      data.entryId  || '',   // Entry_ID (FK)
      data.date,             // Date
      data.supplierName,     // Supplier_Name
      data.address    || '', // Address
      data.tin        || '', // TIN
      data.receiptNo  || '', // Receipt_No / OR No.
      vat.grossAmount,       // Gross_Amount
      vat.vatableSales,      // Vatable_Sales (Less: VAT)
      vat.vatAmount,         // VAT_Amount (12%)
      user,                  // Created_By
      now                    // Created_At
    ]);

    return { success: true, id, vat };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getReceiptByEntryId(entryId) {
  try {
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet     = ss.getSheetByName(SHEETS.RECEIPTS);
    const dataRange = sheet.getDataRange().getValues();

    for (let i = 1; i < dataRange.length; i++) {
      const row = dataRange[i];
      if (row[1] !== entryId) continue;
      return {
        success: true,
        data: {
          receiptId   : row[0],
          entryId     : row[1],
          date        : normalizeDate(row[2]),
          supplierName: row[3],
          address     : row[4],
          tin         : row[5],
          receiptNo   : row[6],
          grossAmount : row[7],
          vatableSales: row[8],
          vatAmount   : row[9],
          createdBy   : row[10],
          createdAt   : row[11]
        }
      };
    }
    return { success: true, data: null };
  } catch(e) {
    return { success: false, message: e.toString(), data: null };
  }
}

function getReceiptsByDate(date) {
  try {
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet     = ss.getSheetByName(SHEETS.RECEIPTS);
    const dataRange = sheet.getDataRange().getValues();
    const records   = [];

    for (let i = 1; i < dataRange.length; i++) {
      const row     = dataRange[i];
      const rowDate = normalizeDate(row[2]);
      if (rowDate !== date) continue;

      records.push({
        receiptId   : row[0],
        entryId     : row[1],
        date        : rowDate,
        supplierName: row[3],
        address     : row[4],
        tin         : row[5],
        receiptNo   : row[6],
        grossAmount : row[7],
        vatableSales: row[8],
        vatAmount   : row[9]
      });
    }
    return { success: true, data: records };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// ─────────────────────────────────────────────
// DENOMINATION RECORDS
// ─────────────────────────────────────────────
function saveDenominationRecord(data) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.DENOMINATIONS);

    const denoms = parseDenomBreakdown(data);
    if (!denoms) return { success: false, message: 'Invalid denomination data — could not parse breakdown.' };

    const total     = calculateDenomTotal(denoms);
    const now       = new Date().toISOString();
    const dataRange = sheet.getDataRange().getValues();

    // Determine prefix based on type
    const prefix = data.type === 'START' ? 'DEN-OC'
                 : data.type === 'END'   ? 'DEN-CC'
                 : 'DEN';

    // Check if a record already exists for this date + type
    let targetRow = -1;
    for (let i = 1; i < dataRange.length; i++) {
      if (normalizeDate(dataRange[i][1]) === data.date && dataRange[i][2] === data.type) {
        targetRow = i + 1;
        break;
      }
    }

    const rowValues = [
      denoms['1000'] ||0, denoms['500']||0, denoms['200']||0,
      denoms['100']  ||0, denoms['50'] ||0, denoms['20'] ||0,
      denoms['10']   ||0, denoms['5']  ||0, denoms['1']  ||0,
      denoms['0.25'] ||0,
      total, data.notes || '', now
    ];

    let recordId;
    if (targetRow !== -1) {
      recordId = dataRange[targetRow - 1][0];
      sheet.getRange(targetRow, 4, 1, 13).setValues([rowValues]);
    } else {
      recordId = generateId(prefix, data.date, sheet);
      sheet.appendRow([
        recordId, data.date, data.type,
        ...rowValues
      ]);
    }

    recalculateDailySummary(data.date);
    return { success: true, id: recordId, total };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getDenominationRecords(date) {
  try {
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet     = ss.getSheetByName(SHEETS.DENOMINATIONS);
    const dataRange = sheet.getDataRange().getValues();
    const records   = [];

    for (let i = 1; i < dataRange.length; i++) {
      const row     = dataRange[i];
      const rowDate = normalizeDate(row[1]);
      if (rowDate !== date) continue;

      records.push({
        id   : row[0], date: rowDate, type: row[2],
        denominations: {
          '1000':row[3], '500':row[4],  '200':row[5],
          '100' :row[6], '50' :row[7],  '20' :row[8],
          '10'  :row[9], '5'  :row[10], '1'  :row[11],
          '0.25':row[12]
        },
        total: row[13], notes: row[14], timestamp: row[15]
      });
    }
    return { success: true, data: records };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// ─────────────────────────────────────────────
// DAILY SUMMARY
// ─────────────────────────────────────────────
function getDailySummary(date) {
  try {
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet     = ss.getSheetByName(SHEETS.SUMMARY);
    const dataRange = sheet.getDataRange().getValues();

    for (let i = 1; i < dataRange.length; i++) {
      const row     = dataRange[i];
      const rowDate = normalizeDate(row[1]);
      if (rowDate !== date) continue;

      return {
        success: true,
        data: {
          id                 : row[0],
          date               : rowDate,
          openingCash        : row[2],
          cashAdvance        : row[3],
          totalWithReceipt   : row[4],
          totalWithoutReceipt: row[5],
          totalExpenses      : row[6],
          totalCashOver      : row[7],
          totalReplenishment : row[8],
          closingCash        : row[9],
          variance           : row[10],
          status             : row[11],
          closedBy           : row[12],
          updatedAt          : row[13]
        }
      };
    }
    return { success: true, data: null };
  } catch(e) {
    return { success: false, message: e.toString(), data: null };
  }
}

// ─────────────────────────────────────────────
// REPORT DATA
// ─────────────────────────────────────────────
function generateReportData(params) {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const entries = [];
    const entryData = ss.getSheetByName(SHEETS.ENTRIES).getDataRange().getValues();

    for (let i = 1; i < entryData.length; i++) {
      const row   = entryData[i];
      if (row.length < 11) continue;
      const rDate = normalizeDate(row[1]);
      if (rDate < params.from || rDate > params.to || row[10] === 'DELETED') continue;

      entries.push({
        id:row[0], date:rDate, type:row[2], category:row[3],
        description:row[4], amount:row[5], hasReceipt:row[6]==='YES',
        referenceNo:row[7], requestedBy:row[8], status:row[10]
      });
    }

    const summaries = [];
    const sumData   = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    for (let j = 1; j < sumData.length; j++) {
      const sRow  = sumData[j];
      const sDate = normalizeDate(sRow[1]);
      if (sDate < params.from || sDate > params.to) continue;
      summaries.push({
        date:sDate, opening:sRow[2], cashAdvance:sRow[3],
        expenses:sRow[6], closing:sRow[9], variance:sRow[10], status:sRow[11]
      });
    }

    return { success: true, data: { entries, summaries } };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function findUnclosedPastDate(beforeDate) {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const data    = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    const unclosed = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i].length < 12) continue;
      const rDate = normalizeDate(data[i][1]);
      if (rDate < beforeDate && data[i][11] === 'OPEN') unclosed.push(rDate);
    }

    if (!unclosed.length) return { success: true, date: null };
    unclosed.sort();
    return { success: true, date: unclosed[0] };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function updateDenominationRecord(data) {
  try {
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet     = ss.getSheetByName(SHEETS.DENOMINATIONS);
    const dataRange = sheet.getDataRange().getValues();

    let targetRow = -1;
    for (let i = 1; i < dataRange.length; i++) {
      const rowDate = normalizeDate(dataRange[i][1]);
      const rowType = dataRange[i][2];
      if (rowDate === data.date && rowType === data.type) {
        targetRow = i + 1;
        break;
      }
    }

    const denoms = parseDenomBreakdown(data);
    if (!denoms) return { success: false, message: 'Invalid denomination data — could not parse breakdown.' };

    const total = calculateDenomTotal(denoms);
    const now   = new Date().toISOString();

    if (targetRow === -1) {
      return saveDenominationRecord(data);
    }

    sheet.getRange(targetRow, 4, 1, 11).setValues([[
      denoms['1000'] ||0, denoms['500']||0, denoms['200']||0,
      denoms['100']  ||0, denoms['50'] ||0, denoms['20'] ||0,
      denoms['10']   ||0, denoms['5']  ||0, denoms['1']  ||0,
      denoms['0.25'] ||0,
      total
    ]]);
    sheet.getRange(targetRow, 15).setValue(data.notes || '');
    sheet.getRange(targetRow, 16).setValue(now);

    recalculateDailySummary(data.date);
    return { success: true, id: dataRange[targetRow - 1][0], total };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function parseDenomBreakdown(data) {
  try {
    if (data.denominations && typeof data.denominations === 'object') {
      return data.denominations;
    }
    if (data.breakdown && typeof data.breakdown === 'object') {
      return data.breakdown;
    }
    if (typeof data.breakdown === 'string' && data.breakdown.trim() !== '') {
      const parsed = JSON.parse(data.breakdown);
      if (typeof parsed === 'object' && parsed !== null) return parsed;
    }
    return null;
  } catch(e) {
    console.error('parseDenomBreakdown failed:', e);
    return null;
  }
}