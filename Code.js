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
  ACCESS       : 'PettyCash_Access',
  AUDIT_LOG    : 'PettyCash_AuditLog',
  FILING       : 'PettyCash_FilingChecklist'
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

  initializeSheets();

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
// AUDIT LOG
// ─────────────────────────────────────────────
function writeAuditLog(action, details, referenceId, date) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
    if (!sheet) return;

    const email    = getUserEmail();
    const roleInfo = getUserRole();
    const role     = roleInfo.success ? roleInfo.role : 'Unknown';
    const now      = new Date().toISOString();
    const logDate  = date || normalizeDate(new Date());
    const id       = generateId('LOG', logDate, sheet);

    sheet.appendRow([
      id,
      now,
      logDate,
      action,
      email,
      role,
      details,
      referenceId || ''
    ]);
  } catch(e) {
    // Never let audit log failure break the main action
    console.error('writeAuditLog error:', e);
  }
}

function getAuditLog(params) {
  // params: { from, to, action (optional) }
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
    if (!sheet) return { success: true, data: [] };

    const rows = sheet.getDataRange().getValues();
    const logs = [];

    for (let i = 1; i < rows.length; i++) {
      const row    = rows[i];
      const logDate = normalizeDate(row[2]);

      if (params.from && logDate < params.from) continue;
      if (params.to   && logDate > params.to)   continue;
      if (params.action && row[3] !== params.action) continue;

      logs.push({
        id         : row[0],
        timestamp  : row[1],
        date       : logDate,
        action     : row[3],
        actorEmail : row[4],
        actorRole  : row[5],
        details    : row[6],
        referenceId: row[7]
      });
    }

    // Most recent first
    logs.sort((a, b) => b.timestamp > a.timestamp ? 1 : -1);
    return { success: true, data: logs };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
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

  if (!ss.getSheetByName(SHEETS.AUDIT_LOG)) {
    const s = ss.insertSheet(SHEETS.AUDIT_LOG);
    s.appendRow([
      'Log_ID', 'Timestamp', 'Date', 'Action',
      'Actor_Email', 'Actor_Role', 'Details', 'Reference_ID'
    ]);
    s.setFrozenRows(1);
    formatHeaderRow(s);
    s.setColumnWidth(1, 160);
    s.setColumnWidth(2, 180);
    s.setColumnWidth(3, 100);
    s.setColumnWidth(4, 160);
    s.setColumnWidth(5, 220);
    s.setColumnWidth(6, 100);
    s.setColumnWidth(7, 400);
    s.setColumnWidth(8, 160);
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
      const row    = entryData[i];
      const rDate  = normalizeDate(row[1]);
      const type   = row[2];
      const status = row[10];

      if (status === 'DELETED') continue;

      const amt = parseFloat(row[5]) || 0;

      if (rDate === date) {
        if (type === 'CASH_ADVANCE')    cashAdvance        += amt;
        else if (type === 'CASH_OVER')  totalCashOver      += amt;
        else if (type === 'REPLENISHMENT') totalReplenishment += amt;
        else {
          totalExp += amt;
          if (row[6] === 'YES') totalReceipt   += amt;
          else                  totalNoReceipt += amt;
        }
        continue;
      }

      // Carried-forward unliquidated advances from BEFORE this date
      if (type === 'CASH_ADVANCE' && rDate < date && status !== 'LIQUIDATED') {
        cashAdvance += amt;
      }
    }

    const denomData = ss.getSheetByName(SHEETS.DENOMINATIONS).getDataRange().getValues();
    let openingCash = 0, closingCash = 0, hasClosing = false;

    for (let i = 1; i < denomData.length; i++) {
      if (normalizeDate(denomData[i][1]) !== date) continue;
      const type  = denomData[i][2];
      const total = parseFloat(denomData[i][13]) || 0;
      if (type === 'START') openingCash = total;
      if (type === 'END')   { closingCash = total; hasClosing = true; }
    }

    // ── FIX: declare sumSheet and sumData BEFORE reading existingStatus ──
    const sumSheet = ss.getSheetByName(SHEETS.SUMMARY);
    const sumData  = sumSheet.getDataRange().getValues();

    let existingStatus = '', targetRow = -1, existingClosedBy = '';
    for (let i = 1; i < sumData.length; i++) {
      if (normalizeDate(sumData[i][1]) === date) {
        existingStatus   = sumData[i][11] || '';
        existingClosedBy = sumData[i][12] || '';
        targetRow = i + 1;
        break;
      }
    }

    const expected = (openingCash + totalReplenishment) - (totalExp + cashAdvance);
    const variance = closingCash - expected;
    const status   = hasClosing
      ? (existingStatus === 'CLOSED' ? 'CLOSED' : 'PENDING_AUDIT')
      : 'OPEN';

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

    // Audit log for cash advance
    if (data.type === 'CASH_ADVANCE') {
      writeAuditLog(
        'CASH_ADVANCE_ISSUED',
        `Cash advance of ₱${parseFloat(data.amount).toFixed(2)} issued to ${data.requestedBy || '—'}. Desc: ${data.description || '—'}`,
        id,
        data.date
      );
    }

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

      // If cashier is correcting a flagged entry, reset its status to ACTIVE
      if (dataRange[i][10] === 'FLAGGED') {
        sheet.getRange(row, 11).setValue('ACTIVE');
        sheet.getRange(row, 15).setValue('');
      }

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

      markNoReceiptDeleted(entryId);
      recalculateDailySummary(date);

      writeAuditLog(
        'ENTRY_DELETED',
        `Entry ${entryId} deleted by ${userEmail}`,
        entryId,
        date
      );
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
      const type    = row[2];
      const status  = row[10];

      if (status === 'DELETED') continue;

      // Always include entries recorded ON this date
      if (rowDate === date) {
        entries.push({
          id              : row[0],
          date            : rowDate,
          type            : type,
          category        : row[3],
          description     : row[4],
          amount          : row[5],
          hasReceipt      : row[6] === 'YES',
          referenceNo     : row[7],
          requestedBy     : row[8],
          approvedBy      : row[9],
          status          : status,
          createdAt       : row[11],
          updatedAt       : row[12],
          carriedForward  : false,
          originalDate    : rowDate
        });
        continue;
      }

      // Carry forward unliquidated cash advances from BEFORE this date
      if (
        type === 'CASH_ADVANCE' &&
        rowDate < date &&
        status !== 'LIQUIDATED' &&
        status !== 'DELETED'
      ) {
        const issuedDate  = rowDate;
        const msPerDay    = 1000 * 60 * 60 * 24;
        const daysOut     = Math.floor(
          (new Date(date) - new Date(issuedDate)) / msPerDay
        );

        entries.push({
        id              : row[0],
        date            : rowDate,
        type            : type,
        category        : row[3],
        description     : row[4],
        amount          : row[5],
        hasReceipt      : row[6] === 'YES',
        referenceNo     : row[7],
        requestedBy     : row[8],
        approvedBy      : row[9],
        status          : status,
        createdAt       : row[11],
        updatedAt       : row[12],
        auditNote       : String(row[14] || '').replace('[AUDITOR FLAG] ', ''),
        carriedForward  : false,
        originalDate    : rowDate
      });
      }
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

    // ── Update Has_Receipt on the linked entry ──
    if (data.entryId) {
      const entrySheet = ss.getSheetByName(SHEETS.ENTRIES);
      const entryRows  = entrySheet.getDataRange().getValues();
      for (let i = 1; i < entryRows.length; i++) {
        if (entryRows[i][0] === data.entryId) {
          entrySheet.getRange(i + 1, 7).setValue('YES');
          entrySheet.getRange(i + 1, 8).setValue(data.receiptNo || '');
          entrySheet.getRange(i + 1, 13).setValue(now);
          markNoReceiptDeleted(data.entryId);
          recalculateDailySummary(data.date);
          break;
        }
      }
    }

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

    const actionLabel = data.type === 'START' ? 'OPENING_SAVED' : 'CLOSING_SAVED';
    writeAuditLog(
      actionLabel,
      `${data.type === 'START' ? 'Opening' : 'Closing'} denomination saved. Total: ₱${total.toFixed(2)}. Notes: ${data.notes || '—'}`,
      recordId,
      data.date
    );

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

    // No Summary row yet — fall back to Denominations sheet directly
    // This happens when only an opening has been saved but no entries yet
    const denomSheet = ss.getSheetByName(SHEETS.DENOMINATIONS);
    const denomData  = denomSheet.getDataRange().getValues();
    let openingCash  = 0;
    let hasOpening   = false;

    for (let i = 1; i < denomData.length; i++) {
      if (normalizeDate(denomData[i][1]) !== date) continue;
      if (denomData[i][2] === 'START') {
        openingCash = parseFloat(denomData[i][13]) || 0;
        hasOpening  = true;
      }
    }

    if (hasOpening) {
      return {
        success: true,
        data: {
          id                 : null,
          date               : date,
          openingCash        : openingCash,
          cashAdvance        : 0,
          totalWithReceipt   : 0,
          totalWithoutReceipt: 0,
          totalExpenses      : 0,
          totalCashOver      : 0,
          totalReplenishment : 0,
          closingCash        : 0,
          variance           : 0,
          status             : 'OPEN',
          closedBy           : '',
          updatedAt          : ''
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
    const summaryDates = new Set();
    const sumData   = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    for (let j = 1; j < sumData.length; j++) {
      const sRow  = sumData[j];
      const sDate = normalizeDate(sRow[1]);
      if (sDate < params.from || sDate > params.to) continue;
      summaryDates.add(sDate);
      summaries.push({
        date:sDate, opening:sRow[2], cashAdvance:sRow[3],
        totalWithReceipt:sRow[4], totalWithoutReceipt:sRow[5],
        expenses:sRow[6], cashOver:sRow[7], replenishment:sRow[8],
        closing:sRow[9], variance:sRow[10], status:sRow[11]
      });
    }

    // ── Fallback: include OPEN days that have a denomination START record
    // but no SUMMARY row yet (opening saved, no entries recorded yet)
    const denomSheet = ss.getSheetByName(SHEETS.DENOMINATIONS);
    if (denomSheet) {
      const denomData = denomSheet.getDataRange().getValues();
      for (let k = 1; k < denomData.length; k++) {
        const dRow  = denomData[k];
        const dDate = normalizeDate(dRow[1]);
        if (dDate < params.from || dDate > params.to) continue;
        if (dRow[2] !== 'START') continue;           // only opening records
        if (summaryDates.has(dDate)) continue;       // already in summary
        summaryDates.add(dDate);
        summaries.push({
          date               : dDate,
          opening            : parseFloat(dRow[13]) || 0,
          cashAdvance        : 0,
          totalWithReceipt   : 0,
          totalWithoutReceipt: 0,
          expenses           : 0,
          cashOver           : 0,
          replenishment      : 0,
          closing            : 0,
          variance           : 0,
          status             : 'OPEN'
        });
      }
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
      const s = data[i][11];
      if (rDate < beforeDate && (s === 'OPEN' || s === 'PENDING_AUDIT' || s === 'FLAGGED')) unclosed.push(rDate);
    }

    if (!unclosed.length) return { success: true, date: null };
    unclosed.sort();
    return { success: true, date: unclosed[0] };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
// AUDITOR APPROVAL
// ─────────────────────────────────────────────

function auditApproveDay(data) {
  // data: { date, notes, carryForward: bool }
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SUMMARY);
    const rows  = sheet.getDataRange().getValues();
    const email = getUserEmail();
    const now   = new Date().toISOString();

    let found = false;

    for (let i = 1; i < rows.length; i++) {
      if (normalizeDate(rows[i][1]) !== data.date) continue;
      const row = i + 1;
      sheet.getRange(row, 12).setValue('CLOSED');
      sheet.getRange(row, 13).setValue(email);
      sheet.getRange(row, 14).setValue(now);
      found = true;

      writeAuditLog(
        'DAY_APPROVED',
        `Auditor approved and closed the day. Notes: ${data.notes || '—'}`,
        '',
        data.date
      );
      break;
    }

    if (!found) return { success: false, message: 'No summary record found for ' + data.date };

    // ── Always carry forward the auditor's END count as next day's START ──
    const nextDate   = getNextDate(data.date);
    const denomSheet = ss.getSheetByName(SHEETS.DENOMINATIONS);
    const denomRows  = denomSheet.getDataRange().getValues();

    // Find auditor's END count for this date
    let endRow = null;
    for (let j = 1; j < denomRows.length; j++) {
      if (normalizeDate(denomRows[j][1]) === data.date && denomRows[j][2] === 'END') {
        endRow = denomRows[j];
        break;
      }
    }

    // Fallback: if no END denom record exists but actualCash was passed, create one now
    if (!endRow && data.actualCash && parseFloat(data.actualCash) > 0) {
      const fallbackId = 'DEN-END-' + data.date.replace(/-/g,'') + '-AU';
      denomSheet.appendRow([
        fallbackId, data.date, 'END',
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        parseFloat(data.actualCash),
        'Auto-saved on audit approval',
        now
      ]);
      // Re-read to get the row we just appended
      const refreshed = denomSheet.getDataRange().getValues();
      for (let j = 1; j < refreshed.length; j++) {
        if (normalizeDate(refreshed[j][1]) === data.date && refreshed[j][2] === 'END') {
          endRow = refreshed[j];
          break;
        }
      }
    }

    if (endRow) {
      // Check if next day START already exists — don't overwrite
      let nextDayStartExists = false;
      for (let j = 1; j < denomRows.length; j++) {
        if (normalizeDate(denomRows[j][1]) === nextDate && denomRows[j][2] === 'START') {
          nextDayStartExists = true;
          break;
        }
      }

      if (!nextDayStartExists) {
        const endTotal = parseFloat(endRow[13]) || 0;
        const newId    = 'DEN-OC-' + nextDate.replace(/-/g,'') + '-CF';

        denomSheet.appendRow([
          newId,          // ID
          nextDate,       // Date
          'START',        // Type
          endRow[3],      // ₱1000
          endRow[4],      // ₱500
          endRow[5],      // ₱200
          endRow[6],      // ₱100
          endRow[7],      // ₱50
          endRow[8],      // ₱20
          endRow[9],      // ₱10
          endRow[10],     // ₱5
          endRow[11],     // ₱1
          endRow[12],     // ₱0.25
          endTotal,       // Total
          'Carried forward from ' + data.date + ' audit closing count',
          now             // Timestamp
        ]);

        recalculateDailySummary(nextDate);

        writeAuditLog(
          'OPENING_SAVED',
          `Opening cash auto-carried from ${data.date} audit count. Total: ₱${endTotal.toFixed(2)}`,
          newId,
          nextDate
        );
      }
    }

    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getNextDate(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  d.setDate(d.getDate() + 1);
  return d.toISOString().split('T')[0];
}

// Helper: get next calendar date string
function getNextDate(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  d.setDate(d.getDate() + 1);
  return d.toISOString().split('T')[0];
}

function getPreviousDayClosing(date) {
  // Returns the closing total of the most recent day before `date` that is CLOSED
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rows  = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    let best    = null;

    for (let i = 1; i < rows.length; i++) {
      const rDate  = normalizeDate(rows[i][1]);
      const status = rows[i][11];
      if (rDate < date && status === 'CLOSED') {
        if (!best || rDate > best.date) {
          best = { date: rDate, closingCash: parseFloat(rows[i][9]) || 0 };
        }
      }
    }
    return { success: true, data: best };
  } catch(e) {
    return { success: false, message: e.toString(), data: null };
  }
}

function getDayStatus(date) {
  // Returns full status info for a given date including outstanding advances
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sumRows = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    let summaryStatus = 'NO_RECORD';
    let openingCash   = 0;
    let closingCash   = 0;
    let replenishment = 0;

    for (let i = 1; i < sumRows.length; i++) {
      if (normalizeDate(sumRows[i][1]) !== date) continue;
      summaryStatus = sumRows[i][11] || 'OPEN';
      openingCash   = parseFloat(sumRows[i][2])  || 0;
      closingCash   = parseFloat(sumRows[i][9])  || 0;
      replenishment = parseFloat(sumRows[i][8])  || 0;
      break;
    }

    // Count outstanding cash advances for this date and earlier
    const entryRows  = ss.getSheetByName(SHEETS.ENTRIES).getDataRange().getValues();
    let outstandingAdvances = 0;
    let outstandingTotal    = 0;

    for (let i = 1; i < entryRows.length; i++) {
      const row    = entryRows[i];
      const rDate  = normalizeDate(row[1]);
      const type   = row[2];
      const status = row[10];
      if (type !== 'CASH_ADVANCE') continue;
      if (status === 'DELETED' || status === 'LIQUIDATED') continue;
      if (rDate > date) continue;
      outstandingAdvances++;
      outstandingTotal += parseFloat(row[5]) || 0;
    }

    // Get flag note from audit log if FLAGGED
  let flagNote = '';
  if (summaryStatus === 'FLAGGED') {
    try {
      const logSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
      if (logSheet) {
        const logs = logSheet.getDataRange().getValues();
        for (let i = logs.length - 1; i >= 1; i--) {
          if (normalizeDate(logs[i][2]) === date && logs[i][3] === 'DAY_FLAGGED') {
            flagNote = logs[i][6] || '';
            break;
          }
        }
      }
    } catch(e) {}
  }

  return {
    success            : true,
    status             : summaryStatus,
    openingCash        : openingCash,
    closingCash        : closingCash,
    replenishment      : replenishment,
    outstandingAdvances: outstandingAdvances,
    outstandingTotal   : outstandingTotal,
    flagNote           : flagNote
  };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
// CASH ADVANCE TRACKING
// ─────────────────────────────────────────────

function getOutstandingCashAdvances() {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const entries = ss.getSheetByName(SHEETS.ENTRIES).getDataRange().getValues();
    const today   = normalizeDate(new Date());
    const advances = [];

    for (let i = 1; i < entries.length; i++) {
      const row    = entries[i];
      const type   = row[2];
      const status = row[10];

      if (type !== 'CASH_ADVANCE') continue;
      if (status === 'DELETED') continue;
      // ACTIVE = outstanding, LIQUIDATION_PENDING = submitted by cashier, LIQUIDATED = closed
      if (status === 'LIQUIDATED') continue;

      const issuedDate = normalizeDate(row[1]);
      const msPerDay   = 1000 * 60 * 60 * 24;
      const daysOut    = Math.floor(
        (new Date(today) - new Date(issuedDate)) / msPerDay
      );

      advances.push({
        id             : row[0],
        date           : issuedDate,
        description    : row[4],
        amount         : row[5],
        requestedBy    : row[8],
        status         : status,
        daysOutstanding: daysOut,
        liquidationNote: row[14] || ''
      });
    }

    // Sort by days outstanding descending (oldest first)
    advances.sort((a, b) => b.daysOutstanding - a.daysOutstanding);
    return { success: true, data: advances };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

function getAllCashAdvances() {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const entries = ss.getSheetByName(SHEETS.ENTRIES).getDataRange().getValues();
    const today   = normalizeDate(new Date());
    const advances = [];

    for (let i = 1; i < entries.length; i++) {
      const row = entries[i];
      if (row[2] !== 'CASH_ADVANCE') continue;
      if (row[10] === 'DELETED') continue;

      const issuedDate = normalizeDate(row[1]);
      const msPerDay   = 1000 * 60 * 60 * 24;
      const daysOut    = Math.floor(
        (new Date(today) - new Date(issuedDate)) / msPerDay
      );

      advances.push({
        id             : row[0],
        date           : issuedDate,
        description    : row[4],
        amount         : row[5],
        requestedBy    : row[8],
        status         : row[10],
        daysOutstanding: daysOut,
        liquidationNote: row[14] || ''
      });
    }

    advances.sort((a, b) => b.daysOutstanding - a.daysOutstanding);
    return { success: true, data: advances };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

function submitLiquidation(data) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const rows  = sheet.getDataRange().getValues();
    const now   = new Date().toISOString();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== data.id) continue;
      const row         = i + 1;
      const advanceDate = normalizeDate(rows[i][1]);
      const requestedBy = rows[i][8] || '';

      // 1. Mark the advance as pending
      sheet.getRange(row, 11).setValue('LIQUIDATION_PENDING');
      sheet.getRange(row, 13).setValue(now);
      sheet.getRange(row, 15).setValue('[LIQUIDATION SUBMITTED] ' + (data.note || ''));

      // 2. Save each line item entry + receipt if present
      if (data.entries && data.entries.length) {
        data.entries.forEach(entry => {
          // Save to PettyCash_Entries (also auto-saves to PettyCash_NoReceipts if no receipt)
          const entryResult = saveExpenseEntry({
            date       : advanceDate,
            type       : 'EXPENSE',
            category   : entry.category   || 'Miscellaneous',
            description: entry.desc       || '',
            amount     : parseFloat(entry.amount) || 0,
            hasReceipt : !!entry.hasReceipt,
            referenceNo: data.id,   // links back to the cash advance
            requestedBy: requestedBy,
            approvedBy : ''
          });

          // Save to PettyCash_Receipts if receipt was attached
          if (entry.hasReceipt && entry.receipt && entryResult.id) {
            saveReceiptRecord({
              ...entry.receipt,
              entryId: entryResult.id,
              date   : advanceDate
            });
          }
        });
      }

      writeAuditLog(
        'LIQUIDATION_SUBMITTED',
        `Liquidation submitted for advance ${data.id}. ${(data.entries || []).length} entries. Notes: ${data.note || '—'}`,
        data.id,
        advanceDate
      );

      return { success: true };
    }
    return { success: false, message: 'Advance not found' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function approveLiquidation(data) {
  // Called by Auditor — marks advance as LIQUIDATED
  // data: { id, note }
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const rows  = sheet.getDataRange().getValues();
    const email = getUserEmail();
    const now   = new Date().toISOString();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== data.id) continue;
      const row = i + 1;
      sheet.getRange(row, 11).setValue('LIQUIDATED');
      sheet.getRange(row, 13).setValue(now);
      sheet.getRange(row, 15).setValue(
        (rows[i][14] || '') + ' | [APPROVED BY ' + email + '] ' + (data.note || '')
      );
      recalculateDailySummary(normalizeDate(rows[i][1]));

      writeAuditLog(
        'LIQUIDATION_APPROVED',
        `Auditor approved liquidation for advance ${data.id}. Notes: ${data.note || '—'}`,
        data.id,
        normalizeDate(rows[i][1])
      );
      return { success: true };
    }
    return { success: false, message: 'Advance not found' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function rejectLiquidation(data) {
  // Called by Auditor — sends back to ACTIVE/outstanding
  // data: { id, note }
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const rows  = sheet.getDataRange().getValues();
    const email = getUserEmail();
    const now   = new Date().toISOString();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== data.id) continue;
      const row = i + 1;
      sheet.getRange(row, 11).setValue('ACTIVE');
      sheet.getRange(row, 13).setValue(now);
      sheet.getRange(row, 15).setValue(
        (rows[i][14] || '') + ' | [REJECTED BY ' + email + '] ' + (data.note || '')
      );

      writeAuditLog(
        'LIQUIDATION_REJECTED',
        `Auditor rejected liquidation for advance ${data.id}. Reason: ${data.note || '—'}`,
        data.id,
        normalizeDate(rows[i][1])
      );
      return { success: true };
    }
    return { success: false, message: 'Advance not found' };
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

function getUserRole() {
  try {
    const email = getUserEmail();
    if (!email || email === 'unknown') return { success: false, role: null, name: null };

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ACCESS);
    if (!sheet) return { success: false, role: null, name: null };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowEmail  = String(data[i][0] || '').trim().toLowerCase();
      const rowStatus = String(data[i][3] || '').trim().toUpperCase();
      if (rowEmail === email.toLowerCase() && rowStatus === 'ACTIVE') {
        return {
          success: true,
          email  : email,
          name   : String(data[i][1] || ''),
          role   : String(data[i][2] || 'Cashier').trim()  // Admin, Auditor, Cashier
        };
      }
    }
    return { success: false, role: null, name: null };
  } catch(e) {
    return { success: false, role: null, name: null, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
// FILING CHECKLIST
// Columns: Filing_ID | Date | Filed_Receipts | Filed_No_Receipts | Filed_Report | Notes | Submitted_By | Submitted_At
// ─────────────────────────────────────────────

function saveFilingChecklist(data) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet   = ss.getSheetByName(SHEETS.FILING);

    // Auto-create sheet if missing
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.FILING);
      sheet.appendRow([
        'Filing_ID','Date','Filed_Receipts','Filed_No_Receipts',
        'Filed_Report','Notes','Submitted_By','Submitted_At'
      ]);
      sheet.setFrozenRows(1);
      formatHeaderRow(sheet);
    }

    const user = getUserEmail();
    const now  = new Date().toISOString();
    const rows = sheet.getDataRange().getValues();

    // Check if a record already exists for this date — update if so
    for (let i = 1; i < rows.length; i++) {
      if (normalizeDate(rows[i][1]) === data.date) {
        sheet.getRange(i + 1, 3, 1, 6).setValues([[
          data.filedReceipts    ? 'YES' : 'NO',
          data.filedNoReceipts  ? 'YES' : 'NO',
          data.filedReport      ? 'YES' : 'NO',
          data.notes || '',
          user,
          now
        ]]);
        writeAuditLog('FILING_UPDATED', `Filing checklist updated for ${data.date}`, '', data.date);
        return { success: true, updated: true };
      }
    }

    // New record
    const id = generateId('FIL', data.date, sheet);
    sheet.appendRow([
      id,
      data.date,
      data.filedReceipts   ? 'YES' : 'NO',
      data.filedNoReceipts ? 'YES' : 'NO',
      data.filedReport     ? 'YES' : 'NO',
      data.notes || '',
      user,
      now
    ]);

    writeAuditLog('FILING_SUBMITTED', `Filing checklist submitted for ${data.date}`, id, data.date);
    return { success: true, updated: false };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getFilingChecklists(params) {
  // params: { from, to }
  // Now cross-references PettyCash_Summary CLOSED days so never-filed days appear too
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── 1. Collect all CLOSED days in range from Summary ──
    const sumSheet = ss.getSheetByName(SHEETS.SUMMARY);
    const closedDates = {};
    if (sumSheet) {
      const sumRows = sumSheet.getDataRange().getValues();
      for (let i = 1; i < sumRows.length; i++) {
        const rDate  = normalizeDate(sumRows[i][1]);
        const status = sumRows[i][11];
        if (params?.from && rDate < params.from) continue;
        if (params?.to   && rDate > params.to)   continue;
        if (status === 'CLOSED') closedDates[rDate] = true;
      }
    }

    // ── 2. Collect actual filing records in range ──
    const filingSheet = ss.getSheetByName(SHEETS.FILING);
    const filingMap   = {};
    if (filingSheet) {
      const rows = filingSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const rowDate = normalizeDate(rows[i][1]);
        if (params?.from && rowDate < params.from) continue;
        if (params?.to   && rowDate > params.to)   continue;
        filingMap[rowDate] = {
          id              : rows[i][0],
          date            : rowDate,
          filedReceipts   : rows[i][2] === 'YES',
          filedNoReceipts : rows[i][3] === 'YES',
          filedReport     : rows[i][4] === 'YES',
          notes           : rows[i][5] || '',
          submittedBy     : rows[i][6] || '',
          submittedAt     : rows[i][7] || '',
          fullyFiled      : rows[i][2] === 'YES' && rows[i][3] === 'YES' && rows[i][4] === 'YES',
          neverFiled      : false
        };
      }
    }

    // ── 3. Merge: every CLOSED day must appear, even if never filed ──
    const records = [];

    // Add all filing records in range (whether CLOSED or not)
    Object.values(filingMap).forEach(r => records.push(r));

    // Add CLOSED days that have no filing record at all
    Object.keys(closedDates).forEach(date => {
      if (!filingMap[date]) {
        records.push({
          id              : null,
          date            : date,
          filedReceipts   : false,
          filedNoReceipts : false,
          filedReport     : false,
          notes           : '',
          submittedBy     : '',
          submittedAt     : '',
          fullyFiled      : false,
          neverFiled      : true
        });
      }
    });

    records.sort((a, b) => b.date > a.date ? 1 : -1);
    return { success: true, data: records };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// ─────────────────────────────────────────────
// NEW: Returns all CLOSED days with filing status for the date picker panel
// ─────────────────────────────────────────────
function getClosedDaysForFiling() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Get all CLOSED days from Summary
    const sumRows   = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    const closedMap = {};
    for (let i = 1; i < sumRows.length; i++) {
      if (sumRows[i][11] === 'CLOSED') closedMap[normalizeDate(sumRows[i][1])] = true;
    }

    // Get all filing records
    const filingSheet = ss.getSheetByName(SHEETS.FILING);
    const filingMap   = {};
    if (filingSheet) {
      const rows = filingSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const d = normalizeDate(rows[i][1]);
        filingMap[d] = {
          fullyFiled : rows[i][2] === 'YES' && rows[i][3] === 'YES' && rows[i][4] === 'YES',
          neverFiled : false
        };
      }
    }

    // Build result: one entry per CLOSED day
    const days = Object.keys(closedMap).map(date => {
      if (filingMap[date]) {
        return {
          date      : date,
          status    : filingMap[date].fullyFiled ? 'filed' : 'incomplete'
        };
      }
      return { date, status: 'never' };
    });

    days.sort((a, b) => b.date > a.date ? 1 : -1);
    return { success: true, data: days };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

function flagDay(data) {
  // data: { date, note, flaggedEntries: [{id, note}] }
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SUMMARY);
    const rows  = sheet.getDataRange().getValues();
    const email = getUserEmail();
    const now   = new Date().toISOString();

    for (let i = 1; i < rows.length; i++) {
      if (normalizeDate(rows[i][1]) !== data.date) continue;
      const row = i + 1;
      sheet.getRange(row, 12).setValue('FLAGGED');
      sheet.getRange(row, 13).setValue(email);
      sheet.getRange(row, 14).setValue(now);

      // Tag each flagged entry with auditor note in col 15 (notes/remarks)
      if (data.flaggedEntries && data.flaggedEntries.length) {
        const entrySheet = ss.getSheetByName(SHEETS.ENTRIES);
        const entryRows  = entrySheet.getDataRange().getValues();
        data.flaggedEntries.forEach(fe => {
          for (let j = 1; j < entryRows.length; j++) {
            if (entryRows[j][0] !== fe.id) continue;
            entrySheet.getRange(j + 1, 11).setValue('FLAGGED');
            entrySheet.getRange(j + 1, 15).setValue('[AUDITOR FLAG] ' + (fe.note || ''));
            break;
          }
        });
      }

      writeAuditLog(
        'DAY_FLAGGED',
        `Auditor flagged day. ${(data.flaggedEntries||[]).length} entries flagged. Note: ${data.note || '—'}`,
        '',
        data.date
      );
      return { success: true };
    }
    return { success: false, message: 'No summary record found for ' + data.date };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
// ADMIN METRICS DASHBOARD
// ─────────────────────────────────────────────
function getAdminMetrics() {
  try {
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sumSheet  = ss.getSheetByName(SHEETS.SUMMARY);
    const filSheet  = ss.getSheetByName(SHEETS.FILING);
    const entSheet  = ss.getSheetByName(SHEETS.ENTRIES);

    const today = normalizeDate(new Date());

    // ── 1. Read all SUMMARY rows ──
    const sumRows = sumSheet ? sumSheet.getDataRange().getValues() : [];
    const unclosedDays   = [];
    const discrepancyDays = [];
    const cashOverDays   = [];

    for (let i = 1; i < sumRows.length; i++) {
      const row    = sumRows[i];
      if (row.length < 12) continue;
      const rDate  = normalizeDate(row[1]);
      if (rDate > today) continue;         // skip future dates
      const status   = row[11] || 'OPEN';
      const variance = parseFloat(row[10]) || 0;
      const cashOver = parseFloat(row[7])  || 0;
      const opening  = parseFloat(row[2])  || 0;
      const expenses = parseFloat(row[6])  || 0;
      const closing  = parseFloat(row[9])  || 0;

      // Unclosed = anything not CLOSED
      if (status !== 'CLOSED') {
        unclosedDays.push({ date: rDate, status, opening, expenses, closing, variance });
      }

      // Discrepancy = |variance| > 0.01
      if (Math.abs(variance) > 0.01) {
        discrepancyDays.push({ date: rDate, status, variance, opening, expenses, closing });
      }

      // Cash Over = cashOver > 0
      if (cashOver > 0) {
        cashOverDays.push({ date: rDate, status, cashOver, opening, expenses, closing });
      }
    }

    // ── 2. Unfiled days (CLOSED days with incomplete filing) ──
    const closedDates = {};
    for (let i = 1; i < sumRows.length; i++) {
      const row = sumRows[i];
      if ((row[11] || '') === 'CLOSED') {
        const d = normalizeDate(row[1]);
        if (d <= today) closedDates[d] = { date: d, opening: parseFloat(row[2])||0, expenses: parseFloat(row[6])||0, closing: parseFloat(row[9])||0 };
      }
    }

    const filingMap = {};
    if (filSheet) {
      const filRows = filSheet.getDataRange().getValues();
      for (let i = 1; i < filRows.length; i++) {
        const d = normalizeDate(filRows[i][1]);
        filingMap[d] = {
          filedReceipts  : filRows[i][2] === 'YES',
          filedNoReceipts: filRows[i][3] === 'YES',
          filedReport    : filRows[i][4] === 'YES',
          fullyFiled     : filRows[i][2] === 'YES' && filRows[i][3] === 'YES' && filRows[i][4] === 'YES'
        };
      }
    }

    const unfiledDays = [];
    Object.keys(closedDates).forEach(d => {
      const fil = filingMap[d];
      if (!fil || !fil.fullyFiled) {
        unfiledDays.push({
          date           : d,
          opening        : closedDates[d].opening,
          expenses       : closedDates[d].expenses,
          closing        : closedDates[d].closing,
          filedReceipts  : fil ? fil.filedReceipts   : false,
          filedNoReceipts: fil ? fil.filedNoReceipts  : false,
          filedReport    : fil ? fil.filedReport      : false,
          neverFiled     : !fil
        });
      }
    });

    // ── 3. Pending cash advances (ACTIVE or LIQUIDATION_PENDING) ──
    const pendingAdvances = [];
    if (entSheet) {
      const entRows = entSheet.getDataRange().getValues();
      const msPerDay = 1000 * 60 * 60 * 24;
      for (let i = 1; i < entRows.length; i++) {
        const row    = entRows[i];
        if (row[2] !== 'CASH_ADVANCE') continue;
        if (row[10] === 'DELETED' || row[10] === 'LIQUIDATED') continue;
        const issuedDate = normalizeDate(row[1]);
        const daysOut    = Math.floor((new Date(today) - new Date(issuedDate)) / msPerDay);
        pendingAdvances.push({
          date        : issuedDate,
          description : row[4],
          amount      : parseFloat(row[5]) || 0,
          requestedBy : row[8],
          status      : row[10],
          daysOutstanding: daysOut
        });
      }
      pendingAdvances.sort((a, b) => b.daysOutstanding - a.daysOutstanding);
    }

    // Sort all lists newest first
    const byDateDesc = (a, b) => b.date > a.date ? 1 : -1;
    unclosedDays.sort(byDateDesc);
    discrepancyDays.sort(byDateDesc);
    cashOverDays.sort(byDateDesc);
    unfiledDays.sort(byDateDesc);

    return {
      success: true,
      data: {
        unclosed    : unclosedDays,
        unfiled     : unfiledDays,
        discrepancies: discrepancyDays,
        cashOver    : cashOverDays,
        pendingAdvances: pendingAdvances
      }
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
// AUDITOR METRICS DASHBOARD
// ─────────────────────────────────────────────
function getAuditorMetrics() {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sumRows = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    const entRows = ss.getSheetByName(SHEETS.ENTRIES).getDataRange().getValues();
    const today   = normalizeDate(new Date());

    const pendingAudit = [];   // status === 'PENDING_AUDIT'
    const flaggedDays  = [];   // status === 'FLAGGED'
    const recentClosed = [];   // status === 'CLOSED', last 7 days

    for (let i = 1; i < sumRows.length; i++) {
      const row    = sumRows[i];
      if (row.length < 12) continue;
      const rDate  = normalizeDate(row[1]);
      if (rDate > today) continue;
      const status   = row[11] || 'OPEN';
      const variance = parseFloat(row[10]) || 0;
      const opening  = parseFloat(row[2])  || 0;
      const expenses = parseFloat(row[6])  || 0;
      const closing  = parseFloat(row[9])  || 0;

      if (status === 'PENDING_AUDIT') {
        pendingAudit.push({ date: rDate, variance, opening, expenses, closing });
      }
      if (status === 'FLAGGED') {
        flaggedDays.push({ date: rDate, variance, opening, expenses, closing });
      }
      if (status === 'CLOSED') {
        const msPerDay   = 1000 * 60 * 60 * 24;
        const daysAgo    = Math.floor((new Date(today) - new Date(rDate)) / msPerDay);
        if (daysAgo <= 7) {
          recentClosed.push({ date: rDate, variance, opening, expenses, closing, daysAgo });
        }
      }
    }

    // Pending liquidations (LIQUIDATION_PENDING advances)
    const pendingLiquidations = [];
    for (let i = 1; i < entRows.length; i++) {
      const row    = entRows[i];
      const type   = row[2];
      const status = row[10];
      if (type !== 'CASH_ADVANCE') continue;
      if (status !== 'LIQUIDATION_PENDING') continue;
      const issuedDate = normalizeDate(row[1]);
      const msPerDay   = 1000 * 60 * 60 * 24;
      const daysOut    = Math.floor((new Date(today) - new Date(issuedDate)) / msPerDay);
      pendingLiquidations.push({
        id         : row[0],
        date       : issuedDate,
        description: row[4],
        amount     : row[5],
        requestedBy: row[8],
        daysOut
      });
    }

    // Sort all by date descending
    const byDateDesc = (a, b) => b.date > a.date ? 1 : -1;
    pendingAudit.sort(byDateDesc);
    flaggedDays.sort(byDateDesc);
    recentClosed.sort(byDateDesc);
    pendingLiquidations.sort((a, b) => b.daysOut - a.daysOut);

    return {
      success: true,
      data: {
        pendingAudit,
        flaggedDays,
        recentClosed,
        pendingLiquidations
      }
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getCategories() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet   = ss.getSheetByName('PettyCash_Categories');

    // Auto-create the sheet if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet('PettyCash_Categories');
      sheet.appendRow(['Category']);
      sheet.getRange('A1').setFontWeight('bold');
      const defaults = [
        'Office Supplies','Transportation','Meals & Entertainment',
        'Utilities','Repairs & Maintenance','Postage & Courier','Miscellaneous'
      ];
      defaults.forEach(c => sheet.appendRow([c]));
    }

    const rows = sheet.getDataRange().getValues();
    const categories = [];
    for (let i = 1; i < rows.length; i++) {
      const val = String(rows[i][0] || '').trim();
      if (val) categories.push(val);
    }
    return { success: true, data: categories };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

function syncReceiptsToFinalSheet() {
  const SOURCE_SHEET_NAME   = "PettyCash_Receipts";
  const DEST_SPREADSHEET_ID = "1p7nptmZh-rJF4gjq1S9ntj4-EwCjtBsu_vTc17wahgw";
  const DEST_SHEET_NAME     = "March sample";

  const COLUMNS_TO_COPY = [
    "Date", "Supplier_Name", "Address",
    "TIN", "Receipt_No", "Gross_Amount",
    "Vatable_Sales", "VAT_Amount"
  ];

  const srcSheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_SHEET_NAME);
  const srcData    = srcSheet.getDataRange().getValues();
  const headers    = srcData[0];
  const colIndices = COLUMNS_TO_COPY.map(col => headers.indexOf(col));

  // We still read Receipt_ID from source to use as a hidden tracking key
  const srcRcpIdIdx = headers.indexOf("Receipt_ID");
  const srcDateIdx  = headers.indexOf("Date");

  const destSS    = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);
  const destSheet = destSS.getSheetByName(DEST_SHEET_NAME);
  const destData  = destSheet.getDataRange().getValues();

  // Write header if destination is empty
  if (destData.length === 0 || !destData[0][0]) {
    destSheet.getRange(1, 1, 1, COLUMNS_TO_COPY.length).setValues([COLUMNS_TO_COPY]);
  }

  // ── Build a set of already-synced Receipt_IDs stored in a hidden Notes column ──
  // We store the Receipt_ID in the row's note (invisible to BIR, used for dedup)
  const destRange     = destSheet.getDataRange();
  const existingNotes = new Set();

  if (destSheet.getLastRow() > 1) {
    const noteRange = destSheet.getRange(2, 1, destSheet.getLastRow() - 1, 1);
    noteRange.getNotes().forEach(([note]) => {
      if (note) existingNotes.add(note);
    });
  }

  // ── Append only new rows ──
  const newRows     = [];
  const newRowIds   = []; // track Receipt_IDs for the notes we'll write

  for (let i = 1; i < srcData.length; i++) {
    const receiptId = srcData[i][srcRcpIdIdx];
    if (!receiptId || existingNotes.has(receiptId)) continue;
    newRows.push(colIndices.map(idx => srcData[i][idx]));
    newRowIds.push(receiptId);
  }

  if (newRows.length > 0) {
    const startRow = destSheet.getLastRow() + 1;
    destSheet.getRange(startRow, 1, newRows.length, COLUMNS_TO_COPY.length).setValues(newRows);

    // Store Receipt_ID as a hidden cell note on column A for dedup tracking
    newRowIds.forEach((id, idx) => {
      destSheet.getRange(startRow + idx, 1).setNote(id);
    });
  }
}

function onChange(e) {
  if (e.changeType === "INSERT_ROW" || e.changeType === "EDIT") {
    syncReceiptsToFinalSheet();
  }
}