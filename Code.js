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
  FILING       : 'PettyCash_FilingChecklist',
  CATEGORIES   : 'PettyCash_Categories',
  REQUESTS     : 'PettyCash_Requests'
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
    SpreadsheetApp.flush();
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
      'Total_Cash_Over','Total_Replenishment','Total_Cash_Return',
      'Total_Reimbursement','Closing_Cash','Variance','Status','Closed_By','Updated_At'
    ]);
    s.setFrozenRows(1);
    formatHeaderRow(s);
    s.getRange('C2:L').setNumberFormat('₱#,##0.00');
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

  if (!ss.getSheetByName(SHEETS.CATEGORIES)) {
    const s = ss.insertSheet(SHEETS.CATEGORIES);
    s.appendRow(['Category']);
    s.setFrozenRows(1);
    formatHeaderRow(s);
    s.setColumnWidth(1, 220);
    const defaults = [
      'Office Supplies','Transportation','Meals & Entertainment',
      'Utilities','Repairs & Maintenance','Postage & Courier','Miscellaneous'
    ];
    defaults.forEach(c => s.appendRow([c]));
  }

  if (!ss.getSheetByName(SHEETS.REQUESTS)) {
    const s = ss.insertSheet(SHEETS.REQUESTS);
    s.appendRow([
      'Request_ID', 'Date', 'Purpose', 'Amount', 'Request_Type',
      'Requested_By', 'Submitted_By', 'Status',
      'Approved_By', 'Approved_At', 'Released_At',
      'Rejection_Note', 'Entry_ID', 'Created_At', 'Updated_At'
    ]);
    s.setFrozenRows(1);
    formatHeaderRow(s);
    s.setColumnWidth(1, 160);
    s.setColumnWidth(2, 110);
    s.setColumnWidth(3, 300);
    s.setColumnWidth(4, 100);
    s.setColumnWidth(5, 220);
    s.setColumnWidth(6, 140);
    s.setColumnWidth(7, 220);
    s.setColumnWidth(8, 180);
    s.setColumnWidth(9, 180);
    s.setColumnWidth(10, 180);
    s.setColumnWidth(11, 280);
    s.setColumnWidth(12, 160);
    s.setColumnWidth(13, 180);
    s.setColumnWidth(14, 180);
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
    let totalExp = 0, totalReceipt = 0, totalNoReceipt = 0;
    let cashAdvance = 0, totalCashOver = 0, totalReplenishment = 0, totalCashReturn = 0, totalReimbursement = 0;
    let cashMovementExp = 0;

    for (let i = 1; i < entryData.length; i++) {
      const row    = entryData[i];
      const rDate  = normalizeDate(row[1]);
      const type   = row[2];
      const status = row[10];

      if (status === 'DELETED') continue;
      if (rDate !== date) continue; // ── Skip ALL entries not on this date — no carry-forward in summary

      const amt = parseFloat(row[5]) || 0;

      if (rDate === date) {
        if (type === 'CASH_ADVANCE' || type === 'PCR_ADVANCE') {
          // Cash leaves the drawer on the advance/release date. Detail rows document
          // the eventual spend but must not rewrite the original day's cash balance.
          cashAdvance += amt;
        }
        else if (type === 'CASH_OVER')                  totalCashOver      += amt;
        else if (type === 'REPLENISHMENT')              totalReplenishment += amt;
        else if (type === 'CASH_RETURN')                totalCashReturn    += amt;
        else if (type === 'CASH_ADVANCE_REIMBURSEMENT') totalReimbursement += amt;
        else if (type === 'LIQ_DETAIL' || type === 'PCR_DETAIL') {
          totalExp += amt;
          if (row[6] === 'YES') totalReceipt   += amt;
          else                  totalNoReceipt += amt;
        }
        else {
          totalExp += amt;
          cashMovementExp += amt;
          if (row[6] === 'YES') totalReceipt   += amt;
          else                  totalNoReceipt += amt;
        }
        continue;
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
        existingStatus   = sumData[i][13] || '';
        existingClosedBy = sumData[i][14] || '';
        targetRow = i + 1;
        break;
      }
    }

    const expected = (openingCash + totalReplenishment + totalCashReturn + totalCashOver) - (cashMovementExp + cashAdvance + totalReimbursement);
    const variance = hasClosing ? (closingCash - expected) : 0;
    const status   = hasClosing
      ? (existingStatus === 'CLOSED' ? 'CLOSED' : 'PENDING_AUDIT')
      : 'OPEN';

    const summaryRow = [
      openingCash, cashAdvance,
      totalReceipt, totalNoReceipt, totalExp,
      totalCashOver, totalReplenishment, totalCashReturn,
      totalReimbursement,                // col K — new
      closingCash, variance, status,
      existingClosedBy,
      new Date().toISOString()
    ];

    if (targetRow === -1) {
      const sumId = generateId('SUM', date, sumSheet);
      sumSheet.appendRow([sumId, date, ...summaryRow]);
    } else {
      sumSheet.getRange(targetRow, 3, 1, 14).setValues([summaryRow]);
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
      id,                          // col 1  — Entry_ID
      data.date,                   // col 2  — Date
      data.type        || 'EXPENSE',// col 3  — Type
      data.category    || 'Miscellaneous', // col 4 — Category
      data.description || '',      // col 5  — Description
      parseFloat(data.amount) || 0,// col 6  — Amount
      data.hasReceipt ? 'YES' : 'NO', // col 7 — Has_Receipt
      data.referenceNo || '',      // col 8  — Reference_No
      data.requestedBy || '',      // col 9  — Requested_By
      data.approvedBy  || '',      // col 10 — Approved_By
      'ACTIVE',                    // col 11 — Status
      now,                         // col 12 — Created_At
      now,                         // col 13 — Updated_At
      '',                          // col 14 — (reserved / blank)
      ''                           // col 15 — Notes/Remarks
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

    // Audit log for cash return
    if (data.type === 'CASH_RETURN') {
      writeAuditLog(
        'CASH_RETURN_RECORDED',
        `Cash return of ₱${parseFloat(data.amount).toFixed(2)} recorded. Ref: ${data.referenceNo || '—'}. Desc: ${data.description || '—'}`,
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
      const wasFlagged = dataRange[i][10] === 'FLAGGED';
      if (wasFlagged) {
        sheet.getRange(row, 11).setValue('ACTIVE');
        sheet.getRange(row, 15).setValue('');
      }

      recalculateDailySummary(payload.date);
      if (oldDate && oldDate !== payload.date) recalculateDailySummary(oldDate);

      if (wasFlagged) {
        writeAuditLog(
          'ENTRY_CORRECTED',
          `Flagged entry corrected and reinstated. Desc: ${payload.description || '—'} | Amount: ₱${parseFloat(payload.amount || 0).toFixed(2)} | Category: ${payload.category || '—'}`,
          payload.id,
          payload.date
        );
      }
      writeAuditLog(
        'ENTRY_UPDATED',
        `Entry updated. Desc: ${payload.description || '—'} | Amount: ₱${parseFloat(payload.amount || 0).toFixed(2)} | Category: ${payload.category || '—'} | Type: ${payload.type || '—'}`,
        payload.id,
        payload.date
      );

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
    const nowReceipted = newPayload.hasReceipt === true || newPayload.hasReceipt === 'YES';
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
      sheet.getRange(row, 13).setValue(now);  // updatedAt (col 13)
      sheet.getRange(row, 15).setValue(`[DELETED BY ${userEmail} @ ${now}]`);  // notes (col 15)

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

      // Include LIQUIDATION_PENDING advances from BEFORE this date
      // so auditor sees them in the Audit Review entry list
      if (
        type === 'CASH_ADVANCE' &&
        rowDate < date &&
        status === 'LIQUIDATION_PENDING'
      ) {
        let liqBreakdown   = null;
        const notesCol     = row[14] || '';
        try {
          const jsonStart  = notesCol.indexOf('{');
          if (jsonStart !== -1) liqBreakdown = JSON.parse(notesCol.substring(jsonStart));
        } catch(e) { liqBreakdown = null; }

        entries.push({
          id             : row[0],
          date           : rowDate,
          type           : type,
          category       : row[3],
          description    : row[4],
          amount         : row[5],
          hasReceipt     : row[6] === 'YES',
          referenceNo    : row[7],
          requestedBy    : row[8],
          approvedBy     : row[9],
          status         : status,
          createdAt      : row[11],
          updatedAt      : row[12],
          carriedForward : false,
          originalDate   : rowDate,
          liqBreakdown   : liqBreakdown
        });
        continue;
      }

      // Carry forward unliquidated ACTIVE advances from BEFORE this date
      // LIQUIDATION_PENDING excluded — handled above
      if (
        type === 'CASH_ADVANCE' &&
        rowDate < date &&
        status !== 'LIQUIDATED' &&
        status !== 'DELETED' &&
        status !== 'LIQUIDATION_PENDING'
      ) {
        const msPerDay = 1000 * 60 * 60 * 24;
        const daysOut  = Math.floor(
          (new Date(date) - new Date(rowDate)) / msPerDay
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
          carriedForward  : true,
          daysOutstanding : daysOut,
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
    SpreadsheetApp.flush();
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
    SpreadsheetApp.flush();
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
          totalCashReturn    : row[9],
          totalReimbursement : row[10],
          closingCash        : row[11],
          variance           : row[12],
          status             : row[13],
          closedBy           : row[14],
          updatedAt          : row[15]
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
          totalCashReturn    : 0,
          totalReimbursement : 0,
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
        cashReturn:sRow[9], closing:sRow[11], variance:sRow[12],
        status:sRow[13]
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
      const s = data[i][13];
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
  // data: { date, actualCash, note, flaggedEntries }
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
      sheet.getRange(row, 14).setValue('CLOSED');
      sheet.getRange(row, 15).setValue(email);
      sheet.getRange(row, 16).setValue(now);
      SpreadsheetApp.flush(); // commit CLOSED status before any recalculation reads it
      found = true;

      // ── Finalize any LIQUIDATION_PENDING advances that were verified ──
      finalizePendingLiquidations(data.date, data.flaggedEntries || [], email, now);

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

    // Flush so the END row saved by the frontend just before this call is committed
    SpreadsheetApp.flush();
    const freshDenomRows = denomSheet.getDataRange().getValues();

    // Find auditor's END count for this date
    let endRow = null;
    for (let j = 1; j < freshDenomRows.length; j++) {
      if (normalizeDate(freshDenomRows[j][1]) === data.date && freshDenomRows[j][2] === 'END') {
        endRow = freshDenomRows[j];
        break;
      }
    }

    // Fallback: if still no END row found, create one from actualCash
    if (!endRow && data.actualCash !== undefined) {
      const fallbackId = 'DEN-END-' + data.date.replace(/-/g,'') + '-AU';
      denomSheet.appendRow([
        fallbackId, data.date, 'END',
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        parseFloat(data.actualCash),
        'Auto-saved on audit approval',
        now
      ]);
      SpreadsheetApp.flush();
      const refreshed2 = denomSheet.getDataRange().getValues();
      for (let j = 1; j < refreshed2.length; j++) {
        if (normalizeDate(refreshed2[j][1]) === data.date && refreshed2[j][2] === 'END') {
          endRow = refreshed2[j];
          break;
        }
      }
    }

    if (endRow) {
      const endTotal = parseFloat(endRow[13]) || 0;
      const cfNote   = 'Carried forward from ' + data.date + ' audit closing count';

      // Auditor's count always takes priority — overwrite any existing next-day START
      SpreadsheetApp.flush();
      const latestDenomRows     = denomSheet.getDataRange().getValues();
      let nextDayStartRowIdx    = -1;
      for (let j = 1; j < latestDenomRows.length; j++) {
        if (normalizeDate(latestDenomRows[j][1]) === nextDate && latestDenomRows[j][2] === 'START') {
          nextDayStartRowIdx = j + 1;
          break;
        }
      }

      const denomVals = [
        endRow[3], endRow[4], endRow[5], endRow[6], endRow[7],
        endRow[8], endRow[9], endRow[10], endRow[11], endRow[12],
        endTotal, cfNote, now
      ];

      let cfId;
      if (nextDayStartRowIdx !== -1) {
        // Overwrite the existing START row in-place
        cfId = latestDenomRows[nextDayStartRowIdx - 1][0];
        denomSheet.getRange(nextDayStartRowIdx, 4, 1, 13).setValues([denomVals]);
      } else {
        cfId = 'DEN-OC-' + nextDate.replace(/-/g,'') + '-CF';
        denomSheet.appendRow([cfId, nextDate, 'START', ...denomVals]);
      }

      // Write or update the next day's summary row with the correct openingCash
      const nextSumSheet = ss.getSheetByName(SHEETS.SUMMARY);
      SpreadsheetApp.flush();
      const nextSumData  = nextSumSheet.getDataRange().getValues();
      let nextSumRowIdx  = -1;
      for (let k = 1; k < nextSumData.length; k++) {
        if (normalizeDate(nextSumData[k][1]) === nextDate) { nextSumRowIdx = k + 1; break; }
      }
      if (nextSumRowIdx === -1) {
        const nextSumId = generateId('SUM', nextDate, nextSumSheet);
        // 16 columns: ID, Date, Opening, CashAdv, ExpRcpt, ExpNoRcpt, Expenses,
        //             CashOver, Repl, CashReturn, Reimb, Closing, Variance, Status, ClosedBy, UpdatedAt
        nextSumSheet.appendRow([
          nextSumId, nextDate,
          endTotal, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'OPEN', '', now
        ]);
      } else {
        nextSumSheet.getRange(nextSumRowIdx, 3).setValue(endTotal);
      }

      recalculateDailySummary(nextDate);

      writeAuditLog(
        'OPENING_SAVED',
        `Opening cash auto-carried from ${data.date} audit count. Total: ₱${endTotal.toFixed(2)}`,
        cfId,
        nextDate
      );

      try { autoCloseNonWorkingDays(nextDate, endTotal, ss, now); } catch(e) {
        console.error('autoCloseNonWorkingDays error:', e);
      }
    }

    // ── Sync approved day's receipts to BIR final sheet ──
    try { syncReceiptsToFinalSheet(data.date); } catch(e) {
      console.error('syncReceiptsToFinalSheet error:', e);
      // Non-fatal — don't block approval if sync fails
    }

    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function finalizePendingLiquidations(approvalDate, flaggedEntries, approverEmail, now) {
  // Called during day approval — finalizes all LIQUIDATION_PENDING advances
  // that were NOT flagged. Saves LIQ_DETAIL entries and CASH_RETURN if change > 0.
  try {
    const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
    const entrySheet = ss.getSheetByName(SHEETS.ENTRIES);
    const entryRows  = entrySheet.getDataRange().getValues();

    const flaggedIds = new Set((flaggedEntries || []).map(f => f.id));

    for (let i = 1; i < entryRows.length; i++) {
      const row    = entryRows[i];
      const type   = row[2];
      const status = row[10];

      if (type !== 'CASH_ADVANCE')           continue;
      if (status !== 'LIQUIDATION_PENDING')  continue;

      const advId       = row[0];
      const advDate     = normalizeDate(row[1]);
      const requestedBy = row[8] || '';
      const advAmount   = parseFloat(row[5]) || 0;
      const notesCol    = row[14] || '';

      // If this advance was flagged by auditor — reset to ACTIVE, skip finalization
      if (flaggedIds.has(advId)) {
        entrySheet.getRange(i + 1, 11).setValue('ACTIVE');
        entrySheet.getRange(i + 1, 13).setValue(now);
        entrySheet.getRange(i + 1, 15).setValue(notesCol + ' | [FLAGGED BY AUDITOR] Returned to outstanding.');
        recalculateDailySummary(advDate);
        continue;
      }

      // Parse the breakdown JSON from notes column
      let breakdown = { entries: [], totalSpent: 0, change: 0, note: '' };
      try {
        const jsonStart = notesCol.indexOf('{');
        if (jsonStart !== -1) {
          breakdown = JSON.parse(notesCol.substring(jsonStart));
        }
      } catch(e) {
        console.error('Failed to parse liquidation breakdown for ' + advId, e);
      }

      const totalSpent = parseFloat(breakdown.totalSpent) || 0;
      const change     = parseFloat(breakdown.change)     || 0;

      // 1. Mark advance as LIQUIDATED
      entrySheet.getRange(i + 1, 11).setValue('LIQUIDATED');
      entrySheet.getRange(i + 1, 13).setValue(now);
      entrySheet.getRange(i + 1, 15).setValue(
        notesCol + ' | [LIQUIDATED BY ' + approverEmail + ' on ' + approvalDate + ']'
      );

      // 2. Save LIQ_DETAIL entries on the current approval date (accounting standard)
      if (breakdown.entries && breakdown.entries.length) {
        breakdown.entries.forEach(entry => {
          const liqEntry = saveExpenseEntry({
            date       : approvalDate,
            type       : 'LIQ_DETAIL',
            category   : entry.category   || 'Miscellaneous',
            description: entry.desc       || '',
            amount     : parseFloat(entry.amount) || 0,
            hasReceipt : !!entry.hasReceipt,
            referenceNo: advId,
            requestedBy: requestedBy,
            approvedBy : approverEmail
          });

          // Save receipt if attached
          if (entry.hasReceipt && entry.receipt && liqEntry.id) {
            saveReceiptRecord({
              ...entry.receipt,
              entryId: liqEntry.id,
              date   : approvalDate
            });
          }
        });
      }

      // 3. If there's change, save a CASH_RETURN entry — but only if not already recorded at submission time
      //    If the employee overspent, guard against a duplicate reimbursement entry too
      if (change > 0.005) {
        const alreadyReturned = entryRows.slice(1).some(r =>
          r[2] === 'CASH_RETURN' && String(r[7]) === String(advId) && r[10] !== 'DELETED'
        );
        if (!alreadyReturned) {
          saveExpenseEntry({
            date       : approvalDate,
            type       : 'CASH_RETURN',
            category   : 'Cash Return',
            description: 'Change return — ' + (row[4] || 'Cash Advance') + ' (' + advId + ')',
            amount     : change,
            hasReceipt : false,
            referenceNo: advId,
            requestedBy: requestedBy,
            approvedBy : approverEmail
          });
        }
      } else if (change < -0.005) {
        const alreadyReimbursed = entryRows.slice(1).some(r =>
          r[2] === 'CASH_ADVANCE_REIMBURSEMENT' && String(r[7]) === String(advId) && r[10] !== 'DELETED'
        );
        if (!alreadyReimbursed) {
          saveExpenseEntry({
            date       : approvalDate,
            type       : 'CASH_ADVANCE_REIMBURSEMENT',
            category   : 'Cash Advance Reimbursement',
            description: 'Overage reimbursement — ' + (row[4] || 'Cash Advance') + ' (' + advId + ')',
            amount     : Math.abs(change),
            hasReceipt : false,
            referenceNo: advId,
            requestedBy: requestedBy,
            approvedBy : approverEmail
          });
        }
      }

      // 4. Recalculate summaries for both dates
      recalculateDailySummary(advDate);
      if (approvalDate !== advDate) recalculateDailySummary(approvalDate);

      writeAuditLog(
        'LIQUIDATION_FINALIZED',
        `Advance ${advId} liquidated on day close. Spent: ₱${totalSpent.toFixed(2)}. Change returned: ₱${change.toFixed(2)}.`,
        advId,
        approvalDate
      );
    }
  } catch(e) {
    console.error('finalizePendingLiquidations error:', e);
  }
}

function getNextDate(dateStr) {
  const parts = String(dateStr).split('-').map(Number);
  if (parts.length !== 3 || parts.some(isNaN)) return '';

  const d = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  d.setUTCDate(d.getUTCDate() + 1);

  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function getPreviousDayClosing(date) {
  // Returns the closing total of the most recent day before `date` that is CLOSED
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rows  = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    let best    = null;

    for (let i = 1; i < rows.length; i++) {
      const rDate  = normalizeDate(rows[i][1]);
      const status = rows[i][12];
      if (rDate < date && status === 'CLOSED') {
        if (!best || rDate > best.date) {
          best = { date: rDate, closingCash: parseFloat(rows[i][10]) || 0 };
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
      summaryStatus = sumRows[i][13] || 'OPEN';
      openingCash   = parseFloat(sumRows[i][2])  || 0;
      closingCash   = parseFloat(sumRows[i][11]) || 0;
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
      if (status === 'DELETED' || status === 'LIQUIDATED' || status === 'LIQUIDATION_PENDING') continue;
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

      // Parse liquidation breakdown if present
      let liqBreakdown = null;
      const notesCol   = row[14] || '';
      if (row[10] === 'LIQUIDATION_PENDING' || row[10] === 'LIQUIDATED') {
        try {
          const jsonStart = notesCol.indexOf('{');
          if (jsonStart !== -1) liqBreakdown = JSON.parse(notesCol.substring(jsonStart));
        } catch(e) { liqBreakdown = null; }
      }

      advances.push({
        id             : row[0],
        date           : issuedDate,
        description    : row[4],
        amount         : row[5],
        requestedBy    : row[8],
        status         : row[10],
        daysOutstanding: daysOut,
        liquidationNote: notesCol,
        liqBreakdown   : liqBreakdown
      });
    }

    advances.sort((a, b) => b.daysOutstanding - a.daysOutstanding);
    return { success: true, data: advances };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

function submitLiquidation(data) {
  // Stores breakdown as JSON in the advance's notes column.
  // No EXPENSE rows created here — those are finalized on audit approval.
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const rows  = sheet.getDataRange().getValues();
    const now   = new Date().toISOString();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== data.id) continue;
      const row         = i + 1;
      const advanceDate = normalizeDate(rows[i][1]);

      // Calculate change (advance amount - total spent)
      const advanceAmount = parseFloat(rows[i][5]) || 0;
      const totalSpent    = (data.entries || []).reduce((s, e) => s + (parseFloat(e.amount) || 0), 0);
      const change        = advanceAmount - totalSpent;

      // Store breakdown as JSON in the notes column (col 15)
      const breakdownJson = JSON.stringify({
        entries    : data.entries || [],
        note       : data.note   || '',
        totalSpent : totalSpent,
        change     : change,
        submittedAt: now
      });

      sheet.getRange(row, 11).setValue('LIQUIDATION_PENDING');
      sheet.getRange(row, 13).setValue(now);
      sheet.getRange(row, 15).setValue('[LIQUIDATION SUBMITTED] ' + breakdownJson);

      // If this is a re-submission (edit), void any previously created CASH_RETURN or
      // CASH_ADVANCE_REIMBURSEMENT entries linked to this advance so we don't double-count
      const entrySheet = ss.getSheetByName(SHEETS.ENTRIES);
      const entryRows  = entrySheet.getDataRange().getValues();
      const voidTypes  = ['CASH_RETURN', 'CASH_ADVANCE_REIMBURSEMENT'];
      for (let j = 1; j < entryRows.length; j++) {
        if (voidTypes.includes(entryRows[j][2]) &&
            String(entryRows[j][7]) === String(data.id) &&
            entryRows[j][10] !== 'DELETED') {
          entrySheet.getRange(j + 1, 11).setValue('DELETED');
          entrySheet.getRange(j + 1, 13).setValue(now);
          entrySheet.getRange(j + 1, 15).setValue('[VOIDED — liquidation re-submitted]');
        }
      }

      // Immediately record the change as CASH_RETURN — cashier has physically returned it on submission
      // Or reimburse the employee from petty cash if they overspent
      const today = normalizeDate(new Date());
      if (change > 0.005) {
        saveExpenseEntry({
          date       : today,
          type       : 'CASH_RETURN',
          category   : 'Cash Return',
          description: 'Change return — ' + (rows[i][4] || 'Cash Advance') + ' (' + data.id + ')',
          amount     : change,
          hasReceipt : false,
          referenceNo: data.id,
          requestedBy: rows[i][8] || '',
          approvedBy : ''
        });
      } else if (change < -0.005) {
        saveExpenseEntry({
          date       : today,
          type       : 'CASH_ADVANCE_REIMBURSEMENT',
          category   : 'Cash Advance Reimbursement',
          description: 'Overage reimbursement — ' + (rows[i][4] || 'Cash Advance') + ' (' + data.id + ')',
          amount     : Math.abs(change),
          hasReceipt : false,
          referenceNo: data.id,
          requestedBy: rows[i][8] || '',
          approvedBy : ''
        });
      }
      recalculateDailySummary(advanceDate);
      if (today !== advanceDate) recalculateDailySummary(today);

      writeAuditLog(
        'LIQUIDATION_SUBMITTED',
        `Liquidation submitted for advance ${data.id}. ${(data.entries || []).length} entries. Total spent: ₱${totalSpent.toFixed(2)}. Change: ₱${change.toFixed(2)} (immediately returned). Notes: ${data.note || '—'}`,
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

    sheet.getRange(targetRow, 4, 1, 13).setValues([[
      denoms['1000'] ||0, denoms['500']||0, denoms['200']||0,
      denoms['100']  ||0, denoms['50'] ||0, denoms['20'] ||0,
      denoms['10']   ||0, denoms['5']  ||0, denoms['1']  ||0,
      denoms['0.25'] ||0,
      total, data.notes || '', now
    ]]);

    recalculateDailySummary(data.date);
    writeAuditLog(
      'DENOMINATION_UPDATED',
      `${data.type} denomination count updated. Total: ₱${total.toFixed(2)}`,
      dataRange[targetRow - 1][0],
      data.date
    );
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
        const status = sumRows[i][13];
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
      if (sumRows[i][13] === 'CLOSED') closedMap[normalizeDate(sumRows[i][1])] = true;
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
      sheet.getRange(row, 14).setValue('FLAGGED');
      sheet.getRange(row, 15).setValue(email);
      sheet.getRange(row, 16).setValue(now);

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
      const status   = row[13] || 'OPEN';
      const variance = parseFloat(row[12]) || 0;
      const cashOver = parseFloat(row[7])  || 0;
      const opening  = parseFloat(row[2])  || 0;
      const expenses = parseFloat(row[6])  || 0;
      const closing  = parseFloat(row[11]) || 0;

      // Unclosed = anything not CLOSED (exclude today as it is still active)
      if (status !== 'CLOSED' && rDate !== today) {
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
      if ((row[13] || '') === 'CLOSED') {
        const d = normalizeDate(row[1]);
        if (d <= today) closedDates[d] = { date: d, opening: parseFloat(row[2])||0, expenses: parseFloat(row[6])||0, closing: parseFloat(row[11])||0 };
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
      const status   = row[13] || 'OPEN';
      const variance = parseFloat(row[12]) || 0;
      const opening  = parseFloat(row[2])  || 0;
      const expenses = parseFloat(row[6])  || 0;
      const closing  = parseFloat(row[11]) || 0;

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

// ─────────────────────────────────────────────
// DATA REPAIR UTILITY
// Run once from the Apps Script editor to fix Summary rows corrupted by the
// old bug where auditApproveDay/flagDay wrote status to col 12 (Variance)
// instead of col 13 (Status).
// ─────────────────────────────────────────────
function repairSummarySheet() {
  try {
    const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sumSheet = ss.getSheetByName(SHEETS.SUMMARY);
    if (!sumSheet) return { success: false, message: 'Summary sheet not found' };

    const rows = sumSheet.getDataRange().getValues();
    let repaired = 0;
    const STATUS_STRINGS = ['CLOSED', 'FLAGGED', 'PENDING_AUDIT', 'OPEN'];

    for (let i = 1; i < rows.length; i++) {
      const row   = rows[i];
      const col12 = row[11]; // index 11 = col L = Closing_Cash (old rows: may hold status string due to prior bug)
      const col13 = row[12]; // index 12 = col M = Variance
      const col14 = row[13]; // index 13 = col N = Status
      const col15 = row[14]; // index 14 = col O = Closed_By
      const date  = normalizeDate(row[1]);

      // Detect corrupted row: col 12 holds a status string (bug wrote status there)
      const isCorrupted = STATUS_STRINGS.includes(String(col12));
      if (!isCorrupted) continue;

      const corruptedStatus = String(col12); // e.g. 'CLOSED'
      const corruptedEmail  = String(col13 || '');
      const corruptedTs     = String(col14 || new Date().toISOString());

      // Recompute all totals from Entries sheet
      const entryData = ss.getSheetByName(SHEETS.ENTRIES).getDataRange().getValues();
      let totalExp = 0, totalReceipt = 0, totalNoReceipt = 0;
      let cashAdvance = 0, totalCashOver = 0, totalReplenishment = 0, totalCashReturn = 0, totalReimbursement = 0;
      let cashMovementExp = 0;

      for (let j = 1; j < entryData.length; j++) {
        const eRow    = entryData[j];
        const rDate   = normalizeDate(eRow[1]);
        const type    = eRow[2];
        const eStatus = eRow[10];
        const amt     = parseFloat(eRow[5]) || 0;
        if (rDate !== date || eStatus === 'DELETED') continue;

        if (type === 'CASH_ADVANCE') {
          cashAdvance += amt;
        } else if (type === 'CASH_ADVANCE_REIMBURSEMENT') { totalReimbursement += amt; }
          else if (type === 'CASH_OVER')     { totalCashOver      += amt; }
          else if (type === 'REPLENISHMENT') { totalReplenishment += amt; }
          else if (type === 'CASH_RETURN')   { totalCashReturn    += amt; }
          else if (type === 'LIQ_DETAIL')    {
            totalExp += amt;
            if (eRow[6] === 'YES') totalReceipt   += amt;
            else                   totalNoReceipt += amt;
          }
          else {
            totalExp += amt;
            cashMovementExp += amt;
            if (eRow[6] === 'YES') totalReceipt   += amt;
            else                   totalNoReceipt += amt;
          }
      }

      // Get opening and closing cash from Denominations sheet
      const denomData = ss.getSheetByName(SHEETS.DENOMINATIONS).getDataRange().getValues();
      let openingCash = parseFloat(row[2]) || 0; // preserve existing opening as fallback
      let closingCash = 0;
      for (let k = 1; k < denomData.length; k++) {
        if (normalizeDate(denomData[k][1]) !== date) continue;
        if (denomData[k][2] === 'START') openingCash = parseFloat(denomData[k][13]) || openingCash;
        if (denomData[k][2] === 'END')   closingCash = parseFloat(denomData[k][13]) || 0;
      }

      const expected = (openingCash + totalReplenishment + totalCashReturn + totalCashOver) - (cashMovementExp + cashAdvance + totalReimbursement);
      const variance = closingCash - expected;
      const rowNum   = i + 1;

      sumSheet.getRange(rowNum, 3, 1, 14).setValues([[
        openingCash, cashAdvance,
        totalReceipt, totalNoReceipt, totalExp,
        totalCashOver, totalReplenishment, totalCashReturn,
        totalReimbursement,            // col K — new
        closingCash, variance,         // col L = Closing, col M = Variance
        corruptedStatus,               // col N = Status
        corruptedEmail,                // col O = Closed_By
        corruptedTs                    // col P = Updated_At
      ]]);

      repaired++;
      console.log('Repaired row for date: ' + date + ' (' + corruptedStatus + ')');
    }

    SpreadsheetApp.flush();
    return { success: true, repaired, message: 'Repaired ' + repaired + ' corrupted Summary row(s).' };
  } catch(e) {
    console.error('repairSummarySheet error:', e);
    return { success: false, message: e.toString() };
  }
}

function getCategories() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.CATEGORIES);
    if (!sheet) return { success: false, message: 'Categories sheet not found', data: [] };

    const rows = sheet.getDataRange().getValues();
    const categories = [];
    for (let i = 1; i < rows.length; i++) {
      const name = String(rows[i][0] || '').trim();
      const desc = String(rows[i][1] || '').trim();
      if (name) categories.push({ name, description: desc });
    }
    return { success: true, data: categories };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

function getEmployees() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PettyCash_Employees');
    if (!sheet) return { success: false, message: 'Employees sheet not found', data: [] };

    const rows = sheet.getDataRange().getValues();
    const employees = [];
    for (let i = 1; i < rows.length; i++) {
      const name = String(rows[i][0] || '').trim();
      if (name) employees.push(name);
    }
    return { success: true, data: employees };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// ─────────────────────────────────────────────
// SUMMARY REPORT DATA
// ─────────────────────────────────────────────
function getSummaryReportData(params) {
  // params: { month: 0-11, year: YYYY }
  try {
    const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    const year     = parseInt(params.year);
    const month    = parseInt(params.month); // 0-indexed
    const from     = new Date(year, month, 1);
    const to       = new Date(year, month + 1, 0);
    const fromStr  = normalizeDate(from);
    const toStr    = normalizeDate(to);

    // ── Read all entries in range ──
    const entryData = ss.getSheetByName(SHEETS.ENTRIES).getDataRange().getValues();
    const categoryTotals   = {};
    let totalWithReceipt   = 0;
    let totalWithoutReceipt= 0;
    let totalExpenses      = 0;
    let advancesIssued     = 0;
    let advancesLiquidated = 0;
    let advancesOutstanding= 0;
    let totalReplenishment = 0;
    let totalCashReturn    = 0;

    for (let i = 1; i < entryData.length; i++) {
      const row    = entryData[i];
      if (row.length < 11) continue;
      const rowDate = normalizeDate(row[1]);
      const type    = row[2];
      const status  = row[10];
      const amount  = parseFloat(row[5]) || 0;

      if (rowDate < fromStr || rowDate > toStr) continue;
      if (status === 'DELETED') continue;

      if (type === 'EXPENSE' || type === 'LIQ_DETAIL') {
        const cat = String(row[3] || 'Miscellaneous').trim();
        categoryTotals[cat] = (categoryTotals[cat] || 0) + amount;
        totalExpenses += amount;
        if (row[6] === 'YES') totalWithReceipt    += amount;
        else                  totalWithoutReceipt += amount;
      } else if (type === 'CASH_ADVANCE') {
        advancesIssued += amount;
        if (status === 'LIQUIDATED')   advancesLiquidated  += amount;
        else if (status !== 'DELETED') advancesOutstanding += amount;
      } else if (type === 'REPLENISHMENT') {
        totalReplenishment += amount;
      } else if (type === 'CASH_RETURN') {
        totalCashReturn += amount;
      } else if (type === 'CASH_ADVANCE_REIMBURSEMENT') {
        advancesIssued += amount; // outflow: petty cash paid employee back for overspend
      }
    }

    // ── Read summaries in range (for daily breakdown) ──
    const sumData  = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    const dailyRows = [];
    for (let i = 1; i < sumData.length; i++) {
      const rowDate = normalizeDate(sumData[i][1]);
      if (rowDate < fromStr || rowDate > toStr) continue;
      const expenses = parseFloat(sumData[i][6]) || 0;
      const opening  = parseFloat(sumData[i][2]) || 0;
      const replenish= parseFloat(sumData[i][8]) || 0;
      const closing  = parseFloat(sumData[i][11]) || 0;
      const status   = sumData[i][13] || 'OPEN';
      // Only include days with actual activity
      if (expenses === 0 && opening === 0 && replenish === 0) continue;
      dailyRows.push({
        date         : rowDate,
        opening      : opening,
        cashAdvance  : parseFloat(sumData[i][3]) || 0,
        expenses     : expenses,
        replenishment: replenish,
        closing      : closing,
        status       : status
      });
    }
    dailyRows.sort((a, b) => a.date > b.date ? 1 : -1);

    // ── Find latest closing cash (cash on hand) ──
    let cashOnHand = 0;
    if (dailyRows.length > 0) {
      const lastClosed = [...dailyRows].reverse().find(d => d.status === 'CLOSED');
      cashOnHand = lastClosed ? lastClosed.closing : 0;
    }

    // ── Compute amount to replenish ──
    const FUND_CEILING    = 28000;
    const accounted       = cashOnHand + advancesOutstanding;
    const toReplenish     = Math.max(0, FUND_CEILING - accounted);

    // ── Format category breakdown ──
    const categories = Object.entries(categoryTotals)
      .map(([name, amount]) => ({
        name,
        amount,
        percent: totalExpenses > 0 ? Math.round((amount / totalExpenses) * 100) : 0
      }))
      .sort((a, b) => b.amount - a.amount);

    return {
      success: true,
      data: {
        period             : { from: fromStr, to: toStr },
        fundCeiling        : FUND_CEILING,
        totalExpenses,
        totalWithReceipt,
        totalWithoutReceipt,
        totalReplenishment,
        totalCashReturn,
        cashOnHand,
        advancesIssued,
        advancesLiquidated,
        advancesOutstanding,
        accounted,
        toReplenish,
        categories,
        dailyRows
      }
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
// REPLENISHMENT PERIOD REPORT
// ─────────────────────────────────────────────
function getReplenishmentPeriodReport() {
  try {
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const entryData = ss.getSheetByName(SHEETS.ENTRIES).getDataRange().getValues();
    const today     = normalizeDate(new Date());

    // ── Collect all unique replenishment dates (sorted) ──
    const replenishDates = [];
    for (let i = 1; i < entryData.length; i++) {
      const row    = entryData[i];
      const type   = row[2];
      const status = row[10];
      const date   = normalizeDate(row[1]);
      if (type !== 'REPLENISHMENT' || status === 'DELETED') continue;
      if (!replenishDates.includes(date)) replenishDates.push(date);
    }
    replenishDates.sort();

    // ── Determine period start ──
    // If 2+ replenishments: period starts day AFTER the second-to-last one
    // If only 1: period starts from the earliest summary date
    // If none: show everything
    const sumData = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    let fromStr = null;

    if (replenishDates.length >= 1) {
      fromStr = replenishDates[replenishDates.length - 1]; // Start ON the day of the last replenishment
    } else {
      const dates = [];
      for (let i = 1; i < sumData.length; i++) {
        const d = normalizeDate(sumData[i][1]);
        if (d) dates.push(d);
      }
      dates.sort();
      fromStr = dates[0] || today;
    }

    const toStr = today;

    // ── Collect all entries in period ──
    let totalExpenses       = 0;
    let totalReplenishment  = 0;
    let advancesOutstanding = 0;
    const periodEntries     = [];

    for (let i = 1; i < entryData.length; i++) {
      const row     = entryData[i];
      if (row.length < 11) continue;
      const rowDate = normalizeDate(row[1]);
      const type    = row[2];
      const status  = row[10];
      const amount  = parseFloat(row[5]) || 0;

      if (rowDate < fromStr || rowDate > toStr) continue;
      if (status === 'DELETED') continue;

      if (type === 'EXPENSE' || type === 'LIQ_DETAIL') {
        totalExpenses += amount;
        periodEntries.push({
          id: row[0], date: rowDate, type, category: row[3],
          description: row[4], amount, requestedBy: row[8], status
        });
      } else if (type === 'REPLENISHMENT') {
        totalReplenishment += amount;
      } else if (type === 'CASH_ADVANCE_REIMBURSEMENT') {
        totalExpenses += amount; // outflow counted against fund usage
        periodEntries.push({
          id: row[0], date: rowDate, type, category: row[3],
          description: row[4], amount, requestedBy: row[8], status
        });
      } else if (type === 'CASH_ADVANCE' &&
                (status === 'ACTIVE' || status === 'LIQUIDATION_PENDING')) {
        advancesOutstanding += amount;
        periodEntries.push({
          id: row[0], date: rowDate, type, category: row[3],
          description: row[4], amount, requestedBy: row[8], status
        });
      }
    }

    periodEntries.sort((a, b) => a.date > b.date ? 1 : -1);

    // ── Daily breakdown from summary — correct column indexes ──
    const dailyRows = [];
    for (let i = 1; i < sumData.length; i++) {
      const rowDate   = normalizeDate(sumData[i][1]);
      if (rowDate < fromStr || rowDate > toStr) continue;
      const opening   = parseFloat(sumData[i][2])  || 0;
      const cashAdv   = parseFloat(sumData[i][3])  || 0;
      const expenses  = parseFloat(sumData[i][6])  || 0;
      const replenish = parseFloat(sumData[i][8])  || 0;
      const closing   = parseFloat(sumData[i][11]) || 0;
      const status    = String(sumData[i][13] || 'OPEN');
      if (expenses === 0 && opening === 0 && replenish === 0) continue;
      dailyRows.push({
        date         : rowDate,
        opening,
        cashAdvance  : cashAdv,
        expenses,
        replenishment: replenish,
        closing,
        status
      });
    }
    dailyRows.sort((a, b) => a.date > b.date ? 1 : -1);

    // ── Cash on hand = latest CLOSED day's closing ──
    const lastClosed   = [...dailyRows].reverse().find(d => d.status === 'CLOSED');
    const cashOnHand   = lastClosed ? lastClosed.closing : 0;
    const FUND_CEILING = 28000;
    const accounted    = cashOnHand + advancesOutstanding;
    const toReplenish  = Math.max(0, FUND_CEILING - accounted);

    return {
      success: true,
      data: {
        periodFrom         : fromStr,
        periodTo           : toStr,
        fundCeiling        : FUND_CEILING,
        totalExpenses,
        totalReplenishment,
        cashOnHand,
        advancesOutstanding,
        accounted,
        toReplenish,
        dailyRows,
        periodEntries
      }
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function syncReceiptsToFinalSheet(approvedDate) {
  const SOURCE_SHEET_NAME   = "PettyCash_Receipts";
  const DEST_SPREADSHEET_ID = "1p7nptmZh-rJF4gjq1S9ntj4-EwCjtBsu_vTc17wahgw";

  const MONTHS = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE',
                  'JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];

  const COLUMNS_TO_COPY = [
    "Date", "Supplier_Name", "Address",
    "TIN", "Receipt_No", "Gross_Amount",
    "Vatable_Sales", "VAT_Amount"
  ];

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const srcSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!srcSheet) return;
  const srcData    = srcSheet.getDataRange().getValues();
  const headers    = srcData[0];
  const colIndices = COLUMNS_TO_COPY.map(col => headers.indexOf(col));

  const srcRcpIdIdx    = headers.indexOf("Receipt_ID");
  const srcDateIdx     = headers.indexOf("Date");
  const srcEntryIdIdx  = headers.indexOf("Entry_ID");

  // ── Build a map of entryId → {status, hasReceipt} to filter out orphaned receipts ──
  // Receipts linked to DELETED entries or entries where receipt was later removed must be excluded.
  const entrySheet = ss.getSheetByName(SHEETS.ENTRIES);
  const entryRows  = entrySheet ? entrySheet.getDataRange().getValues() : [];
  const entryMap   = {};
  for (let i = 1; i < entryRows.length; i++) {
    entryMap[entryRows[i][0]] = { status: entryRows[i][10], hasReceipt: entryRows[i][6] };
  }

  const destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);

  // ── Group source receipts by their actual receipt month ──
  // Only process receipts matching the approved date
  const receiptsByMonth = {};

  for (let i = 1; i < srcData.length; i++) {
    const receiptId   = srcData[i][srcRcpIdIdx];
    const receiptDate = normalizeDate(srcData[i][srcDateIdx]);
    const entryId     = srcData[i][srcEntryIdIdx];

    if (!receiptId)                          continue;
    if (approvedDate && receiptDate !== approvedDate) continue; // only this day's receipts

    // Skip receipts whose linked entry was deleted or no longer marked as receipted
    const entry = entryMap[entryId];
    if (entry && (entry.status === 'DELETED' || entry.hasReceipt !== 'YES')) continue;

    // Determine which monthly tab this receipt belongs to
    const dateObj   = new Date(receiptDate + 'T00:00:00');
    const tabName   = MONTHS[dateObj.getMonth()];

    if (!receiptsByMonth[tabName]) receiptsByMonth[tabName] = [];
    receiptsByMonth[tabName].push({ row: srcData[i], id: receiptId });
  }

  // ── For each monthly tab, append only new receipts ──
  for (const [tabName, receipts] of Object.entries(receiptsByMonth)) {
    let destSheet = destSS.getSheetByName(tabName);

    // Auto-create the monthly tab if it doesn't exist yet
    if (!destSheet) {
      destSheet = destSS.insertSheet(tabName);
      destSheet.getRange(1, 1, 1, COLUMNS_TO_COPY.length).setValues([
        ['DATE','NAME OF SUPPLIER','ADDRESS','TIN #','RECEIPT No.','AMOUNT','LESS:VAT','VAT-12%']
      ]);
      destSheet.setFrozenRows(1);
      destSheet.getRange('A1:H1').setFontWeight('bold');
    }

    // Write header if sheet is empty
    const destData = destSheet.getDataRange().getValues();
    if (destData.length === 0 || !destData[0][0]) {
      destSheet.getRange(1, 1, 1, COLUMNS_TO_COPY.length).setValues([
        ['DATE','NAME OF SUPPLIER','ADDRESS','TIN #','RECEIPT No.','AMOUNT','LESS:VAT','VAT-12%']
      ]);
    }

    // Build set of already-synced Receipt_IDs from hidden cell notes
    const existingNotes = new Set();
    if (destSheet.getLastRow() > 1) {
      destSheet.getRange(2, 1, destSheet.getLastRow() - 1, 1)
        .getNotes()
        .forEach(([note]) => { if (note) existingNotes.add(note); });
    }

    // Append only new rows
    const newRows   = [];
    const newRowIds = [];

    for (const { row, id } of receipts) {
      if (existingNotes.has(id)) continue;
      newRows.push(colIndices.map(idx => row[idx]));
      newRowIds.push(id);
    }

    if (newRows.length > 0) {
      const startRow = destSheet.getLastRow() + 1;
      destSheet.getRange(startRow, 1, newRows.length, COLUMNS_TO_COPY.length).setValues(newRows);
      newRowIds.forEach((id, idx) => {
        destSheet.getRange(startRow + idx, 1).setNote(id);
      });
    }
  }
}

// ─────────────────────────────────────────────
// AUTO-CLOSE NON-WORKING DAYS (Sundays)
// ─────────────────────────────────────────────
function autoCloseNonWorkingDays(startDate, openingCash, ss, now) {
  try {
    let checkDate = startDate;

    while (true) {
      const dateObj  = new Date(checkDate + 'T00:00:00');
      const isSunday = dateObj.getDay() === 0; // 0 = Sunday

      if (!isSunday) break;

      const sumSheet   = ss.getSheetByName(SHEETS.SUMMARY);
      const denomSheet = ss.getSheetByName(SHEETS.DENOMINATIONS);
      const sumData    = sumSheet.getDataRange().getValues();

      // Skip if summary record already exists for this date
      let alreadyExists = false;
      for (let i = 1; i < sumData.length; i++) {
        if (normalizeDate(sumData[i][1]) === checkDate) {
          alreadyExists = true;
          break;
        }
      }

      if (!alreadyExists) {
        // ── Create Summary row — CLOSED, zero activity ──
        const sumId = generateId('SUM', checkDate, sumSheet);
        sumSheet.appendRow([
          sumId,        // Summary_ID
          checkDate,    // Date
          openingCash,  // Opening_Cash
          0,            // Cash_Advance
          0,            // Total_Exp_With_Receipt
          0,            // Total_Exp_No_Receipt
          0,            // Total_Expenses
          0,            // Total_Cash_Over
          0,            // Total_Replenishment
          0,            // Total_Cash_Return
          0,            // Total_Reimbursement
          openingCash,  // Closing_Cash (same as opening)
          0,            // Variance
          'CLOSED',     // Status
          'system',     // Closed_By
          now           // Updated_At
        ]);

        // ── Create Denomination START row (carry-forward from previous day END) ──
        const prevDate   = getPreviousDate(checkDate);
        const denomData  = denomSheet.getDataRange().getValues();
        let   prevEndRow = null;

        for (let i = 1; i < denomData.length; i++) {
          if (normalizeDate(denomData[i][1]) === prevDate && denomData[i][2] === 'END') {
            prevEndRow = denomData[i];
            break;
          }
        }

        if (prevEndRow) {
          const startId = 'DEN-OC-' + checkDate.replace(/-/g,'') + '-CF';
          denomSheet.appendRow([
            startId,         // Record_ID
            checkDate,       // Date
            'START',         // Type
            prevEndRow[3],   // ₱1000
            prevEndRow[4],   // ₱500
            prevEndRow[5],   // ₱200
            prevEndRow[6],   // ₱100
            prevEndRow[7],   // ₱50
            prevEndRow[8],   // ₱20
            prevEndRow[9],   // ₱10
            prevEndRow[10],  // ₱5
            prevEndRow[11],  // ₱1
            prevEndRow[12],  // ₱0.25
            openingCash,     // Total
            'Auto-closed: Non-working day (Sunday)',
            now
          ]);

          // ── Create Denomination END row (same as START — no activity) ──
          const endId = 'DEN-CC-' + checkDate.replace(/-/g,'') + '-CF';
          denomSheet.appendRow([
            endId,           // Record_ID
            checkDate,       // Date
            'END',           // Type
            prevEndRow[3],   // ₱1000
            prevEndRow[4],   // ₱500
            prevEndRow[5],   // ₱200
            prevEndRow[6],   // ₱100
            prevEndRow[7],   // ₱50
            prevEndRow[8],   // ₱20
            prevEndRow[9],   // ₱10
            prevEndRow[10],  // ₱5
            prevEndRow[11],  // ₱1
            prevEndRow[12],  // ₱0.25
            openingCash,     // Total
            'Auto-closed: Non-working day (Sunday)',
            now
          ]);
        }

        writeAuditLog(
          'DAY_APPROVED',
          `Sunday auto-closed as non-working day. Opening/Closing: ₱${openingCash.toFixed(2)}`,
          sumId,
          checkDate
        );
      }

      // Move to next day and keep checking
      // (handles holiday Monday after Sunday, etc.)
      checkDate = getNextDate(checkDate);
    }
  } catch(e) {
    console.error('autoCloseNonWorkingDays error:', e);
  }
}

function getPreviousDate(dateStr) {
  const parts = String(dateStr).split('-').map(Number);
  if (parts.length !== 3 || parts.some(isNaN)) return '';

  const d = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  d.setUTCDate(d.getUTCDate() - 1);

  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

// ─────────────────────────────────────────────
// SUPPLIER AUTOCOMPLETE
// Reads from PettyCash_Receipts (main) + BIR dest sheet
// Returns all unique name+address+TIN combinations
// ─────────────────────────────────────────────
function getSuppliers() {
  try {
    const BIR_SPREADSHEET_ID = '1p7nptmZh-rJF4gjq1S9ntj4-EwCjtBsu_vTc17wahgw';
    const seen    = new Map(); // key: "NAME|||ADDRESS|||TIN" → true (dedup)
    const grouped = new Map(); // key: NAME_LOWER → { name, variants: [] }

    function processRow(name, address, tin) {
      name    = String(name    || '').trim();
      address = String(address || '').trim();
      tin     = String(tin     || '').trim();
      if (!name) return;

      const dedupeKey = `${name.toLowerCase()}|||${address.toLowerCase()}|||${tin.toLowerCase()}`;
      if (seen.has(dedupeKey)) return;
      seen.set(dedupeKey, true);

      const groupKey = name.toLowerCase();
      if (!grouped.has(groupKey)) grouped.set(groupKey, { name, variants: [] });
      grouped.get(groupKey).variants.push({ address, tin });
    }

    // ── 1. Read main PettyCash_Receipts ──
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.RECEIPTS);
    if (sheet) {
      const rows    = sheet.getDataRange().getValues();
      const headers = rows[0];
      const nameIdx = headers.indexOf('Supplier_Name');
      const addrIdx = headers.indexOf('Address');
      const tinIdx  = headers.indexOf('TIN');
      for (let i = 1; i < rows.length; i++) {
        processRow(rows[i][nameIdx], rows[i][addrIdx], rows[i][tinIdx]);
      }
    }

    // ── 2. Read BIR historical spreadsheet (all monthly tabs) ──
    // Columns: A=Date, B=Supplier_Name, C=Address, D=TIN
    try {
      const birSS    = SpreadsheetApp.openById(BIR_SPREADSHEET_ID);
      const birSheets = birSS.getSheets();
      for (const tab of birSheets) {
        const rows = tab.getDataRange().getValues();
        for (let i = 1; i < rows.length; i++) {
          processRow(rows[i][1], rows[i][2], rows[i][3]); // B, C, D
        }
      }
    } catch(e) {
      console.warn('getSuppliers: BIR sheet read failed (non-fatal):', e);
    }

    // ── 3. Sort alphabetically and return ──
    const result = [];
    for (const [, entry] of grouped) {
      result.push(entry);
    }
    result.sort((a, b) => a.name.localeCompare(b.name));

    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// ─────────────────────────────────────────────
// SUMMARY REPORT — date-range aware
// params: { from: 'YYYY-MM-DD', to: 'YYYY-MM-DD' }
// ─────────────────────────────────────────────
function getSummaryReportDataByRange(params) {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const fromStr = params.from;
    const toStr   = params.to;

    const entryData = ss.getSheetByName(SHEETS.ENTRIES).getDataRange().getValues();
    const categoryTotals    = {};
    let totalWithReceipt    = 0;
    let totalWithoutReceipt = 0;
    let totalExpenses       = 0;
    let advancesIssued      = 0;
    let advancesLiquidated  = 0;
    let advancesOutstanding = 0;
    let totalReplenishment  = 0;
    let totalCashReturn     = 0;

    // Also collect entry lists for receipt modal
    const withReceiptEntries    = [];
    const withoutReceiptEntries = [];

    for (let i = 1; i < entryData.length; i++) {
      const row     = entryData[i];
      if (row.length < 11) continue;
      const rowDate = normalizeDate(row[1]);
      const type    = row[2];
      const status  = row[10];
      const amount  = parseFloat(row[5]) || 0;

      if (rowDate < fromStr || rowDate > toStr) continue;
      if (status === 'DELETED') continue;

      if (type === 'EXPENSE' || type === 'LIQ_DETAIL') {
        const cat = String(row[3] || 'Miscellaneous').trim();
        categoryTotals[cat] = (categoryTotals[cat] || 0) + amount;
        totalExpenses += amount;

        const entry = {
          id         : row[0],
          date       : rowDate,
          category   : row[3],
          description: row[4],
          amount     : amount,
          referenceNo: row[7],
          requestedBy: row[8]
        };

        if (row[6] === 'YES') {
          totalWithReceipt += amount;
          withReceiptEntries.push(entry);
        } else {
          totalWithoutReceipt += amount;
          withoutReceiptEntries.push(entry);
        }
      } else if (type === 'CASH_ADVANCE') {
        advancesIssued += amount;
        if (status === 'LIQUIDATED')   advancesLiquidated  += amount;
        else if (status !== 'DELETED') advancesOutstanding += amount;
      } else if (type === 'REPLENISHMENT') {
        totalReplenishment += amount;
      } else if (type === 'CASH_RETURN') {
        totalCashReturn += amount;
      } else if (type === 'CASH_ADVANCE_REIMBURSEMENT') {
        advancesIssued += amount; // outflow: petty cash paid employee back for overspend
      }
    }

    // ── Join receipt details from PettyCash_Receipts ──
    const rcptSheet = ss.getSheetByName(SHEETS.RECEIPTS);
    const rcptData  = rcptSheet ? rcptSheet.getDataRange().getValues() : [];
    const rcptMap   = {};
    for (let i = 1; i < rcptData.length; i++) {
      const r = rcptData[i];
      rcptMap[r[1]] = {
        supplierName: r[3],
        address     : r[4],
        tin         : r[5],
        receiptNo   : r[6],
        grossAmount : r[7]
      };
    }
    withReceiptEntries.forEach(e => {
      if (rcptMap[e.id]) Object.assign(e, rcptMap[e.id]);
    });

    withReceiptEntries.sort((a, b) => a.date > b.date ? 1 : -1);
    withoutReceiptEntries.sort((a, b) => a.date > b.date ? 1 : -1);

    // ── Summary rows ──
    const sumData  = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    const dailyRows = [];
    for (let i = 1; i < sumData.length; i++) {
      const rowDate = normalizeDate(sumData[i][1]);
      if (rowDate < fromStr || rowDate > toStr) continue;
      const expenses = parseFloat(sumData[i][6]) || 0;
      const opening  = parseFloat(sumData[i][2]) || 0;
      const replenish= parseFloat(sumData[i][8]) || 0;
      const closing  = parseFloat(sumData[i][11]) || 0;
      const status   = String(sumData[i][13] || 'OPEN');
      if (expenses === 0 && opening === 0 && replenish === 0) continue;
      dailyRows.push({
        date: rowDate, opening,
        cashAdvance: parseFloat(sumData[i][3]) || 0,
        expenses, replenishment: replenish, closing, status
      });
    }
    dailyRows.sort((a, b) => a.date > b.date ? 1 : -1);

    const lastClosed  = [...dailyRows].reverse().find(d => d.status === 'CLOSED');
    const cashOnHand  = lastClosed ? lastClosed.closing : 0;
    const FUND_CEILING = 28000;
    const accounted   = cashOnHand + advancesOutstanding;
    const toReplenish = Math.max(0, FUND_CEILING - accounted);

    const categories = Object.entries(categoryTotals)
      .map(([name, amount]) => ({
        name, amount,
        percent: totalExpenses > 0 ? Math.round((amount / totalExpenses) * 100) : 0
      }))
      .sort((a, b) => b.amount - a.amount);

    return {
      success: true,
      data: {
        period: { from: fromStr, to: toStr },
        fundCeiling: FUND_CEILING,
        totalExpenses, totalWithReceipt, totalWithoutReceipt,
        totalReplenishment, totalCashReturn,
        cashOnHand, advancesIssued, advancesLiquidated,
        advancesOutstanding, accounted, toReplenish,
        categories, dailyRows,
        withReceiptEntries, withoutReceiptEntries
      }
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
// PETTY CASH REQUESTS (PCR) — Pre-approval workflow
// Sheet columns:
//   1: Request_ID    2: Date          3: Purpose        4: Amount
//   5: Request_Type  6: Requested_By  7: Submitted_By   8: Status
//   9: Approved_By  10: Approved_At  11: Released_At   12: Rejection_Note
//  13: Entry_ID     14: Created_At   15: Updated_At
//
// Request_Type  = 'Expense' | 'Cash Advance'
// Requested_By  = employee name (who the cash is for)
// Submitted_By  = cashier email (who created the request)
// ─────────────────────────────────────────────

function savePettyCashRequest(data) {
  // data: { date, purpose, amount, requestedBy, requestType }
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REQUESTS);
    if (!sheet) return { success: false, message: 'Requests sheet not found. Please run setup.' };

    const now      = new Date().toISOString();
    const email    = getUserEmail();
    const id       = generateId('PCR', data.date, sheet);
    const rType    = data.requestType === 'Cash Advance' ? 'Cash Advance' : 'Expense';
    const roleInfo = getUserRole();
    const creatorIsAdmin = roleInfo.success && roleInfo.role === 'Admin';

    // Admin-created requests skip approval — they start as APPROVED immediately
    const initialStatus = creatorIsAdmin ? 'APPROVED' : 'PENDING_APPROVAL';

    sheet.appendRow([
      id,                                // 1  Request_ID
      data.date,                         // 2  Date
      data.purpose || '',                // 3  Purpose
      parseFloat(data.amount) || 0,      // 4  Amount
      rType,                             // 5  Request_Type
      data.requestedBy || '',            // 6  Requested_By (employee name)
      email,                             // 7  Submitted_By (cashier email)
      initialStatus,                     // 8  Status
      creatorIsAdmin ? email : '',       // 9  Approved_By
      creatorIsAdmin ? now   : '',       // 10 Approved_At
      '', '', '',                        // 11-13 Released_At, Rejection_Note, Entry_ID
      now, now                           // 14-15 Created_At, Updated_At
    ]);

    writeAuditLog('REQUEST_CREATED',
      `PCR [${rType}] submitted by ${email}. For: ${data.requestedBy || '—'} | Purpose: ${data.purpose || '—'} | Amount: ₱${parseFloat(data.amount || 0).toFixed(2)}`,
      id, data.date);

    if (creatorIsAdmin) {
      writeAuditLog('REQUEST_APPROVED',
        `Auto-approved: created by Admin (${email}).`,
        id, data.date);
    }

    return { success: true, id };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getPettyCashRequests() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REQUESTS);
    if (!sheet) return { success: true, data: [] };

    const rows  = sheet.getDataRange().getValues();
    const email = getUserEmail();
    const role  = getUserRole();
    const isPrivileged = role.success && (role.role === 'Admin' || role.role === 'Auditor');

    // Build a map of PCR_DETAIL entries grouped by referenceNo for settlement breakdowns
    const settlementMap = {};
    const entrySheet = ss.getSheetByName(SHEETS.ENTRIES);
    if (entrySheet) {
      const eRows = entrySheet.getDataRange().getValues();
      for (let j = 1; j < eRows.length; j++) {
        const e = eRows[j];
        if (!e[0] || e[2] !== 'PCR_DETAIL' || e[10] === 'VOID') continue;
        const ref = e[7];
        if (!ref) continue;
        if (!settlementMap[ref]) settlementMap[ref] = [];
        settlementMap[ref].push({
          category   : e[3],
          description: e[4],
          amount     : parseFloat(e[5]) || 0,
          hasReceipt : e[6] === 'YES'
        });
      }
    }

    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      // Cashiers only see requests they submitted; Admin/Auditor see all
      if (!isPrivileged && r[6] !== email) continue;
      data.push({
        id              : r[0],
        date            : normalizeDate(r[1]),
        purpose         : r[2],
        amount          : parseFloat(r[3]) || 0,
        requestType     : r[4] || 'Expense',  // 'Expense' | 'Cash Advance'
        requestedByName : r[5],   // employee name
        submittedBy     : r[6],   // cashier email
        status          : r[7],
        approvedBy      : r[8],
        approvedAt      : r[9],
        releasedAt      : r[10],
        rejectionNote   : r[11],
        entryId         : r[12],
        createdAt       : r[13],
        settlementItems : r[7] === 'SETTLED' ? (settlementMap[r[0]] || []) : undefined
      });
    }
    data.sort((a, b) => (a.createdAt > b.createdAt ? -1 : 1));
    return { success: true, data };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

function approvePettyCashRequest(requestId) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REQUESTS);
    const rows  = sheet.getDataRange().getValues();
    const email = getUserEmail();
    const now   = new Date().toISOString();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== requestId) continue;
      if (rows[i][7] !== 'PENDING_APPROVAL') return { success: false, message: 'Request is no longer pending approval.' };
      const row = i + 1;
      sheet.getRange(row, 8).setValue('APPROVED');
      sheet.getRange(row, 9).setValue(email);
      sheet.getRange(row, 10).setValue(now);
      sheet.getRange(row, 15).setValue(now);

      writeAuditLog('REQUEST_APPROVED',
        `PCR approved for ${rows[i][5]}. Type: ${rows[i][4]} | Purpose: ${rows[i][2]} | Amount: ₱${parseFloat(rows[i][3] || 0).toFixed(2)}`,
        requestId, normalizeDate(rows[i][1]));

      return { success: true };
    }
    return { success: false, message: 'Request not found.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function rejectPettyCashRequest(data) {
  // data: { requestId, note }
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REQUESTS);
    const rows  = sheet.getDataRange().getValues();
    const email = getUserEmail();
    const now   = new Date().toISOString();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== data.requestId) continue;
      if (rows[i][7] !== 'PENDING_APPROVAL') return { success: false, message: 'Request is no longer pending approval.' };
      const row = i + 1;
      sheet.getRange(row, 8).setValue('REJECTED');
      sheet.getRange(row, 9).setValue(email);
      sheet.getRange(row, 12).setValue(data.note || '');
      sheet.getRange(row, 15).setValue(now);

      writeAuditLog('REQUEST_REJECTED',
        `PCR rejected for ${rows[i][5]}. Purpose: ${rows[i][2]} | Note: ${data.note || '—'}`,
        data.requestId, normalizeDate(rows[i][1]));

      return { success: true };
    }
    return { success: false, message: 'Request not found.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function releasePettyCashRequest(requestId) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REQUESTS);
    const rows  = sheet.getDataRange().getValues();
    const now   = new Date().toISOString();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== requestId) continue;
      if (rows[i][7] !== 'APPROVED') return { success: false, message: 'Request must be approved before releasing.' };

      const releaseDate   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const requestType   = rows[i][4] || 'Expense';   // col 5
      const amount        = parseFloat(rows[i][3]) || 0;
      const purpose       = rows[i][2] || '';
      const requestedFor  = rows[i][5] || '';           // col 6 - employee name
      const approvedBy    = rows[i][8] || '';           // col 9

      // Expense → PCR_ADVANCE (settled via PCR page)
      // Cash Advance → CASH_ADVANCE (liquidated via Entries/Advances page like a manual advance)
      const entryType     = requestType === 'Cash Advance' ? 'CASH_ADVANCE' : 'PCR_ADVANCE';
      const entryCategory = requestType === 'Cash Advance' ? 'Cash Advance' : 'Petty Cash Request';

      const entryResult = saveExpenseEntry({
        date       : releaseDate,
        type       : entryType,
        category   : entryCategory,
        description: purpose,
        amount     : amount,
        hasReceipt : false,
        referenceNo: requestId,
        requestedBy: requestedFor,
        approvedBy : approvedBy
      });
      if (!entryResult.success) return { success: false, message: 'Failed to create entry: ' + entryResult.message };

      const row = i + 1;
      sheet.getRange(row, 8).setValue('RELEASED');
      sheet.getRange(row, 11).setValue(now);
      sheet.getRange(row, 13).setValue(entryResult.id);
      sheet.getRange(row, 15).setValue(now);

      recalculateDailySummary(releaseDate);

      writeAuditLog('REQUEST_RELEASED',
        `PCR [${requestType}] released for ${requestedFor}. ₱${amount.toFixed(2)} disbursed. Entry: ${entryType}`,
        requestId, releaseDate);

      return { success: true, entryId: entryResult.id };
    }
    return { success: false, message: 'Request not found.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function settlePettyCashRequest(data) {
  // data: { requestId, entries: [{category, description, amount, hasReceipt, receipt}], note }
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REQUESTS);
    const rows  = sheet.getDataRange().getValues();
    const now   = new Date().toISOString();

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== data.requestId) continue;
      if (rows[i][7] !== 'RELEASED') return { success: false, message: 'Request must be released before settling.' };

      const requestType  = rows[i][4] || 'Expense';
      if (requestType === 'Cash Advance') {
        return { success: false, message: 'Cash Advance requests are liquidated from the Cash Advances page, not settled here.' };
      }

      const reqAmount    = parseFloat(rows[i][3]) || 0;
      const requestedFor = rows[i][5] || '';   // col 6 - employee name
      const approvedBy   = rows[i][8] || '';   // col 9
      const advEntryId   = rows[i][12];        // col 13 - entry ID
      const entries      = data.entries || [];
      const settleDate   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

      const totalSpent = entries.reduce((s, e) => s + (parseFloat(e.amount) || 0), 0);
      const change     = parseFloat((reqAmount - totalSpent).toFixed(2));

      // 1. Create PCR_DETAIL expense entries for each item (on settlement date)
      for (const entry of entries) {
        const amt = parseFloat(entry.amount) || 0;
        if (amt <= 0) continue;

        const expResult = saveExpenseEntry({
          date       : settleDate,
          type       : 'PCR_DETAIL',
          category   : entry.category   || 'Miscellaneous',
          description: entry.description|| '',
          amount     : amt,
          hasReceipt : !!entry.hasReceipt,
          referenceNo: data.requestId,
          requestedBy: requestedFor,
          approvedBy : approvedBy
        });

        if (expResult.success && entry.hasReceipt && entry.receipt) {
          saveReceiptRecord({
            entryId     : expResult.id,
            date        : settleDate,
            supplierName: entry.receipt.supplierName || '',
            address     : entry.receipt.address      || '',
            tin         : entry.receipt.tin          || '',
            receiptNo   : entry.receipt.receiptNo    || '',
            grossAmount : amt
          });
        }
      }

      // 2. Record change return on settlement date so today's balance reflects cash back
      if (change > 0) {
        saveExpenseEntry({
          date       : settleDate,
          type       : 'CASH_RETURN',
          category   : 'Cash Return',
          description: `Change returned from ${data.requestId}`,
          amount     : change,
          hasReceipt : false,
          referenceNo: data.requestId
        });
      }

      // 3. Mark the PCR_ADVANCE entry as LIQUIDATED
      if (advEntryId) {
        const entrySheet = ss.getSheetByName(SHEETS.ENTRIES);
        const entryRows  = entrySheet.getDataRange().getValues();
        for (let j = 1; j < entryRows.length; j++) {
          if (entryRows[j][0] !== advEntryId) continue;
          // Read the advance's date so we can recalculate that day's summary too
          const advDate = normalizeDate(entryRows[j][1]);
          entrySheet.getRange(j + 1, 11).setValue('LIQUIDATED');
          entrySheet.getRange(j + 1, 13).setValue(now);
          entrySheet.getRange(j + 1, 15).setValue('[SETTLED] ' + (data.note || ''));
          if (advDate && advDate !== settleDate) recalculateDailySummary(advDate);
          break;
        }
      }

      // 4. Mark request as SETTLED
      const row = i + 1;
      sheet.getRange(row, 8).setValue('SETTLED');
      sheet.getRange(row, 12).setValue(data.note || '');
      sheet.getRange(row, 15).setValue(now);

      recalculateDailySummary(settleDate);

      writeAuditLog('REQUEST_SETTLED',
        `PCR settled. Spent: ₱${totalSpent.toFixed(2)} | Change: ₱${change.toFixed(2)} | Items: ${entries.length}`,
        data.requestId, settleDate);

      return { success: true };
    }
    return { success: false, message: 'Request not found.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}