// =============================================
// PETTY CASH SYSTEM — Google Apps Script
// v3.0 — BIR-compliant, race-safe IDs, full audit trail
// =============================================

const SPREADSHEET_ID = '12RGOYbXlHz70wtVskB_SNRjO1pRuhN9pdwauqzBEI-E';

const SHEETS = {
  ENTRIES      : 'PettyCash_Entries',
  DENOMINATIONS: 'PettyCash_Denominations',
  SUMMARY      : 'PettyCash_Summary',
  RECEIPTS     : 'PettyCash_Receipts'   // BIR Purchases Journal
};

// ─────────────────────────────────────────────
// WEB APP ENTRY POINT
// ─────────────────────────────────────────────
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Petty Cash System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getUserEmail() {
  try { return Session.getActiveUser().getEmail() || 'unknown'; }
  catch(e) { return 'unknown'; }
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

function generateId(prefix) {
  const ts  = Date.now().toString(36).toUpperCase();
  const rnd = Math.floor(Math.random() * 65536).toString(36).toUpperCase().padStart(3, '0');
  return `${prefix}-${ts}-${rnd}`;
}

/**
 * Compute Philippine VAT breakdown from a gross (VAT-inclusive) amount.
 * VAT rate: 12%
 * Net of VAT (Vatable Sales) = Gross / 1.12
 * VAT Amount = Gross - (Gross / 1.12)
 * Net of VAT removed from output per business requirement.
 */
function computeVAT(grossAmount) {
  const gross      = parseFloat(grossAmount) || 0;
  const vatableSales = gross / 1.12;
  const vatAmount    = gross - vatableSales;
  return {
    grossAmount  : parseFloat(gross.toFixed(2)),
    vatableSales : parseFloat(vatableSales.toFixed(2)),  // Less: VAT line
    vatAmount    : parseFloat(vatAmount.toFixed(2))      // VAT Amount (12%)
    // netOfVat removed
  };
}

function calculateDenomTotal(denominations) {
  return [1000, 500, 200, 100, 50, 20, 10, 5, 1].reduce((sum, bill) => {
    return sum + (parseInt(denominations[String(bill)]) || 0) * bill;
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
      'D_20','D_10','D_5','D_1',
      'Total','Notes','Created_At'
    ]);
    s.setFrozenRows(1);
    formatHeaderRow(s);
    s.getRange('M2:M').setNumberFormat('₱#,##0.00');
  }

  if (!ss.getSheetByName(SHEETS.SUMMARY)) {
    const s = ss.insertSheet(SHEETS.SUMMARY);
    s.appendRow([
      'Summary_ID','Date','Opening_Cash','Cash_Advance',
      'Total_Exp_With_Receipt','Total_Exp_No_Receipt','Total_Expenses',
      'Closing_Cash','Variance','Status','Closed_By','Updated_At'
    ]);
    s.setFrozenRows(1);
    formatHeaderRow(s);
    s.getRange('C2:I').setNumberFormat('₱#,##0.00');
  }

  // ── PettyCash_Receipts (BIR Purchases Journal) ──────
  // Net_Of_VAT column removed — only Gross, Vatable_Sales (Less VAT), VAT_Amount
  // 1:Receipt_ID  2:Entry_ID  3:Date  4:Supplier_Name
  // 5:Address  6:TIN  7:Receipt_No
  // 8:Gross_Amount  9:Vatable_Sales  10:VAT_Amount
  // 11:Created_By  12:Created_At
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

  return { success: true };
}

// ─────────────────────────────────────────────
// SUMMARY RECALCULATION
// ─────────────────────────────────────────────
function recalculateDailySummary(date) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const entryData = ss.getSheetByName(SHEETS.ENTRIES).getDataRange().getValues();
    let totalExp = 0, totalReceipt = 0, totalNoReceipt = 0, cashAdvance = 0;

    for (let i = 1; i < entryData.length; i++) {
      const row   = entryData[i];
      const rDate = normalizeDate(row[1]);
      if (rDate !== date || row[10] === 'DELETED') continue;

      const amt  = parseFloat(row[5]) || 0;
      const type = row[2];

      if (type === 'CASH_ADVANCE') {
        cashAdvance += amt;
      } else if (type === 'CASH_OVER' || type === 'REPLENISHMENT') {
        // excluded from expense totals
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
        existingClosedBy = sumData[i][10] || '';
        break;
      }
    }

    const summaryRow = [
      openingCash, cashAdvance,
      totalReceipt, totalNoReceipt, totalExp,
      closingCash, variance, status,
      existingClosedBy,
      new Date().toISOString()
    ];

    if (targetRow === -1) {
      sumSheet.appendRow([generateId('SUM'), date, ...summaryRow]);
    } else {
      sumSheet.getRange(targetRow, 3, 1, 10).setValues([summaryRow]);
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
    initializeSheets();
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const now   = new Date().toISOString();
    const id    = generateId('EXP');

    sheet.appendRow([
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

      recalculateDailySummary(payload.date);
      if (oldDate && oldDate !== payload.date) recalculateDailySummary(oldDate);

      return { success: true };
    }
    return { success: false, message: 'Entry not found' };
  } catch(e) {
    return { success: false, message: e.toString() };
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
    initializeSheets();
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
// BIR RECEIPTS — Purchases Journal
// Net_Of_VAT column removed from sheet schema
// ─────────────────────────────────────────────
function saveReceiptRecord(data) {
  try {
    initializeSheets();
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.RECEIPTS);
    const vat   = computeVAT(data.grossAmount);
    const id    = generateId('RCP');
    const user  = getUserEmail();
    const now   = new Date().toISOString();

    sheet.appendRow([
      id,                    // Receipt_ID (PK)
      data.entryId  || '',   // Entry_ID (FK)
      data.date,             // Date
      data.supplierName,     // Supplier_Name
      data.address    || '', // Address
      data.tin        || '', // TIN
      data.receiptNo  || '', // Receipt_No / OR No.
      vat.grossAmount,       // Gross_Amount
      vat.vatableSales,      // Vatable_Sales (Less: VAT)
      vat.vatAmount,         // VAT_Amount (12%)
      // Net_Of_VAT removed
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
    initializeSheets();
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
          vatableSales: row[8],  // Less: VAT
          vatAmount   : row[9],  // VAT Amount (12%)
          // netOfVat removed (was row[10])
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
    initializeSheets();
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
        // netOfVat removed
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
    initializeSheets();
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.DENOMINATIONS);

    let denoms = data.denominations || {};
    if (typeof data.breakdown === 'string') {
      try { denoms = JSON.parse(data.breakdown); } catch(e) { denoms = {}; }
    } else if (data.breakdown && typeof data.breakdown === 'object') {
      denoms = data.breakdown;
    }

    const total    = calculateDenomTotal(denoms);
    const prefix   = data.type === 'START' ? 'DEN-OC' : data.type === 'END' ? 'DEN-CC' : 'DEN';
    const recordId = generateId(prefix);

    sheet.appendRow([
      recordId, data.date, data.type,
      denoms['1000']||0, denoms['500']||0, denoms['200']||0,
      denoms['100'] ||0, denoms['50'] ||0, denoms['20'] ||0,
      denoms['10']  ||0, denoms['5']  ||0, denoms['1']  ||0,
      total, data.notes || '',
      new Date().toISOString()
    ]);

    recalculateDailySummary(data.date);
    return { success: true, id: recordId, total };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getDenominationRecords(date) {
  try {
    initializeSheets();
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
          '1000':row[3],'500':row[4],'200':row[5],
          '100' :row[6],'50' :row[7],'20' :row[8],
          '10'  :row[9],'5'  :row[10],'1' :row[11]
        },
        total: row[12], notes: row[13], timestamp: row[14]
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
    initializeSheets();
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
          closingCash        : row[7],
          variance           : row[8],
          status             : row[9],
          closedBy           : row[10],
          updatedAt          : row[11]
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
        expenses:sRow[6], closing:sRow[7], variance:sRow[8], status:sRow[9]
      });
    }

    return { success: true, data: { entries, summaries } };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getDateRange(params) {
  try {
    const startDate = params.from || '';
    const endDate   = params.to   || '';
    initializeSheets();
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const data  = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    const results = [];

    for (let i = 1; i < data.length; i++) {
      const rowDate = normalizeDate(data[i][1]);
      if (rowDate >= startDate && rowDate <= endDate) {
        results.push({
          id:data[i][0], date:rowDate,
          openingCash:data[i][2], cashAdvance:data[i][3],
          totalWithReceipt:data[i][4], totalWithoutReceipt:data[i][5],
          totalExpenses:data[i][6], closingCash:data[i][7],
          variance:data[i][8], status:data[i][9]
        });
      }
    }
    return { success: true, data: results };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

function findUnclosedPastDate(beforeDate) {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const data    = ss.getSheetByName(SHEETS.SUMMARY).getDataRange().getValues();
    const unclosed = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i].length < 10) continue;
      const rDate = normalizeDate(data[i][1]);
      if (rDate < beforeDate && data[i][9] === 'OPEN') unclosed.push(rDate);
    }

    if (!unclosed.length) return { success: true, date: null };
    unclosed.sort();
    return { success: true, date: unclosed[0] };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getCategories() {
  return {
    success: true,
    data: [
      'Office Supplies','Transportation','Meals & Entertainment',
      'Utilities','Repairs & Maintenance','Postage & Courier','Miscellaneous'
    ]
  };
}

function recalculateRange(params) {
  const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  const entryData = ss.getSheetByName(SHEETS.ENTRIES).getDataRange().getValues();
  const dates     = new Set();
  for (let i = 1; i < entryData.length; i++) {
    const d = normalizeDate(entryData[i][1]);
    if (d >= params.from && d <= params.to) dates.add(d);
  }
  const results = [];
  dates.forEach(d => results.push({ date: d, result: recalculateDailySummary(d) }));
  return { success: true, data: results };
}