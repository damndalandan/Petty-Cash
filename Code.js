// =============================================
// PETTY CASH SYSTEM - Google Apps Script
// =============================================

const SPREADSHEET_ID = '12RGOYbXlHz70wtVskB_SNRjO1pRuhN9pdwauqzBEI-E';

const SHEETS = {
  ENTRIES: 'PettyCash_Entries',
  DENOMINATIONS: 'PettyCash_Denominations',
  SUMMARY: 'PettyCash_Summary'
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Petty Cash System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =============================================
// HELPERS
// =============================================

// Normalize a value that may be a Date object or a date string to YYYY-MM-DD
function normalizeDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return new Date(val.getTime() - (val.getTimezoneOffset() * 60000))
      .toISOString().split('T')[0];
  }
  return String(val).split('T')[0]; // already a string, trim any time part
}

function generateId(prefix) {
  const now = new Date();
  return `${prefix}-${now.getTime().toString(36).toUpperCase()}`;
}

// Denomination IDs include a timestamp so multiple saves on the same day are unique
function generateDenomId(prefix, dateStr) {
  const [y, m, d] = dateStr.split('-');
  const datePart = `${m}${d}${y}`;
  const timePart = new Date().getTime().toString(36).toUpperCase();
  return `${prefix}-${datePart}-${timePart}`;
}

function generateExpenseId(dateStr) {
  try {
    const [y, m, d] = dateStr.split('-');
    const datePart = `${m}${d}${y}`;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const data = sheet.getDataRange().getValues();
    let maxNum = 0;
    const regex = new RegExp(`^exp-(\\d+)-${datePart}$`, 'i');
    for (let i = 1; i < data.length; i++) {
      const match = String(data[i][0]).match(regex);
      if (match) {
        const num = parseInt(match[1]);
        if (num > maxNum) maxNum = num;
      }
    }
    return `exp-${maxNum + 1}-${datePart}`;
  } catch(e) {
    return `EXP-ERR-${Date.now()}`;
  }
}

function calculateDenomTotal(denominations) {
  const bills = [1000, 500, 200, 100, 50, 20, 10, 5, 1];
  let total = 0;
  bills.forEach(bill => {
    total += (parseInt(denominations[String(bill)]) || 0) * bill;
  });
  return total;
}

function formatHeaderRow(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setBackground('#1a1a2e');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
}

// =============================================
// SHEET INITIALIZATION
// =============================================
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let entriesSheet = ss.getSheetByName(SHEETS.ENTRIES);
  if (!entriesSheet) {
    entriesSheet = ss.insertSheet(SHEETS.ENTRIES);
    entriesSheet.appendRow([
      'Entry ID', 'Date', 'Type', 'Category', 'Description',
      'Amount', 'Has Receipt', 'Reference No', 'Requested By',
      'Approved By', 'Status', 'Timestamp'
    ]);
    entriesSheet.setFrozenRows(1);
    formatHeaderRow(entriesSheet);
  }

  let denomSheet = ss.getSheetByName(SHEETS.DENOMINATIONS);
  if (!denomSheet) {
    denomSheet = ss.insertSheet(SHEETS.DENOMINATIONS);
    denomSheet.appendRow([
      'Record ID', 'Date', 'Type',
      'D_1000', 'D_500', 'D_200', 'D_100', 'D_50',
      'D_20', 'D_10', 'D_5', 'D_1',
      'Total', 'Notes', 'Timestamp'
    ]);
    denomSheet.setFrozenRows(1);
    formatHeaderRow(denomSheet);
  }

  let summarySheet = ss.getSheetByName(SHEETS.SUMMARY);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SHEETS.SUMMARY);
    summarySheet.appendRow([
      'Summary ID', 'Date', 'Opening Cash', 'Cash Advance',
      'Total Expenses With Receipt', 'Total Expenses Without Receipt',
      'Total Expenses', 'Closing Cash', 'Variance',
      'Status', 'Closed By', 'Timestamp'
    ]);
    summarySheet.setFrozenRows(1);
    formatHeaderRow(summarySheet);
  }

  return { success: true, message: 'Sheets initialized' };
}

// =============================================
// CENTRALIZED SUMMARY RECALCULATION
// =============================================
function recalculateDailySummary(date) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── Tally entries ──
    const entrySheet = ss.getSheetByName(SHEETS.ENTRIES);
    const entryData  = entrySheet.getDataRange().getValues();

    let totalExp       = 0;
    let totalReceipt   = 0;
    let totalNoReceipt = 0;
    let cashAdvance    = 0;
    let cashOver       = 0; // CASH_OVER entries (overage found during closing)

    for (let i = 1; i < entryData.length; i++) {
      const row   = entryData[i];
      const rDate = normalizeDate(row[1]);
      if (rDate !== date || row[10] === 'DELETED') continue;

      const amt        = parseFloat(row[5]) || 0;
      const type       = row[2];
      const hasReceipt = row[6] === 'YES';

      if (type === 'CASH_ADVANCE') {
        cashAdvance += amt;
      } else if (type === 'CASH_OVER') {
        // Overage is NOT an expense — it actually adds back to the fund.
        // Track separately; does not affect totalExp.
        cashOver += amt;
      } else if (type === 'REPLENISHMENT') {
        // Replenishment adds funds — also tracked separately, not an outflow.
        // (No change to totalExp)
      } else {
        // EXPENSE
        totalExp += amt;
        if (hasReceipt) totalReceipt   += amt;
        else            totalNoReceipt += amt;
      }
    }

    // ── Get cash counts from denominations ──
    const denomSheet = ss.getSheetByName(SHEETS.DENOMINATIONS);
    const denomData  = denomSheet.getDataRange().getValues();

    let openingCash = 0;
    let closingCash = 0;
    let hasClosingRecord = false;

    for (let i = 1; i < denomData.length; i++) {
      const rDate = normalizeDate(denomData[i][1]);
      if (rDate !== date) continue;
      const type  = denomData[i][2];
      const total = parseFloat(denomData[i][12]) || 0;
      if (type === 'START') openingCash = total;
      if (type === 'END')   { closingCash = total; hasClosingRecord = true; }
    }

    // ── Variance ──
    // Expected = Opening - Expenses - Cash Advances (outflows)
    // Overage and Replenishment don't change the expected calculation;
    // they explain WHY closing differs from expected.
    const totalOutflows = totalExp + cashAdvance;
    const expected      = openingCash - totalOutflows;
    const variance      = closingCash - expected;

    const newStatus = hasClosingRecord ? 'CLOSED' : 'OPEN';

    // ── Upsert Summary row ──
    const sumSheet = ss.getSheetByName(SHEETS.SUMMARY);
    const sumData  = sumSheet.getDataRange().getValues();
    let targetRow  = -1;

    for (let i = 1; i < sumData.length; i++) {
      if (normalizeDate(sumData[i][1]) === date) { targetRow = i + 1; break; }
    }

    const now = new Date().toISOString();
    if (targetRow === -1) {
      sumSheet.appendRow([
        generateId('SUM'), date,
        openingCash, cashAdvance, totalReceipt, totalNoReceipt,
        totalExp, closingCash, variance, newStatus, '', now
      ]);
    } else {
      sumSheet.getRange(targetRow, 3, 1, 10).setValues([[
        openingCash, cashAdvance, totalReceipt, totalNoReceipt,
        totalExp, closingCash, variance, newStatus, '', now
      ]]);
    }

    return { success: true };
  } catch(e) {
    console.error(e);
    return { success: false, message: e.toString() };
  }
}

// =============================================
// EXPENSE ENTRY CRUD
// =============================================
function saveExpenseEntry(data) {
  try {
    initializeSheets();
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const id    = generateExpenseId(data.date);

    sheet.appendRow([
      id, data.date, data.type || 'EXPENSE',
      data.category, data.description,
      parseFloat(data.amount),
      data.hasReceipt ? 'YES' : 'NO',
      data.referenceNo || '',
      data.requestedBy || '',
      data.approvedBy  || '',
      'ACTIVE',
      new Date().toISOString()
    ]);

    recalculateDailySummary(data.date);
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ── Cash Overage entry (separate from expense so reports can distinguish) ──
// Stored in the same Entries sheet with type = CASH_OVER
function saveCashOverEntry(data) {
  try {
    initializeSheets();
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const id    = generateExpenseId(data.date);

    sheet.appendRow([
      id, data.date, 'CASH_OVER',
      data.category || 'Miscellaneous',
      data.description || 'Cash Overage',
      parseFloat(data.amount),
      'NO',                       // overages never have receipts
      data.referenceNo || 'OVER-ADJ',
      data.requestedBy || 'System',
      data.approvedBy  || 'Auto',
      'ACTIVE',
      new Date().toISOString()
    ]);

    recalculateDailySummary(data.date);
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ── Replenishment entry — records that a top-up was requested/received ──
function saveReplenishmentEntry(data) {
  try {
    initializeSheets();
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const id    = generateExpenseId(data.date);

    sheet.appendRow([
      id, data.date, 'REPLENISHMENT',
      'Replenishment',
      data.description || 'Fund Replenishment',
      parseFloat(data.amount),
      'NO',
      data.referenceNo || 'REPLENISH-REQ',
      data.requestedBy || 'System',
      data.approvedBy  || '',
      'ACTIVE',
      new Date().toISOString()
    ]);

    // Note: recalculate does NOT add replenishment to expenses —
    // it's tracked separately in the summary for visibility.
    recalculateDailySummary(data.date);
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// FIX: updateExpenseEntry now accepts a SINGLE object {id, ...data}
// because google.script.run only supports one argument.
// The frontend should call: callServer('updateExpenseEntry', { id, ...data })
function updateExpenseEntry(payload) {
  try {
    const entryId = payload.id;
    if (!entryId) return { success: false, message: 'No entry ID provided' };

    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet     = ss.getSheetByName(SHEETS.ENTRIES);
    const dataRange = sheet.getDataRange().getValues();

    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] === entryId) {
        const row     = i + 1;
        const oldDate = normalizeDate(dataRange[i][1]);

        sheet.getRange(row, 2, 1, 11).setValues([[
          payload.date,
          payload.type       || 'EXPENSE',
          payload.category,
          payload.description,
          parseFloat(payload.amount),
          payload.hasReceipt ? 'YES' : 'NO',
          payload.referenceNo  || '',
          payload.requestedBy  || '',
          payload.approvedBy   || '',
          'ACTIVE',
          new Date().toISOString()
        ]]);

        recalculateDailySummary(payload.date);
        if (oldDate && oldDate !== payload.date) recalculateDailySummary(oldDate);

        return { success: true };
      }
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

    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] === entryId) {
        const date = normalizeDate(dataRange[i][1]);
        sheet.getRange(i + 1, 11).setValue('DELETED');
        recalculateDailySummary(date);
        return { success: true };
      }
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
        id: row[0], date: rowDate,
        type: row[2], category: row[3], description: row[4],
        amount: row[5], hasReceipt: row[6] === 'YES',
        referenceNo: row[7], requestedBy: row[8],
        approvedBy: row[9], status: row[10]
      });
    }
    return { success: true, data: entries };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// =============================================
// DENOMINATION RECORDS
// =============================================
function saveDenomination(data) {
  try {
    initializeSheets();
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.DENOMINATIONS);

    const prefix   = data.type === 'START' ? 'DEN-OC' : data.type === 'END' ? 'DEN-CC' : 'DEN';
    const recordId = generateDenomId(prefix, data.date);
    const denoms   = data.denominations || {};
    const total    = calculateDenomTotal(denoms);

    sheet.appendRow([
      recordId, data.date, data.type,
      denoms['1000'] || 0, denoms['500'] || 0, denoms['200'] || 0,
      denoms['100']  || 0, denoms['50']  || 0, denoms['20']  || 0,
      denoms['10']   || 0, denoms['5']   || 0, denoms['1']   || 0,
      total, data.notes || '',
      new Date().toISOString()
    ]);

    recalculateDailySummary(data.date);
    return { success: true, id: recordId, total: total };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getDenominations(date) {
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
        id: row[0], date: rowDate, type: row[2],
        denominations: {
          '1000': row[3], '500': row[4], '200': row[5],
          '100':  row[6], '50':  row[7], '20':  row[8],
          '10':   row[9], '5':   row[10],'1':   row[11]
        },
        total: row[12], notes: row[13], timestamp: row[14]
      });
    }
    return { success: true, data: records };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// Front-end adapter: breakdown comes as JSON string from the form
function saveDenominationRecord(data) {
  const payload = JSON.parse(JSON.stringify(data));
  if (typeof payload.breakdown === 'string') {
    try { payload.denominations = JSON.parse(payload.breakdown); }
    catch(e) { payload.denominations = {}; }
  } else {
    payload.denominations = payload.breakdown || {};
  }
  return saveDenomination(payload);
}

function getDenominationRecords(date) {
  return getDenominations(date);
}

// =============================================
// DAILY SUMMARY
// =============================================
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
          id: row[0], date: rowDate,
          openingCash: row[2], cashAdvance: row[3],
          totalWithReceipt: row[4], totalWithoutReceipt: row[5],
          totalExpenses: row[6], closingCash: row[7],
          variance: row[8], status: row[9], closedBy: row[10]
        }
      };
    }
    return { success: true, data: null };
  } catch(e) {
    return { success: false, message: e.toString(), data: null };
  }
}

// FIX: getDateRange now accepts an object {from, to} and normalizes dates
function getDateRange(params) {
  try {
    const startDate = params.from || params;
    const endDate   = params.to   || arguments[1];
    initializeSheets();
    const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    const summarySheet= ss.getSheetByName(SHEETS.SUMMARY);
    const summaryData = summarySheet.getDataRange().getValues();
    const results     = [];

    for (let i = 1; i < summaryData.length; i++) {
      const row     = summaryData[i];
      const rowDate = normalizeDate(row[1]);
      if (rowDate >= startDate && rowDate <= endDate) {
        results.push({
          id: row[0], date: rowDate,
          openingCash: row[2], cashAdvance: row[3],
          totalWithReceipt: row[4], totalWithoutReceipt: row[5],
          totalExpenses: row[6], closingCash: row[7],
          variance: row[8], status: row[9]
        });
      }
    }
    return { success: true, data: results };
  } catch(e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// =============================================
// REPORT DATA
// =============================================
function generateReportData(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const sheet     = ss.getSheetByName(SHEETS.ENTRIES);
    const dataRange = sheet.getDataRange().getValues();
    const entries   = [];

    for (let i = 1; i < dataRange.length; i++) {
      const row   = dataRange[i];
      if (row.length < 11) continue;
      const rDate = normalizeDate(row[1]);
      if (rDate < data.from || rDate > data.to || row[10] === 'DELETED') continue;

      entries.push({
        id: row[0], date: rDate,
        type: row[2], category: row[3], description: row[4],
        amount: row[5], hasReceipt: row[6] === 'YES',
        referenceNo: row[7], requestedBy: row[8],
        approvedBy: row[9], status: row[10]
      });
    }

    const summaries = [];
    const sumSheet  = ss.getSheetByName(SHEETS.SUMMARY);
    if (sumSheet) {
      const sumData = sumSheet.getDataRange().getValues();
      for (let j = 1; j < sumData.length; j++) {
        const sRow  = sumData[j];
        if (sRow.length < 8) continue;
        const sDate = normalizeDate(sRow[1]);
        if (sDate < data.from || sDate > data.to) continue;
        summaries.push({
          date: sDate,
          opening: sRow[2], cashAdvance: sRow[3],
          expenses: sRow[6], closing: sRow[7],
          variance: sRow[8], status: sRow[9]
        });
      }
    }

    return { success: true, data: { entries, summaries } };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// =============================================
// UNCLOSED DAY CHECK
// =============================================
function findUnclosedPastDate(beforeDate) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SUMMARY);
    const data  = sheet.getDataRange().getValues();
    const unclosed = [];

    for (let i = 1; i < data.length; i++) {
      const row   = data[i];
      if (row.length < 10) continue;
      const rDate = normalizeDate(row[1]);
      if (rDate < beforeDate && row[9] === 'OPEN') unclosed.push(rDate);
    }

    if (unclosed.length === 0) return { success: true, date: null };
    unclosed.sort();
    return { success: true, date: unclosed[0] };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// =============================================
// MISC
// =============================================
function getCategories() {
  return {
    success: true,
    data: [
      'Office Supplies', 'Transportation', 'Meals & Entertainment',
      'Utilities', 'Repairs & Maintenance', 'Postage & Courier',
      'Miscellaneous', 'Cash Advance'
    ]
  };
}