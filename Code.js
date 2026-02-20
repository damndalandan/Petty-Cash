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
// SHEET INITIALIZATION
// =============================================
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // PettyCash_Entries sheet
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

  // PettyCash_Denominations sheet
  let denomSheet = ss.getSheetByName(SHEETS.DENOMINATIONS);
  if (!denomSheet) {
    denomSheet = ss.insertSheet(SHEETS.DENOMINATIONS);
    denomSheet.appendRow([
      'Record ID', 'Date', 'Type', // Type: START, END, ADVANCE
      'D_1000', 'D_500', 'D_200', 'D_100', 'D_50',
      'D_20', 'D_10', 'D_5', 'D_1',
      'Total', 'Notes', 'Timestamp'
    ]);
    denomSheet.setFrozenRows(1);
    formatHeaderRow(denomSheet);
  }

  // PettyCash_Summary sheet
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

  return { success: true, message: 'Sheets initialized successfully' };
}

function formatHeaderRow(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setBackground('#1a1a2e');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
}

// =============================================
// EXPENSE ENTRY CRUD
// =============================================
function saveExpenseEntry(data) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);

    const entryId = generateId('EXP');
    const timestamp = new Date().toISOString();

    sheet.appendRow([
      entryId,
      data.date,
      data.type, // EXPENSE, CASH_ADVANCE
      data.category,
      data.description,
      parseFloat(data.amount),
      data.hasReceipt ? 'YES' : 'NO',
      data.referenceNo || '',
      data.requestedBy || '',
      data.approvedBy || '',
      'ACTIVE',
      timestamp
    ]);

    return { success: true, id: entryId, message: 'Entry saved successfully' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function updateExpenseEntry(entryId, data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const dataRange = sheet.getDataRange().getValues();

    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] === entryId) {
        const row = i + 1;
        sheet.getRange(row, 2).setValue(data.date);
        sheet.getRange(row, 3).setValue(data.type);
        sheet.getRange(row, 4).setValue(data.category);
        sheet.getRange(row, 5).setValue(data.description);
        sheet.getRange(row, 6).setValue(parseFloat(data.amount));
        sheet.getRange(row, 7).setValue(data.hasReceipt ? 'YES' : 'NO');
        sheet.getRange(row, 8).setValue(data.referenceNo || '');
        sheet.getRange(row, 9).setValue(data.requestedBy || '');
        sheet.getRange(row, 10).setValue(data.approvedBy || '');
        sheet.getRange(row, 12).setValue(new Date().toISOString());
        return { success: true, message: 'Entry updated successfully' };
      }
    }
    return { success: false, message: 'Entry not found' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deleteExpenseEntry(entryId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const dataRange = sheet.getDataRange().getValues();

    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] === entryId) {
        sheet.getRange(i + 1, 11).setValue('DELETED');
        return { success: true, message: 'Entry deleted successfully' };
      }
    }
    return { success: false, message: 'Entry not found' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getExpenseEntries(date) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const dataRange = sheet.getDataRange().getValues();

    const entries = [];
    for (let i = 1; i < dataRange.length; i++) {
      const row = dataRange[i];
      if (row[1] === date && row[10] !== 'DELETED') {
        entries.push({
          id: row[0],
          date: row[1],
          type: row[2],
          category: row[3],
          description: row[4],
          amount: row[5],
          hasReceipt: row[6] === 'YES',
          referenceNo: row[7],
          requestedBy: row[8],
          approvedBy: row[9],
          status: row[10]
        });
      }
    }
    return { success: true, data: entries };
  } catch (e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// =============================================
// DENOMINATION RECORDS
// =============================================
function saveDenomination(data) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.DENOMINATIONS);

    const recordId = generateId('DEN');
    const denominations = data.denominations;
    const total = calculateDenomTotal(denominations);

    sheet.appendRow([
      recordId,
      data.date,
      data.type, // START, END, ADVANCE
      denominations['1000'] || 0,
      denominations['500'] || 0,
      denominations['200'] || 0,
      denominations['100'] || 0,
      denominations['50'] || 0,
      denominations['20'] || 0,
      denominations['10'] || 0,
      denominations['5'] || 0,
      denominations['1'] || 0,
      total,
      data.notes || '',
      new Date().toISOString()
    ]);

    return { success: true, id: recordId, total: total };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getDenominations(date) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.DENOMINATIONS);
    const dataRange = sheet.getDataRange().getValues();

    const records = [];
    for (let i = 1; i < dataRange.length; i++) {
      const row = dataRange[i];
      if (row[1] === date) {
        records.push({
          id: row[0],
          date: row[1],
          type: row[2],
          denominations: {
            '1000': row[3], '500': row[4], '200': row[5],
            '100': row[6], '50': row[7], '20': row[8],
            '10': row[9], '5': row[10], '1': row[11]
          },
          total: row[12],
          notes: row[13]
        });
      }
    }
    return { success: true, data: records };
  } catch (e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// =============================================
// DAILY SUMMARY
// =============================================
function saveDailySummary(data) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SUMMARY);

    // Check if summary already exists for this date
    const existing = getDailySummary(data.date);
    if (existing.success && existing.data) {
      // Update existing
      const dataRange = sheet.getDataRange().getValues();
      for (let i = 1; i < dataRange.length; i++) {
        if (dataRange[i][1] === data.date) {
          const row = i + 1;
          sheet.getRange(row, 3).setValue(parseFloat(data.openingCash));
          sheet.getRange(row, 4).setValue(parseFloat(data.cashAdvance));
          sheet.getRange(row, 5).setValue(parseFloat(data.totalWithReceipt));
          sheet.getRange(row, 6).setValue(parseFloat(data.totalWithoutReceipt));
          sheet.getRange(row, 7).setValue(parseFloat(data.totalExpenses));
          sheet.getRange(row, 8).setValue(parseFloat(data.closingCash));
          sheet.getRange(row, 9).setValue(parseFloat(data.variance));
          sheet.getRange(row, 10).setValue(data.status || 'OPEN');
          sheet.getRange(row, 11).setValue(data.closedBy || '');
          sheet.getRange(row, 12).setValue(new Date().toISOString());
          return { success: true, message: 'Summary updated' };
        }
      }
    }

    const summaryId = generateId('SUM');
    sheet.appendRow([
      summaryId,
      data.date,
      parseFloat(data.openingCash) || 0,
      parseFloat(data.cashAdvance) || 0,
      parseFloat(data.totalWithReceipt) || 0,
      parseFloat(data.totalWithoutReceipt) || 0,
      parseFloat(data.totalExpenses) || 0,
      parseFloat(data.closingCash) || 0,
      parseFloat(data.variance) || 0,
      data.status || 'OPEN',
      data.closedBy || '',
      new Date().toISOString()
    ]);

    return { success: true, id: summaryId };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getDailySummary(date) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SUMMARY);
    const dataRange = sheet.getDataRange().getValues();

    for (let i = 1; i < dataRange.length; i++) {
        const row = dataRange[i];
        let rowDate = row[1];
        
        // Normalize Date Object if needed, similar to generateReportData
        if (rowDate instanceof Date) {
            rowDate = new Date(rowDate.getTime() - (rowDate.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
        }
    
        if (rowDate === date) {
        return {
          success: true,
          data: {
            id: row[0], date: rowDate,
            openingCash: row[2], cashAdvance: row[3],
            totalWithReceipt: row[4], totalWithoutReceipt: row[5],
            totalExpenses: row[6], closingCash: row[7],
            variance: row[8], status: row[9],
            closedBy: row[10]
          }
        };
      }
    }
    return { success: true, data: null };
  } catch (e) {
    return { success: false, message: e.toString(), data: null };
  }
}

function getDateRange(startDate, endDate) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const summarySheet = ss.getSheetByName(SHEETS.SUMMARY);
    const summaryData = summarySheet.getDataRange().getValues();

    const results = [];
    for (let i = 1; i < summaryData.length; i++) {
      const row = summaryData[i];
      const rowDate = row[1];
      if (rowDate >= startDate && rowDate <= endDate) {
        results.push({
          id: row[0], date: row[1],
          openingCash: row[2], cashAdvance: row[3],
          totalWithReceipt: row[4], totalWithoutReceipt: row[5],
          totalExpenses: row[6], closingCash: row[7],
          variance: row[8], status: row[9]
        });
      }
    }
    return { success: true, data: results };
  } catch (e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// =============================================
// HELPERS
// =============================================
function generateId(prefix) {
  const now = new Date();
  const timestamp = now.getTime().toString(36).toUpperCase();
  return `${prefix}-${timestamp}`;
}

function calculateDenomTotal(denominations) {
  const bills = [1000, 500, 200, 100, 50, 20, 10, 5, 1];
  let total = 0;
  bills.forEach(bill => {
    total += (parseInt(denominations[bill.toString()]) || 0) * bill;
  });
  return total;
}

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

// =============================================
// FRONTEND ADAPTERS
// =============================================

function saveDenominationRecord(data) {
  var payload = JSON.parse(JSON.stringify(data)); // Clone
  if (typeof payload.breakdown === 'string') {
    try {
      payload.denominations = JSON.parse(payload.breakdown);
    } catch (e) {
      payload.denominations = {};
    }
  } else {
    payload.denominations = payload.breakdown || {};
  }
  return saveDenomination(payload); // Ensure this function exists in your main logic
}

function getDenominationRecords(date) {
  return getDenominations(date); // Ensure this function exists in your main logic
}

function generateReportData(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. Get Entries
    var sheet = ss.getSheetByName(SHEETS.ENTRIES);
    if (!sheet) return { success: false, message: 'Entries sheet not found' };
    
    var dataRange = sheet.getDataRange().getValues();
    var entries = [];
    
    // Skip header
    for (var i = 1; i < dataRange.length; i++) {
        var row = dataRange[i];
        if (row.length < 11) continue; // Safety check
        
        var rDate = row[1];
        // Normalize date to string YYYY-MM-DD if it's a Date object
        if (rDate instanceof Date) {
            // Adjust for timezone offset to ensure correct day
            var localDate = new Date(rDate.getTime() - (rDate.getTimezoneOffset() * 60000));
            rDate = localDate.toISOString().split('T')[0];
        }
        
        if (rDate >= data.from && rDate <= data.to && row[10] !== 'DELETED') {
           entries.push({
              id: row[0],
              date: rDate, // Return the normalized string
              type: row[2],
              category: row[3],
              description: row[4],
              amount: row[5],
              hasReceipt: row[6] === 'YES',
              referenceNo: row[7],
              requestedBy: row[8],
              approvedBy: row[9],
              status: row[10]
           });
        }
    }
    
    // 2. Get Daily Summaries
    var summaries = [];
    var sumSheet = ss.getSheetByName(SHEETS.SUMMARY); // Ensure SHEETS.SUMMARY is defined
    if (sumSheet) {
      var sumData = sumSheet.getDataRange().getValues();
      for(var j = 1; j < sumData.length; j++){
         var sRow = sumData[j];
         if (sRow.length < 8) continue; // Safety check
         
         var sDate = sRow[1];
         if (sDate instanceof Date) {
            var localSDate = new Date(sDate.getTime() - (sDate.getTimezoneOffset() * 60000));
            sDate = localSDate.toISOString().split('T')[0];
         }
         
         if(sDate >= data.from && sDate <= data.to){
           summaries.push({
              date: sDate,
              opening: sRow[2],
              cashAdvance: sRow[3],
              expenses: sRow[6],
              closing: sRow[7],
              variance: sRow[8],
              status: sRow[9]
           });
         }
      }
    }
    
    return { success: true, data: { entries: entries, summaries: summaries } };
    
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}
