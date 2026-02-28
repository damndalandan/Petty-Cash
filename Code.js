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

  return { success: true, message: 'Sheets initialized successfully' };
}

function formatHeaderRow(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setBackground('#1a1a2e');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
}

// =============================================
// CENTRALIZED SUMMARY LOGIC
// =============================================
function recalculateDailySummary(date) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    const entrySheet = ss.getSheetByName(SHEETS.ENTRIES);
    const entryData = entrySheet.getDataRange().getValues();
    
    let totalExp = 0;
    let totalReceipt = 0;
    let totalNoReceipt = 0;
    let cashAdvance = 0;
    
    for (let i = 1; i < entryData.length; i++) {
        let rDate = entryData[i][1];
        if (rDate instanceof Date) {
             rDate = new Date(rDate.getTime() - (rDate.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
        }
        
        if (rDate === date && entryData[i][10] !== 'DELETED') {
            const amt = parseFloat(entryData[i][5]) || 0;
            const type = entryData[i][2];
            const hasReceipt = entryData[i][6] === 'YES';
            
            if (type === 'CASH_ADVANCE') {
                cashAdvance += amt;
            } else {
                totalExp += amt;
                if(hasReceipt) totalReceipt += amt;
                else totalNoReceipt += amt;
            }
        }
    }
    
    const denomSheet = ss.getSheetByName(SHEETS.DENOMINATIONS);
    const denomData = denomSheet.getDataRange().getValues();
    
    let openingCash = 0;
    let closingCash = 0;
    
    for (let i = 1; i < denomData.length; i++) {
        let rDate = denomData[i][1];
        if (rDate instanceof Date) {
             rDate = new Date(rDate.getTime() - (rDate.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
        }
        
        if (rDate === date) {
            const type = denomData[i][2];
            const total = parseFloat(denomData[i][12]) || 0;
            if (type === 'START') openingCash = total;
            if (type === 'END') closingCash = total;
        }
    }
    
    const totalOutflows = totalExp + cashAdvance; 
    const expected = openingCash - totalOutflows; 
    const variance = closingCash - expected; 
    
    let newStatus = 'OPEN';
    for (let i = 1; i < denomData.length; i++) {
        let rDate = denomData[i][1];
        if (rDate instanceof Date) {
             rDate = new Date(rDate.getTime() - (rDate.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
        }
        if (rDate === date && denomData[i][2] === 'END') {
            newStatus = 'CLOSED';
            break;
        }
    }

    const sumSheet = ss.getSheetByName(SHEETS.SUMMARY);
    const sumData = sumSheet.getDataRange().getValues();
    let targetRow = -1;
    
    for (let i = 1; i < sumData.length; i++) {
        let rDate = sumData[i][1];
         if (rDate instanceof Date) {
             rDate = new Date(rDate.getTime() - (rDate.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
        }
        if (rDate === date) {
            targetRow = i + 1;
            break;
        }
    }

    if (targetRow === -1) {
        const summaryId = generateId('SUM');
        sumSheet.appendRow([
            summaryId, date, openingCash, cashAdvance,
            totalReceipt, totalNoReceipt, totalExp, 
            closingCash, variance, newStatus, '',
            new Date().toISOString()
        ]);
    } else {
        sumSheet.getRange(targetRow, 3).setValue(openingCash);
        sumSheet.getRange(targetRow, 4).setValue(cashAdvance);
        sumSheet.getRange(targetRow, 5).setValue(totalReceipt);
        sumSheet.getRange(targetRow, 6).setValue(totalNoReceipt);
        sumSheet.getRange(targetRow, 7).setValue(totalExp);
        sumSheet.getRange(targetRow, 8).setValue(closingCash);
        sumSheet.getRange(targetRow, 9).setValue(variance);
        sumSheet.getRange(targetRow, 10).setValue(newStatus);
        sumSheet.getRange(targetRow, 12).setValue(new Date().toISOString());
    }
    
    return { success: true };
    
  } catch (e) {
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
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);

    const entryId = generateExpenseId(data.date);
    const timestamp = new Date().toISOString();

    sheet.appendRow([
      entryId,
      data.date,
      data.type,
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
    
    recalculateDailySummary(data.date);

    return { success: true, id: entryId, message: 'Entry saved successfully' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// FIX: updateExpenseEntry is now a proper top-level callable function
// (previously it existed but wasn't wired up — the frontend was calling saveExpenseEntry
// which created duplicate rows instead of updating the existing one)
function updateExpenseEntry(entryId, data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ENTRIES);
    const dataRange = sheet.getDataRange().getValues();

    for (let i = 1; i < dataRange.length; i++) {
        if (dataRange[i][0] === entryId) {
             const row = i + 1;
             
             // Capture old date so we can recalc it too if date changed
             let oldDate = dataRange[i][1];
             if (oldDate instanceof Date) {
               oldDate = new Date(oldDate.getTime() - (oldDate.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
             }
             
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
             
             // Recalc new date's summary
             recalculateDailySummary(data.date);
             
             // Also recalc old date if date was changed
             if (oldDate && oldDate !== data.date) {
               recalculateDailySummary(oldDate);
             }
             
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
                 const row = i + 1;
                 const date = dataRange[i][1];
                 
                 sheet.getRange(row, 11).setValue('DELETED');
                 
                 let dStr = date;
                 if (date instanceof Date) dStr = new Date(date.getTime() - (date.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
                 
                 recalculateDailySummary(dStr);
                 
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
      let rowDate = row[1];

      if (rowDate instanceof Date) {
        rowDate = new Date(rowDate.getTime() - (rowDate.getTimezoneOffset() * 60000))
                  .toISOString().split('T')[0];
      }

      if (rowDate === date && row[10] !== 'DELETED') {
        entries.push({
          id: row[0],
          date: rowDate,
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

    let prefix = 'DEN';
    if (data.type === 'START') prefix = 'DEN-OC';
    else if (data.type === 'END') prefix = 'DEN-CC';
    
    // FIX: generateDenomId now appends a timestamp suffix so multiple saves
    // on the same day for the same type get unique IDs instead of colliding.
    const recordId = generateDenomId(prefix, data.date);
    const denominations = data.denominations;
    const total = calculateDenomTotal(denominations);

    sheet.appendRow([
      recordId,
      data.date,
      data.type,
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
    
    recalculateDailySummary(data.date);

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
        let rowDate = row[1];
        
        if (rowDate instanceof Date) {
            rowDate = new Date(rowDate.getTime() - (rowDate.getTimezoneOffset() * 60000))
                      .toISOString().split('T')[0];
        }

      if (rowDate === date) {
        records.push({
          id: row[0],
          date: rowDate,
          type: row[2],
          denominations: {
            '1000': row[3], '500': row[4], '200': row[5],
            '100': row[6], '50': row[7], '20': row[8],
            '10': row[9], '5': row[10], '1': row[11]
          },
          total: row[12],
          notes: row[13],
          timestamp: row[14] // pass timestamp to frontend for display
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

    const existing = getDailySummary(data.date);
    if (existing.success && existing.data) {
      const dataRange = sheet.getDataRange().getValues();
      for (let i = 1; i < dataRange.length; i++) {
        let rowDate = dataRange[i][1];
        if (rowDate instanceof Date) {
          rowDate = new Date(rowDate.getTime() - (rowDate.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
        }
        if (rowDate === data.date) {
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

// FIX: getDateRange now normalizes Date objects to strings before comparing,
// matching the same pattern used throughout the rest of the codebase.
// Previously it did raw comparison of Date objects against strings which always silently returned 0 rows.
function getDateRange(params) {
  try {
    // Support both getDateRange({from, to}) object and legacy positional args
    const startDate = (params && params.from) ? params.from : params;
    const endDate = (params && params.to) ? params.to : arguments[1];

    initializeSheets();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const summarySheet = ss.getSheetByName(SHEETS.SUMMARY);
    const summaryData = summarySheet.getDataRange().getValues();

    const results = [];
    for (let i = 1; i < summaryData.length; i++) {
      const row = summaryData[i];
      // FIX: normalize date before string comparison
      let rowDate = row[1];
      if (rowDate instanceof Date) {
        rowDate = new Date(rowDate.getTime() - (rowDate.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
      }

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
  } catch (e) {
    return { success: false, message: e.toString(), data: [] };
  }
}

// =============================================
// HELPERS
// =============================================
function generateId(prefix) {
  const now = new Date();
  return `${prefix}-${now.getTime().toString(36).toUpperCase()}`;
}

// FIX: Denomination IDs now include a millisecond timestamp to guarantee uniqueness.
// Previously DEN-OC-MMDDYYYY was generated identically for every save on the same day,
// meaning only the first was truly unique — subsequent counts silently used the same ID.
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
     
     for(let i=1; i<data.length; i++) {
        const id = data[i][0];
        const match = id.match(regex);
        if(match) {
           const num = parseInt(match[1]);
           if(num > maxNum) maxNum = num;
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
  var payload = JSON.parse(JSON.stringify(data));
  if (typeof payload.breakdown === 'string') {
    try {
      payload.denominations = JSON.parse(payload.breakdown);
    } catch (e) {
      payload.denominations = {};
    }
  } else {
    payload.denominations = payload.breakdown || {};
  }
  return saveDenomination(payload);
}

function getDenominationRecords(date) {
  return getDenominations(date);
}

function generateReportData(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    var sheet = ss.getSheetByName(SHEETS.ENTRIES);
    if (!sheet) return { success: false, message: 'Entries sheet not found' };
    
    var dataRange = sheet.getDataRange().getValues();
    var entries = [];
    
    for (var i = 1; i < dataRange.length; i++) {
        var row = dataRange[i];
        if (row.length < 11) continue;
        
        var rDate = row[1];
        if (rDate instanceof Date) {
            var localDate = new Date(rDate.getTime() - (rDate.getTimezoneOffset() * 60000));
            rDate = localDate.toISOString().split('T')[0];
        }
        
        if (rDate >= data.from && rDate <= data.to && row[10] !== 'DELETED') {
           entries.push({
              id: row[0],
              date: rDate,
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
    
    var summaries = [];
    var sumSheet = ss.getSheetByName(SHEETS.SUMMARY);
    if (sumSheet) {
      var sumData = sumSheet.getDataRange().getValues();
      for(var j = 1; j < sumData.length; j++){
         var sRow = sumData[j];
         if (sRow.length < 8) continue;
         
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

function findUnclosedPastDate(beforeDate) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.SUMMARY);
    const data = sheet.getDataRange().getValues();
    
    let unclosed = [];
    
    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (row.length < 10) continue;
        
        var rDate = row[1];
        if (rDate instanceof Date) {
            rDate = new Date(rDate.getTime() - (rDate.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
        }
        
        if (rDate < beforeDate && row[9] === 'OPEN') {
            unclosed.push(rDate);
        }
    }
    
    if (unclosed.length === 0) return { success: true, date: null };
    
    unclosed.sort();
    return { success: true, date: unclosed[0] }; 
    
  } catch(e) {
     return { success: false, message: e.toString() };
  }
}