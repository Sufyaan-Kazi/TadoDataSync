/**
 * TESTING UTILITY: Wipes only Today's data
 */
function wipeTodayForTesting() {
  const today = new Date();
  const dateStr = Utilities.formatDate(today, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);
  removeDateFromSheets(dateStr);
}

/**
 * TESTING UTILITY: Wipes only Yesterday's data
 */
function wipeYesterdayForTesting() {
  const yesterday = new Date(new Date().getTime() - 24 * 60 * 60 * 1000);
  const dateStr = Utilities.formatDate(yesterday, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);
  removeDateFromSheets(dateStr);
}

/**
 * TESTING UTILITY: Wipes all data from the Tado sheets but keeps the headers.
 * Run this from the editor before testing sync functions for a clean slate.
 */
function wipeAllDataForTesting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ["Measured Data", "Settings", "Call For Heat", "Weather"];
  
  sheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return; // Skip if the sheet doesn't exist yet
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // Deletes all rows from row 2 down to the bottom of the data
      sheet.deleteRows(2, lastRow - 1);
    }
  });
  
  console.log("🧹 All test data wiped! Sheets are clean and ready for a fresh sync.");
}

/**
 * Super-fast array rewrite to drop a specific date without using deleteRows()
 */
function removeDateFromSheets(targetDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ["Measured Data", "Settings", "Call For Heat", "Weather"];
  
  sheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    
    const fullData = sheet.getDataRange().getValues();
    if (fullData.length <= 1) return; // Skip if empty or only headers
    
    const headers = fullData[0];
    const rows = fullData.slice(1);
    
    // Weather uses Col A (index 0) for dates, others use Col C (index 2)
    const dateColIdx = (name === "Weather") ? 0 : 2; 
    
    // Filter out the target date
    const keptRows = rows.filter(row => {
      let val = row[dateColIdx];
      let rowDateStr = "";
      
      if (val instanceof Date) {
        rowDateStr = Utilities.formatDate(val, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);
      } else if (val) {
        rowDateStr = String(val).substring(0, 10);
      }
      
      return rowDateStr !== targetDateStr; // Keep the row if the date DOES NOT match
    });
    
    // Only rewrite the sheet if rows were actually removed
    if (keptRows.length < rows.length) {
      sheet.clearContents(); // Instantly wipes data but preserves your formatting/frozen rows
      const finalData = [headers, ...keptRows];
      sheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
    }
  });
  
  console.log(`🧹 Wiped all data for ${targetDateStr}. Ready for a fresh sync!`);
}