/** --- UTILS --- **/

function saveToSheet(ss, name, headers, newRows) {
  if (newRows.length === 0) return;
  let sheet = ss.getSheetByName(name) || ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  const lastRow = sheet.getLastRow();
  let filteredRows = newRows;

  if (lastRow > 1) {
    // Weather uses Col 1 (FROM) for duplicates, the others use Col 3 (DATE/FROM)
    const checkCol = (name === "Weather") ? 1 : 3;
    const existingData = sheet.getRange(2, checkCol, lastRow - 1, 1).getValues();
    const existingKeys = new Set(existingData.map(r => String(r[0])));
    filteredRows = newRows.filter(r => !existingKeys.has(String(r[checkCol - 1])));
  }

  if (filteredRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, filteredRows.length, filteredRows[0].length).setValues(filteredRows);

    // Sort: Newest at the top (Descending)
    let sortCriteria = [];
    if (name === "Measured Data") {
      sortCriteria = [{ column: 3, ascending: false }, { column: 4, ascending: false }];
    } else {
      const sortCol = (name === "Weather") ? 1 : 3;
      sortCriteria = [{ column: sortCol, ascending: false }];
    }
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort(sortCriteria);
  }
}

function fetchTadoData(dateStr, token, zoneId) {
  const url = `https://my.tado.com/api/v2/homes/${CONFIG.HOME_ID}/zones/${zoneId}/dayReport?date=${dateStr}`;
  const response = UrlFetchApp.fetch(url, { "headers": { "Authorization": "Bearer " + token }, "muteHttpExceptions": true });
  return response.getResponseCode() === 200 ? JSON.parse(response.getContentText()) : null;
}

function getAccessToken() {
  const props = PropertiesService.getScriptProperties();
  let refreshToken = props.getProperty('REFRESH_TOKEN') || CONFIG.INITIAL_REFRESH_TOKEN;
  if (!refreshToken || refreshToken === "YOUR_REFRESH_TOKEN_HERE") throw new Error("Check Refresh Token.");
  const payload = { "client_id": CONFIG.CLIENT_ID, "grant_type": "refresh_token", "refresh_token": refreshToken };
  const response = UrlFetchApp.fetch("https://login.tado.com/oauth2/token", { "method": "post", "payload": payload, "muteHttpExceptions": true });
  const data = JSON.parse(response.getContentText());
  if (data.access_token) {
    if (data.refresh_token) props.setProperty('REFRESH_TOKEN', data.refresh_token);
    return data.access_token;
  }
  throw new Error("Auth Error: " + response.getContentText());
}

function resetHistoryProgress() {
  PropertiesService.getScriptProperties().deleteProperty("LAST_PROCESSED_DATE");
  console.log("✅ History progress has been reset. Next run will start from your current target date.");
}

function formatTadoDate(isoString) {
  if (!isoString || isoString === "") return "";
  const d = new Date(isoString);
  if (isNaN(d.getTime())) return isoString;
  return Utilities.formatDate(d, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");
}

function filterInterval(interval) {
  return interval && interval.from && interval.to;
}

/**
 * --- UNIVERSAL CLEANUP TOOLS WITH DETAILED LOGGING AND CELL REFERENCES ---
 */

function cleanAllSheetsDryRun() {
  runUniversalCleanup(true);
}

function cleanAllSheetsEffective() {
  runUniversalCleanup(false);
}

/**
 * --- UNIVERSAL CLEANUP TOOLS (VERSION WITH DETAILED ROW LOGS) ---
 */

function runUniversalCleanup(isDryRun) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsConfig = [
    // CONFIGURATION FOR THE FIRST CASE: DATE > TIME
    { name: "Measured Data", keyCols: [2, 3], sortCols: [{ col: 3, asc: false }, { col: 4, asc: false }] },
    { name: "Settings", keyCols: [2, 3], sortCols: [{ col: 3, asc: false }, { col: 4, asc: false }] },
    { name: "Call For Heat", keyCols: [2, 3], sortCols: [{ col: 3, asc: false }, { col: 4, asc: false }] },
    { name: "Weather", keyCols: [1, 2], sortCols: [{ col: 2, asc: false }, { col: 3, asc: false }] }
  ];

  sheetsConfig.forEach(conf => {
    const sheet = ss.getSheetByName(conf.name);
    if (!sheet) return;

    const range = sheet.getDataRange();
    const fullData = range.getValues();
    if (fullData.length <= 1) return;

    const headers = fullData[0];
    const dataRows = fullData.slice(1);

    const seenKeys = new Map();
    const uniqueRows = [];
    const duplicatesLog = {};

    dataRows.forEach((row, index) => {
      const currentRowNum = index + 2;

      const keyParts = conf.keyCols.map(idx => {
        let val = row[idx];
        if (val instanceof Date) {
          let format = (idx === 2 || (conf.name === "Weather" && idx === 1)) ? "yyyy-MM-dd" : "HH:mm:ss";
          return Utilities.formatDate(val, CONFIG.TIMEZONE, format);
        }
        return String(val).trim();
      });

      const key = keyParts.join(" | ");

      if (!key || key === " | ") {
        uniqueRows.push(row);
        return;
      }

      if (!seenKeys.has(key)) {
        seenKeys.set(key, currentRowNum);
        uniqueRows.push(row);
      } else {
        if (!duplicatesLog[key]) {
          duplicatesLog[key] = { kept: seenKeys.get(key), removed: [] };
        }
        duplicatesLog[key].removed.push(currentRowNum);
      }
    });

    if (!isDryRun && uniqueRows.length > 0) {
      sheet.clear();
      const finalOutput = [headers, ...uniqueRows];
      sheet.getRange(1, 1, finalOutput.length, headers.length).setValues(finalOutput);

      // APPLY THE CORRECT SORTING
      const sortCriteria = conf.sortCols.map(s => ({ column: s.col + 1, ascending: s.asc }));
      sheet.getRange(2, 1, uniqueRows.length, headers.length).sort(sortCriteria);

      console.log(`✅ [${conf.name}] Sorted by Day and then by Time.`);
    }
  });
}

/**
 * HELPER: Configures a specific date to start the historical download.
 * You can run this directly from the Google Apps Script editor.
 */
function setHistoricalStartDate() {
  // --- CONFIGURE THE DESIRED DATE HERE (YYYY-MM-DD Format) ---
  const desiredDate = "2024-12-26";
  // -----------------------------------------------------------

  const props = PropertiesService.getScriptProperties();
  props.setProperty("LAST_PROCESSED_DATE", desiredDate);

  // We use console.log instead of getUi().alert() to avoid the error in the editor
  console.log("✅ Date configured correctly.");
  console.log("📅 The history will start on: " + desiredDate);
  console.log("💡 Now you can run the 'historicalSync' function.");
}

function getZones(homeId, headers) {
  const url = `https://my.tado.com/api/v2/homes/${homeId}/zones`;
  const options = {
    'method': 'get',
    'headers': headers,
    'muteHttpExceptions': true
  };
  const response = UrlFetchApp.fetch(url, options);

  //console.log("Status Code: " + response.getResponseCode());
  //console.log("Error Details: " + response.getContentText());
  const zones = JSON.parse(response.getContentText());

  const zoneList = Array.isArray(zones) ? zones : (zones.zones || Object.values(zones));
  zoneList.sort((a, b) => a.id - b.id);

  const textLog = zoneList.map(z => `Zone ${z.id} -> ${z.name}`);
  console.log(textLog.join("\n"));

  return zoneList;
}

/**
 * Wipes all existing data for a specific date to prevent duplicates
 */
function clearDataForDate(targetDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ["Measured Data", "Settings", "Call For Heat", "Weather"];

  sheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return; // Skip if sheet doesn't exist yet

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    const dateCol = (name === "Weather") ? 1 : 3;
    const dates = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();

    let deleteBlocks = [];
    let currentBlock = null;

    // Iterate backwards to safely collect row numbers for deletion
    for (let i = dates.length - 1; i >= 0; i--) {
      let val = dates[i][0];
      let rowDateStr = "";

      if (val instanceof Date) {
        rowDateStr = Utilities.formatDate(val, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);
      } else if (val) {
        rowDateStr = String(val).substring(0, 10); // Extract YYYY-MM-DD
      }

      if (rowDateStr === targetDateStr) {
        let rowNum = i + 2;
        if (!currentBlock) {
          currentBlock = { start: rowNum, count: 1 };
        } else {
          // Extend block upwards
          if (rowNum === currentBlock.start - 1) {
            currentBlock.start = rowNum;
            currentBlock.count++;
          } else {
            deleteBlocks.push(currentBlock);
            currentBlock = { start: rowNum, count: 1 };
          }
        }
      } else {
        if (currentBlock) {
          deleteBlocks.push(currentBlock);
          currentBlock = null;
        }
      }
    }
    if (currentBlock) deleteBlocks.push(currentBlock);

    // Execute all identified block deletions for this sheet
    deleteBlocks.forEach(block => {
      sheet.deleteRows(block.start, block.count);
    });
  });
}

function validateConfig() {
  const missing = [];
  if (!CONFIG.HOME_ID) missing.push("HOME_ID");
  if (!CONFIG.INITIAL_REFRESH_TOKEN && !PropertiesService.getScriptProperties().getProperty('REFRESH_TOKEN')) {
    missing.push("INITIAL_REFRESH_TOKEN");
  }

  if (missing.length > 0) {
    throw new Error(`❌ CRITICAL SETUP ERROR: Missing Script Properties: [${missing.join(", ")}]. Please add them via Project Settings.`);
  }
}