# advance-appscript-code-
const WEBHOOK_URL = "";



///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////     DO NOT TOUCH BELOW!!!     ////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



/**
 * function execute it self when sheet is open
 */
function onOpen(e) {
  Logger.log('onOpen started');
  try {
    var userProps = PropertiesService.getUserProperties();
    var triggersCreated = userProps.getProperty('triggersInstalled');
    Logger.log('Trigger installed ? ' + triggersCreated);

    var ui = SpreadsheetApp.getUi();
    var addonMenu = ui.createAddonMenu();
    addonMenu.addItem('Install Triggers', 'createTriggersOnce');

    // Only show "Resend Data" option if triggers are installed
    if (triggersCreated === 'true') {
      addonMenu.addItem('Resend Data (Select Rows)', 'showResendPrompt');
    }
    addonMenu.addToUi();
    setupSheetDetails();
  } catch (e) {
    Logger.log('Error in onOpen: ' + e.toString());
  }
}

function setupSheetDetails() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  PropertiesService.getScriptProperties().setProperty("spreadsheetId", spreadsheet.getId());
  PropertiesService.getScriptProperties().setProperty("sheetName", sheet.getName());
}

/**
 * Helper function to check if a row is empty
 */
function isRowEmpty(row) {
  return row.every(function (cell) {
    return row.every(cell => !cell || cell.toString().trim() === "");
  });
}

/**
 * this function triggerd by onEdit event via triggerd
 */
function onEditChanges(e) {
  var lastChange = PropertiesService.getScriptProperties().getProperty('lastColumnChangeTime');
  var scriptProps = PropertiesService.getScriptProperties();
  var lastChange = scriptProps.getProperty('lastColumnChangeTime');
  if (lastChange) {
    var now = Date.now();
    var delta = now - parseInt(lastChange, 10);

    // If onChange occurred less than 1 second ago, skip this edit
    if (delta < 1000) {
      console.log("Skipping onEdit due to recent onChange.");
      return;
    }
  }

  var endpointUrl = WEBHOOK_URL;
  var leadList = [];


  // Check if the event object is available
  if (e && e.range) {
    var editedRange = e.range;
    var sheet = editedRange.getSheet();

    var rowStart = editedRange.getRow();  // Row that was edited
    var colStart = editedRange.getColumn(); // Column that was edited
    var lastColumn = sheet.getLastColumn();
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();


    // Get the header row
    var headerRow = sheet.getRange(1, 1, 1, lastColumn);
    var headerRowVal = headerRow.getValues()[0]; // Values of the header row
    var currentLastRow = sheet.getLastRow();
    var lastStoredRow = parseInt(scriptProps.getProperty('lastRowCount') || '1', 10);

    // Create map object to send as payload
    var map = {
      sheetId: spreadsheet.getId(),
      gId: SpreadsheetApp.getActiveSpreadsheet().getSheetId(),
      headerList: getHeaderList(headerRowVal), // Send the updated headers
      sheetName: SpreadsheetApp.getActiveSpreadsheet().getName() // Updated to use sheet name rather than spreadsheet name
    };

    // Check if the edited column has a header
    var editedColumnHeader = headerRowVal[colStart - 1]; // Header for the edited column
    if (!editedColumnHeader || editedColumnHeader.trim() === "") {
      // Send updated headers to the webhook
      console.log("Edited column has no header, skipping webhook.");
      return; // Do nothing if the edited column's header is empty
    }

    // If the first row (headers) is edited
    if (rowStart === 1) {
      var editedValue = editedRange.getValue();
      var oldValue = e.oldValue || "";
      if (oldValue !== "") {
        showWarningBelt();
      }
      if (endpointUrl) {
        sendDataToWebhook(endpointUrl, map);
      }
      return;
    }

    var dataStartRow, dataRowCount;
    if (rowStart > lastStoredRow) {
      // New row(s) inserted
      dataStartRow = lastStoredRow + 1;
      dataRowCount = rowStart - lastStoredRow;
    } else {
      // Existing row edited
      dataStartRow = rowStart;
      dataRowCount = editedRange.getNumRows();
    }

    var dataRows = sheet.getRange(dataStartRow, 1, dataRowCount, lastColumn).getValues();

    dataRows.forEach(function (rowValues) {
      var hasData = rowValues.some(cellValue => cellValue !== "");
      if (hasData) {
        rowValues = rowValues.map(cellValue => convertUTCToLocalDateOnSave(cellValue));
        var mapped = mapToColumnLetters(headerRowVal, rowValues);
        if (Object.keys(mapped).length > 0) {
          leadList.push(mapped);
        }
      }
    });


    // Add leads to the map and send if data rows were edited
    if (leadList.length > 0) {
      if (leadList.length === 0) {
        console.log("No valid data rows. Skipping API call.");
        return;
      }
      map["leads"] = leadList;

      if (endpointUrl) {
        sendDataToWebhook(endpointUrl, map);
        updateRowChecksums(sheet, [rowStart]);
      }
    } else {
      console.error('LeadList found empty');
    }
  }
  PropertiesService.getScriptProperties().setProperty('lastRowCount', sheet.getLastRow().toString());
}


/**
 * check cloumn changes by onChange event via triggerd
 */
function onColumnChanges(e) {
  if (changeType === 'INSERT_ROW') {
    console.log("Skipping INSERT_ROW handling here; letting onEditChanges handle multi-row pastes.");
    return;
  }
  var changeType = e.changeType;
  console.log("Change Type: ", changeType);

  // Show warning only for column structure changes
  if (changeType === 'INSERT_COLUMN' || changeType === 'REMOVE_COLUMN') {
    showWarningBelt();
  }

  if (
    changeType === 'INSERT_COLUMN' ||
    changeType === 'REMOVE_COLUMN' ||
    changeType === 'INSERT_ROW'
  ) {
    PropertiesService.getScriptProperties().setProperty('lastColumnChangeTime', Date.now().toString());

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastColumn = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();

    var headerRowVal = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    var map = {
      "sheetId": SpreadsheetApp.getActiveSpreadsheet().getId(),
      "gId": SpreadsheetApp.getActiveSpreadsheet().getSheetId(),
      "sheetName": SpreadsheetApp.getActiveSpreadsheet().getName(),
      "headerList": getHeaderList(headerRowVal)
    };

    var endpointUrl = WEBHOOK_URL;
    if (!endpointUrl) return;

    let leadList = [];

    // If only header exists, send only headers
    if (lastRow <= 1) {
      console.log("Only header exists, sending headers.");
      sendDataToWebhook(endpointUrl, map);
      return;
    }

    // Get previous row count from property
    let previousRowCount = parseInt(PropertiesService.getScriptProperties().getProperty('lastRowCount') || '1');

    // Reset count if all rows were deleted and repasted
    if (lastRow < previousRowCount) {
      previousRowCount = 1;
    }

    if (changeType === 'INSERT_ROW') {
      let addedRowCount = lastRow - previousRowCount;
      let startRow = previousRowCount + 1;
      let rowsToFetch = addedRowCount > 0 ? addedRowCount : 1;

      try {
        var newRows = sheet.getRange(startRow, 1, rowsToFetch, lastColumn).getValues();
        newRows.forEach(function (row) {
          if (!isRowEmpty(row)) {
            row = row.map(cell => convertUTCToLocalDateOnSave(cell));
            var mapped = mapToColumnLetters(headerRowVal, row);
            if (Object.keys(mapped).length > 0) {
              leadList.push(mapped);
            }
          }
        });
      } catch (err) {
        console.error("Error fetching new rows: ", err);
      }

      // Update last row count
      PropertiesService.getScriptProperties().setProperty('lastRowCount', lastRow.toString());
    } else {
      // For column insert/remove: send all rows
      var dataRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
      var dataValues = dataRange.getValues();
      dataValues.forEach(function (row) {
        if (!isRowEmpty(row)) {
          row = row.map(cell => convertUTCToLocalDateOnSave(cell));
          var mapped = mapToColumnLetters(headerRowVal, row);
          if (Object.keys(mapped).length > 0) {
            leadList.push(mapped);
          }
        }
      });

      // Update last row count
      PropertiesService.getScriptProperties().setProperty('lastRowCount', lastRow.toString());
    }

    if (leadList.length > 0) {
      map["leads"] = leadList;
      sendDataToWebhook(endpointUrl, map);
      var newRowNumbers = [];
      for (var r = startRow; r <= startRow + rowsToFetch - 1; r++) {
        newRowNumbers.push(r);
      }
      updateRowChecksums(sheet, newRowNumbers);
    }
  }
}




// Helper function to compare two arrays
function getHeaderList(headers) {
  var headerList = {};
  headers.forEach(function (header, colIndex) {
    if (header) {  // Only add non-empty headers
      var columnLetter = getColumnLetter(colIndex + 1);
      headerList[columnLetter] = header;
    }
  });
  return headerList;
}


/**
 * Helper function for the map key value pair
 */
function mapToColumnLetters(headers, rowValues) {
  var keyValuePairs = {};
  headers.forEach(function (header, colIndex) {
    if (header && rowValues[colIndex]) {  // Only add non-empty headers and values
      var columnLetter = getColumnLetter(colIndex + 1);
      keyValuePairs[columnLetter] = rowValues[colIndex];
    }
  });
  return keyValuePairs;
}


/**
 * Helper function to get cloumn letter
 */
function getColumnLetter(column) {
  var letter = '';
  while (column > 0) {
    var temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Helper function to convert UTC date to LocalDate
 */
function convertUTCToLocalDateOnSave(value) {

  value = convertCustomDateFormat(value);

  if (value instanceof Date && !isNaN(value.getTime())) {
    // Format the date to 'yyyy-MM-dd' or another desired format
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return value; // Return the value if it's not a date
}

/**
 * Helper function to convert date into custom Date Format
 */
function convertCustomDateFormat(value) {
  // Check if the value is a string and matches either DD/MM/YYYY or DD-MM-YYYY
  if (typeof value === "string" && value.match(/^d{2}[/-]d{2}[/-]d{4}$/)) {
    var dateParts = value.split(/[/-]/); // Split on either '/' or '-'
    var day = parseInt(dateParts[0], 10);
    var month = parseInt(dateParts[1], 10) - 1; // Month is zero-based in JavaScript Date
    var year = parseInt(dateParts[2], 10);


    // Validate day, month, and year explicitly
    if (month >= 0 && month <= 11 && day >= 1 && day <= daysInMonth(year, month)) {
      // Create a new Date object if valid
      return new Date(year, month, day);
    }
  }

  // Return the original value if it's not a recognized or valid date format
  return value;
}

/**
 * Helper function to convert days in month
 */
function daysInMonth(year, month) {
  // Return the number of days in a given month and year
  return new Date(year, month + 1, 0).getDate();
}

/**
 * Function to send data to webhook url
 */
function sendDataToWebhook(endpointUrl, map) {
  if (endpointUrl) {
    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(map)
    };
    try {
      console.log(options);
      UrlFetchApp.fetch(endpointUrl, options);
    } catch (error) {
      console.error('Failed to send data to API: ' + error);
    }
  }
}

/**
 * this function loads warning modal using html service
 */
function showWarningBelt() {
  var htmlContent = String.raw`
     <div id="warning-belt">
    <!-- Custom styled title -->
    <h2 style="font-size: 14px; color: red; margin: 2; font-family: Karla;"> ⚠️ Header value changed! Make sure it's
      also updated in Callyzer.</h2>
  </div>
  `;

  var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(60);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}



/**
 * create trigger programatically by calling this function 
 */
function createTriggersOnce() {
  // Check if already created
  var userProps = PropertiesService.getUserProperties();
  if (userProps.getProperty('triggersInstalled') === 'true') {
    SpreadsheetApp.getUi().alert("Triggers have already been installed");
    return;
  }

  // Avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  var sheet = SpreadsheetApp.getActive();
  var installedFunctions = triggers.map(t => t.getHandlerFunction());

  if (!installedFunctions.includes('onColumnChanges')) {
    ScriptApp.newTrigger('onColumnChanges').forSpreadsheet(sheet).onChange().create();
  }
  if (!installedFunctions.includes('onEditChanges')) {
    ScriptApp.newTrigger('onEditChanges').forSpreadsheet(sheet).onEdit().create();
  }
  if (!installedFunctions.includes('checkAndSendHeadersOnLoad')) {
    ScriptApp.newTrigger('checkAndSendHeadersOnLoad').forSpreadsheet(sheet).onOpen().create();
  }

  // Add time-driven trigger (every hour)
  var triggers = ScriptApp.getProjectTriggers();
  var timeDrivenExists = triggers.some(trigger =>
    trigger.getHandlerFunction() === 'checkDataBasedOnTime' &&
    trigger.getEventType() === ScriptApp.EventType.CLOCK
  );
  if (!timeDrivenExists) {
    ScriptApp.newTrigger('checkDataBasedOnTime')
      .timeBased()
      .everyMinutes(10)
      .create();
  }


  // Set flag so menu doesn't show again
  userProps.setProperty('triggersInstalled', 'true');

  // Optional: show alert to confirm
  SpreadsheetApp.getUi().alert("Triggers installed successfully!");
  checkAndSendHeadersOnLoad()
}

/**
 * function reset trigger property
 */
function resetTriggerMenu() {
  var userProps = PropertiesService.getUserProperties();
  userProps.deleteProperty('triggersInstalled');
  SpreadsheetApp.getUi().alert("Trigger menu has been reset. Reload the sheet to see the menu again.");
}

/**
 * setup sheet details on loead
 */
function setupSheetDetails() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  PropertiesService.getScriptProperties().setProperty("spreadsheetId", spreadsheet.getId());
  PropertiesService.getScriptProperties().setProperty("sheetName", sheet.getName());
}

/**
 * check and send headers to webhook on sheet load
 * 
 */
function checkAndSendHeadersOnLoad() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var scriptProps = PropertiesService.getScriptProperties();

    var lastColumn = sheet.getLastColumn();
    var currentLastRow = sheet.getLastRow();
    var headerRowVal = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    var hasHeaderData = headerRowVal.some(cell => cell && cell.toString().trim() !== "");
    if (!hasHeaderData) {
      Logger.log("First row is empty. Skipping header and lead data send.");
      return;
    }

    // Check for header changes
    var currentHeaderString = JSON.stringify(headerRowVal);
    var previousHeaderString = scriptProps.getProperty('lastHeaderString');

    var headerChanged = currentHeaderString !== previousHeaderString;

    // Check for new row data
    var lastStoredRow = parseInt(scriptProps.getProperty('lastRowCount') || '1', 10); // Default to 1
    var newRowCount = currentLastRow - lastStoredRow;

    var leadList = [];
    if (newRowCount > 0) {
      var newRows = sheet.getRange(lastStoredRow + 1, 1, newRowCount, lastColumn).getValues();

      newRows.forEach(function (rowValues) {
        var hasData = rowValues.some(cellValue => cellValue !== "");
        if (hasData) {
          rowValues = rowValues.map(cellValue => convertUTCToLocalDateOnSave(cellValue));
          var mapped = mapToColumnLetters(headerRowVal, rowValues);
          if (Object.keys(mapped).length > 0) {
            leadList.push(mapped);
          }
        }
      });
    }

    // Only send if headers changed or new leads found
    if (headerChanged || leadList.length > 0) {
      var map = {
        sheetId: spreadsheet.getId(),
        gId: sheet.getSheetId(),
        sheetName: SpreadsheetApp.getActiveSpreadsheet().getName(),
        headerList: getHeaderList(headerRowVal)
      };

      if (leadList.length > 0) {
        map.leads = leadList;
      }

      if (typeof WEBHOOK_URL !== 'undefined') {
        sendDataToWebhook(WEBHOOK_URL, map);
        Logger.log("Header and/or new lead data sent on sheet load.");
        var newRowNumbers = [];
        for (var r = lastStoredRow + 1; r <= currentLastRow; r++) {
          newRowNumbers.push(r);
        }
        updateRowChecksums(sheet, newRowNumbers);
      } else {
        Logger.log("WEBHOOK_URL not defined.");
      }

      // Update properties
      if (headerChanged) {
        scriptProps.setProperty('lastHeaderString', currentHeaderString);
      }
      if (newRowCount > 0 && leadList.length > 0) {
        scriptProps.setProperty('lastRowCount', currentLastRow.toString());
      }
    } else {
      Logger.log("No header changes or new rows found. Skipping send.");
    }
  } catch (e) {
    Logger.log("Error in checkAndSendHeadersOnLoad: " + e.toString());
  }
}


/** 
 * resend popup in UI with validation
 */
function showResendPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Resend Data", "Enter row number or range (e.g. 2 or 2-5):", ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const input = response.getResponseText().trim();

    // Match either a single number (e.g., 3) or a range (e.g., 2-5)
    const singleMatch = input.match(/^(\d+)$/);
    const rangeMatch = input.match(/^(\d+)\s*-\s*(\d+)$/);

    if (singleMatch) {
      const row = parseInt(singleMatch[1], 10);
      resendRowsToWebhook(row, row);
    } else if (rangeMatch) {
      const fromRow = parseInt(rangeMatch[1], 10);
      const toRow = parseInt(rangeMatch[2], 10);

      if (fromRow <= toRow) {
        resendRowsToWebhook(fromRow, toRow);
      } else {
        ui.alert("Invalid range. Start row should be less than or equal to end row.");
      }
    } else {
      ui.alert("Invalid input. Please enter a single row like '3' or a range like '2-5'.");
    }
  }
}


/**
 * this function resend data from google sheet by providing range
 */
function resendRowsToWebhook(startRow, endRow) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const endpointUrl = WEBHOOK_URL;
  const lastColumn = sheet.getLastColumn();
  const headerRowVal = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const leadList = [];

  for (let row = startRow; row <= endRow; row++) {
    const rowValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
    const hasData = rowValues.some(cellValue => cellValue !== "");

    if (hasData) {
      const localDateRow = rowValues.map(cellValue => convertUTCToLocalDateOnSave(cellValue));
      const mapped = mapToColumnLetters(headerRowVal, localDateRow);
      if (Object.keys(mapped).length > 0) {
        leadList.push(mapped);
      }
    }
  }

  if (leadList.length > 0) {
    const payload = {
      sheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
      gId: sheet.getSheetId(),
      headerList: getHeaderList(headerRowVal),
      sheetName: SpreadsheetApp.getActiveSpreadsheet().getName(),
      leads: leadList
    };

    sendDataToWebhook(endpointUrl, payload);
    var resendRows = [];
    for (let row = startRow; row <= endRow; row++) {
      resendRows.push(row);
    }
    updateRowChecksums(sheet, resendRows);
  } else {
    Logger.log("No valid data found in specified row range.");
    SpreadsheetApp.getUi().alert("No valid data found in specified row range.");
  }
}

function createOneTimeTestTrigger() {
  // Create a trigger that fires in 2 minutes, only once
  ScriptApp.newTrigger("testModifyAndAddRow")
    .timeBased()
    .after(2 * 60 * 1000) // 2 minutes from now
    .create();
}

function testModifyAndAddRow() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = props.getProperty("spreadsheetId");
  var sheetName = props.getProperty("sheetName");

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  // --- Modify a random row between 2 and 40 ---
  var randomRow = Math.floor(Math.random() * (40 - 2 + 1)) + 2;
  var randomCol = Math.floor(Math.random() * lastCol) + 1;
  var newValue = "Modified-" + new Date().toLocaleTimeString();

  sheet.getRange(randomRow, randomCol).setValue(newValue);
  Logger.log("Modified Row " + randomRow + " Col " + randomCol + " with " + newValue);

  // --- Add a new row at the bottom ---
  var newRowValues = [
    "Test", "Lead", "test@example.com", "1234567890", "New",
    "123 Fake Street", "Apt. 101", "Test City", "", "Test State", "99999", "qa test (+91-9999999999)"
  ];
  sheet.appendRow(newRowValues);
  Logger.log("Appended new test row at bottom.");

  // --- Clean up: delete this trigger after running once ---
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "testModifyAndAddRow") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function checkDataBasedOnTime() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var props = PropertiesService.getScriptProperties();
    var spreadsheetId = props.getProperty("spreadsheetId") || ss.getId();
    var sheetName = props.getProperty("sheetName") || SpreadsheetApp.getActiveSpreadsheet().getName();
    var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
    if (!sheet) return;

    var checksumSheet = getOrCreateChecksumSheet(ss);

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < 2) {
      Logger.log("No data rows.");
      return;
    }

    var headerRowVal = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    var oldChecksums = getStoredChecksums(checksumSheet);
    var newChecksums = {};
    var modifiedRows = [];

    // --- RESET DETECTION ---
    var resetDetected = false;
    if (Object.keys(oldChecksums).length > 0) {
      if (data.length === 0) {
        resetDetected = true; // sheet cleared
      } else {
        var changedCount = 0;
        data.forEach(function (row, idx) {
          var rowNum = idx + 2;
          var checksum = hashRow(row);
          if (oldChecksums[rowNum] && oldChecksums[rowNum] !== checksum) {
            changedCount++;
          }
        });
        if (changedCount > data.length * 0.8) {
          resetDetected = true; // bulk overwrite
        }
      }
    }

    if (resetDetected) {
      Logger.log("⚠️ Reset detected — sending all rows.");
      modifiedRows = data
        .map(function (row) {
          var hasData = row.some(cell => cell !== "" && cell !== null);
          if (!hasData) return null;
          row = row.map(cellValue => convertUTCToLocalDateOnSave(cellValue));
          return mapToColumnLetters(headerRowVal, row);
        })
        .filter(mapped => mapped && Object.keys(mapped).length > 0);
    } else {
      // --- NORMAL CHECK FOR NEW/MODIFIED ROWS ---
      data.forEach(function (row, idx) {
        var rowNum = idx + 2;
        var checksum = hashRow(row);
        newChecksums[rowNum] = checksum;

        var hasData = row.some(cell => cell !== "" && cell !== null);

        if (hasData && oldChecksums[rowNum] !== checksum) {
          row = row.map(cellValue => convertUTCToLocalDateOnSave(cellValue));
          var mapped = mapToColumnLetters(headerRowVal, row);
          if (Object.keys(mapped).length > 0) {
            modifiedRows.push(mapped);
            Logger.log("Row " + rowNum + " changed.");
          }
        }
      });
    }

    if (modifiedRows.length > 0) {
      var map = {
        sheetId: ss.getId(),
        gId: sheet.getSheetId(),
        sheetName: SpreadsheetApp.getActiveSpreadsheet().getName(),
        headerList: getHeaderList(headerRowVal),
        leads: modifiedRows
      };

      if (typeof WEBHOOK_URL !== 'undefined') {
        sendDataToWebhook(WEBHOOK_URL, map);
        Logger.log("✅ Sent " + modifiedRows.length + " leads.");
      }
    } else {
      Logger.log("No new/modified rows.");
    }

    // ✅ Always update checksums (after detecting changes)
    data.forEach(function (row, idx) {
      var rowNum = idx + 2;
      newChecksums[rowNum] = hashRow(row);
    });
    writeChecksums(checksumSheet, newChecksums);

  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}

function getOrCreateChecksumSheet(ss) {
  var checksumSheet = ss.getSheetByName("_RowChecksums");
  if (!checksumSheet) {
    checksumSheet = ss.insertSheet("_RowChecksums");
    checksumSheet.hideSheet();
    checksumSheet.appendRow(["RowNum", "Checksum"]);
  }
  return checksumSheet;
}

function getStoredChecksums(sheet) {
  var values = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    map[values[i][0]] = values[i][1];
  }
  return map;
}

function writeChecksums(sheet, newStored) {
  sheet.clearContents();
  sheet.appendRow(["RowNum", "Checksum"]);
  for (var rowNum in newStored) {
    sheet.appendRow([rowNum, newStored[rowNum]]);
  }
}

function hashRow(row) {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, row.join("|"))
  );
}


////////////// testing function which add one row at last of data and one rondom row modified //////////

function createOneTimeTestTrigger() {
  // Create a trigger that fires in 2 minutes, only once
  ScriptApp.newTrigger("testModifyAndAddRow")
    .timeBased()
    .after(2 * 60 * 1000) // 2 minutes from now
    .create();
}

function testModifyAndAddRow() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = props.getProperty("spreadsheetId");
  var sheetName = props.getProperty("sheetName");

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  // --- Modify a random row between 2 and 40 ---
  var randomRow = Math.floor(Math.random() * (40 - 2 + 1)) + 2;
  var randomCol = Math.floor(Math.random() * lastCol) + 1;
  var newValue = "Modified-" + new Date().toLocaleTimeString();

  sheet.getRange(randomRow, randomCol).setValue(newValue);
  Logger.log("Modified Row " + randomRow + " Col " + randomCol + " with " + newValue);

  // --- Add a new row at the bottom ---
  var newRowValues = [
    "Test", "Lead", "test@example.com", "1234567890", "New",
    "123 Fake Street", "Apt. 101", "Test City", "", "Test State", "99999", "qa test (+91-9999999999)"
  ];
  sheet.appendRow(newRowValues);
  Logger.log("Appended new test row at bottom.");

  // --- Clean up: delete this trigger after running once ---
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "testModifyAndAddRow") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


function updateRowChecksums(sheet, rowNumbers) {
  var ss = sheet.getParent();
  var checksumSheet = getOrCreateChecksumSheet(ss);
  var stored = getStoredChecksums(checksumSheet);

  var lastCol = sheet.getLastColumn();
  rowNumbers.forEach(function(rowNum) {
    if (rowNum <= 1) return; // skip header

    var rowValues = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var checksum = hashRow(rowValues);
    stored[rowNum] = checksum;
  });

  // Write back only once
  writeChecksums(checksumSheet, stored);
}
