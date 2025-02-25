function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Custom')
    menu.addItem('Go to MASTER', 'goToMasterSheet')
    menu.addItem('Go to a Sheet', 'showGoToSheetDialog')
    menu.addSeparator();
    menu.addItem('Update Translation Status', 'openTranslationDialog') 
    menu.addToUi();
}

function goToMasterSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("MASTER");
  
  if (masterSheet) {
    ss.setActiveSheet(masterSheet);
  } else {
    SpreadsheetApp.getUi().alert("The sheet titled 'MASTER' was not found.");
  }
}

function showGoToSheetDialog() {
  var html = HtmlService.createHtmlOutputFromFile('SheetDialog')
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showDialog(html);
}

function goToSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert("The sheet titled '" + sheetName + "' was not found.");
  }
}

function updateReleaseDates() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MASTER");
  if (!sheet) {
    Logger.log("Sheet 'MASTER' not found.");
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  Logger.log("Last row of the sheet: " + lastRow);
  Logger.log("Last column of the sheet: " + lastColumn);

  // Define the maximum row to process (79) to exclude rows 80 and above
  var maxDataRow = 79; // Setting maximum processing row to 79
  var effectiveLastRow = Math.min(lastRow, maxDataRow); // Ensures we don't exceed maxDataRow

  // Get today's date (clear time for comparison)
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  // Fetch headers once (assuming headers are in row 24)
  var headers = sheet.getRange(24, 1, 1, lastColumn).getValues()[0];
  
  // Identify all 'Schedule Release' column indices
  var scheduleReleaseCols = [];
  for (var col = 1; col <= headers.length; col++) {
    if (headers[col - 1] === 'Schedule Release') {
      scheduleReleaseCols.push(col);
      Logger.log("'Schedule Release' column found at: " + col + " (" + getColumnLetter(col) + ")");
    }
  }

  if (scheduleReleaseCols.length === 0) {
    Logger.log("No 'Schedule Release' columns found in headers.");
    return;
  }

  // Identify adjacent columns for each 'Schedule Release' column
  var publishedCols = [];
  var releaseDateCols = [];
  scheduleReleaseCols.forEach(function(col) {
    publishedCols.push(col - 1); // 1 column to the left
    releaseDateCols.push(col + 1); // 1 column to the right
  });

  // Fetch all relevant data in one go, up to effectiveLastRow
  var numRows = effectiveLastRow - 25; // Since starting from row 26
  var dataRange = sheet.getRange(26, 1, numRows, lastColumn);
  var data = dataRange.getValues();

  // Prepare arrays to hold updates for 'Release Date' and 'Published' columns
  var releaseDateUpdates = {};
  var publishedUpdates = {};

  // Initialize updates arrays
  releaseDateCols.forEach(function(col) {
    releaseDateUpdates[col] = sheet.getRange(26, col, numRows, 1).getValues();
  });
  publishedCols.forEach(function(col) {
    publishedUpdates[col] = sheet.getRange(26, col, numRows, 1).getValues();
  });

  // Track the number of updates for logging
  var totalUpdates = 0;
  var updatesPerColumn = {};

  // Initialize updatesPerColumn
  scheduleReleaseCols.forEach(function(col) {
    updatesPerColumn[col] = 0;
  });

  // Loop through each row and process each 'Schedule Release' column
  for (var i = 0; i < data.length; i++) {
    var rowNumber = 26 + i;
    Logger.log("Processing row: " + rowNumber);

    scheduleReleaseCols.forEach(function(scheduleColIndex, index) {
      var publishedColIndex = publishedCols[index];
      var releaseDateColIndex = releaseDateCols[index];

      var scheduleDate = data[i][scheduleColIndex - 1]; // Arrays are 0-based
      Logger.log("Schedule date for row " + rowNumber + " in column " + scheduleColIndex + " (" + getColumnLetter(scheduleColIndex) + "): " + scheduleDate);

      // Check if the Schedule Release date is valid and matches today's date
      if (scheduleDate instanceof Date && !isNaN(scheduleDate)) {
        // Clone the date object to avoid mutating the original data
        var scheduleDateCopy = new Date(scheduleDate);
        scheduleDateCopy.setHours(0, 0, 0, 0); // Clear time for accurate comparison

        if (scheduleDateCopy.getTime() === today.getTime()) {
          Logger.log("Valid schedule date found and matches today in column " + scheduleColIndex + " (" + getColumnLetter(scheduleColIndex) + ") for row: " + rowNumber);

          // Update 'Release Date' column
          releaseDateUpdates[releaseDateColIndex][i][0] = scheduleDateCopy;
          Logger.log("Scheduled to move date to 'Release Date' column " + releaseDateColIndex + " (" + getColumnLetter(releaseDateColIndex) + ") for row: " + rowNumber);

          // Update 'Published' column
          publishedUpdates[publishedColIndex][i][0] = true;
          Logger.log("Scheduled to tick 'Published' checkbox in column " + publishedColIndex + " (" + getColumnLetter(publishedColIndex) + ") for row: " + rowNumber);

          // Optionally clear the 'Schedule Release' date
          // Uncomment the line below if you want to clear the 'Schedule Release' date after processing
          // sheet.getRange(rowNumber, scheduleColIndex).setValue('');
          Logger.log("Optionally scheduled to clear 'Schedule Release' date in column " + scheduleColIndex + " (" + getColumnLetter(scheduleColIndex) + ") for row: " + rowNumber);

          // Increment update counters
          totalUpdates++;
          updatesPerColumn[scheduleColIndex]++;
        } else {
          Logger.log("Schedule date does not match today's date for row: " + rowNumber + ", column: " + scheduleColIndex);
        }
      } else {
        Logger.log("Schedule date is not valid for row: " + rowNumber + ", column: " + scheduleColIndex);
      }
    });
  }

  // Bulk update 'Release Date' columns
  releaseDateCols.forEach(function(col) {
    // Check if there are any updates in this column
    var hasUpdates = false;
    for (var i = 0; i < releaseDateUpdates[col].length; i++) {
      if (releaseDateUpdates[col][i][0] !== data[i][col - 1]) {
        hasUpdates = true;
        break;
      }
    }
    if (hasUpdates) {
      sheet.getRange(26, col, releaseDateUpdates[col].length, 1).setValues(releaseDateUpdates[col]);
      Logger.log("Updated 'Release Date' column " + col + " (" + getColumnLetter(col) + ") with " + updatesPerColumn[col] + " updates.");
    }
  });

  // Bulk update 'Published' columns
  publishedCols.forEach(function(col) {
    // Check if there are any updates in this column
    var hasUpdates = false;
    for (var i = 0; i < publishedUpdates[col].length; i++) {
      if (publishedUpdates[col][i][0] !== data[i][col - 1]) {
        hasUpdates = true;
        break;
      }
    }
    if (hasUpdates) {
      sheet.getRange(26, col, publishedUpdates[col].length, 1).setValues(publishedUpdates[col]);
      Logger.log("Updated 'Published' column " + col + " (" + getColumnLetter(col) + ") with " + updatesPerColumn[col] + " updates.");
    }
  });

  Logger.log("Total updates performed: " + totalUpdates);
  Logger.log("updateReleaseDates script completed.");
}

// Helper function to convert column number to letter
function getColumnLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function updateSheetLinks() {
  var lock = LockService.getScriptLock();
  try {
    // Wait for up to 30 seconds to acquire the lock
    lock.waitLock(30000);
  } catch (e) {
    Logger.log("Could not obtain lock: " + e.message);
    return;
  }

  try {
    Logger.log("Starting updateSheetLinks function...");

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MASTER");
    if (!sheet) {
      Logger.log("Sheet 'MASTER' not found.");
      return;
    }

    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    Logger.log("Last row of the sheet: " + lastRow);
    Logger.log("Last column of the sheet: " + lastColumn);

    // Fetch headers once (assuming headers are in row 24)
    var headers = sheet.getRange(24, 1, 1, lastColumn).getValues()[0];

    var chapterCols = [];
    var sheetLinkCols = [];

    // Identify all "CHAPTER" and "Sheet Link" columns
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i].toString().trim().toUpperCase();
      if (header === 'CHAPTER') {
        chapterCols.push(i + 1); // Store the column index of 'CHAPTER'
        Logger.log("'CHAPTER' column found at: " + (i + 1) + " (" + getColumnLetter(i + 1) + ")");
      } else if (header === 'SHEET LINK') {
        sheetLinkCols.push(i + 1); // Store the column index of 'Sheet Link'
        Logger.log("'Sheet Link' column found at: " + (i + 1) + " (" + getColumnLetter(i + 1) + ")");
      }
    }

    if (chapterCols.length === 0 || sheetLinkCols.length === 0) {
      Logger.log('No "CHAPTER" or "Sheet Link" columns found.');
      return;
    }

    // Map each 'CHAPTER' column to its nearest right 'Sheet Link' column
    var chapterToSheetLinkMap = {};

    chapterCols.forEach(function(chapterCol) {
      var nearestLinkCol = findNearestRightSheetLinkColumn(chapterCol, sheetLinkCols);
      if (nearestLinkCol) {
        chapterToSheetLinkMap[chapterCol] = nearestLinkCol;
        Logger.log("'CHAPTER' column " + chapterCol + " (" + getColumnLetter(chapterCol) + ") mapped to 'Sheet Link' column " + nearestLinkCol + " (" + getColumnLetter(nearestLinkCol) + ")");
      } else {
        Logger.log("'CHAPTER' column " + chapterCol + " (" + getColumnLetter(chapterCol) + ") has no corresponding 'Sheet Link' column to the right.");
      }
    });

    // Fetch all data from row 26 to lastRow in bulk
    var dataRange = sheet.getRange(26, 1, lastRow - 25, lastColumn);
    var allData = dataRange.getValues();

    // Prepare an object to store updates for 'Sheet Link' columns
    var sheetLinkUpdates = {};

    // Initialize sheetLinkUpdates with existing 'Sheet Link' values
    sheetLinkCols.forEach(function(linkCol) {
      sheetLinkUpdates[linkCol] = [];
      for (var i = 0; i < allData.length; i++) {
        sheetLinkUpdates[linkCol].push([allData[i][linkCol - 1]]); // Initialize with existing values
      }
    });

    var totalUpdates = 0;

    // Iterate through each row and process 'CHAPTER' columns
    for (var rowIndex = 0; rowIndex < allData.length; rowIndex++) {
      var rowNumber = 26 + rowIndex;
      var rowData = allData[rowIndex];

      for (var chapterCol in chapterToSheetLinkMap) {
        var sheetLinkCol = chapterToSheetLinkMap[chapterCol];
        var chapterText = rowData[chapterCol - 1];

        if (chapterText) {
          var existingLink = rowData[sheetLinkCol - 1];
          if (!existingLink) { // Only update if 'Sheet Link' cell is empty
            var chapterNumber = extractChapterNumber(chapterText);
            if (chapterNumber === null) {
              Logger.log("Invalid chapter format in row " + rowNumber + " column " + chapterCol + ": " + chapterText);
              continue;
            }

            var sheetToLink = findSheetByChapterNumber(chapterNumber);
            if (sheetToLink) {
              var sheetUrl = constructSheetUrl(sheetToLink);
              sheetLinkUpdates[sheetLinkCol][rowIndex][0] = sheetUrl;
              Logger.log("Row " + rowNumber + ", Column " + sheetLinkCol + " (" + getColumnLetter(sheetLinkCol) + ") updated with URL: " + sheetUrl);
              totalUpdates++;
            } else {
              Logger.log("No sheet found for Chapter " + chapterNumber + " in row " + rowNumber);
            }
          }
        }
      }
    }

    // Prepare an array for each 'Sheet Link' column to be updated
    sheetLinkCols.forEach(function(linkCol) {
      var hasUpdates = false;
      for (var i = 0; i < sheetLinkUpdates[linkCol].length; i++) {
        if (sheetLinkUpdates[linkCol][i][0] !== allData[i][linkCol - 1]) {
          hasUpdates = true;
          break;
        }
      }
      if (hasUpdates) {
        sheet.getRange(26, linkCol, sheetLinkUpdates[linkCol].length, 1).setValues(sheetLinkUpdates[linkCol]);
        Logger.log("Updated 'Sheet Link' column " + linkCol + " (" + getColumnLetter(linkCol) + ") with updates.");
      }
    });

    Logger.log("Total 'Sheet Link' updates performed: " + totalUpdates);
    Logger.log("updateSheetLinks function completed.");
  } finally {
    // Release the lock after script execution
    lock.releaseLock();
  }
}

function findSheetByChapterNumber(chapterNumber) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNameLower = "c" + chapterNumber;
  var sheetNameUpper = "C" + chapterNumber;

  var sheet = spreadsheet.getSheetByName(sheetNameLower);
  if (sheet) {
    return sheet;
  }

  sheet = spreadsheet.getSheetByName(sheetNameUpper);
  if (sheet) {
    return sheet;
  }

  return null;
}

function constructSheetUrl(sheet) {
  var spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var sheetId = sheet.getSheetId();
  return spreadsheetUrl + "#gid=" + sheetId;
}

function extractChapterNumber(chapterText) {
  var regex = /Chapter\s*(\d+)/i;
  var match = chapterText.match(regex);
  if (match && match[1]) {
    return match[1];
  }
  return null;
}

function findNearestRightSheetLinkColumn(chapterCol, sheetLinkCols) {
  for (var i = 0; i < sheetLinkCols.length; i++) {
    if (sheetLinkCols[i] > chapterCol) {
      return sheetLinkCols[i]; // Return the first "Sheet Link" column to the right of the "CHAPTER" column
    }
  }
  return null; // Return null if no "Sheet Link" column is found to the right
}

function patternColumnD() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Define the starting row and column
  var startRow = 4;
  var lastRow = sheet.getLastRow();
  
  if (lastRow < startRow) {
    return; // No data to move if fewer than 4 rows
  }
  
  // Get values from column D starting from row 4
  var range = sheet.getRange(startRow, 4, lastRow - startRow + 1, 1);
  var values = range.getValues();
  
  var newValues = [];
  
  // Create the new pattern of text followed by a blank row
  for (var i = 0; i < values.length; i++) {
    var value = values[i][0];
    if (value !== '') {
      newValues.push([value]);
      newValues.push(['']); // Add a blank row after the text
    }
  }
  
  // Fill remaining rows with empty values if newValues has fewer rows
  var numRowsToFill = lastRow - startRow + 1;
  while (newValues.length < numRowsToFill) {
    newValues.push(['']);
  }
  
  // Set the updated values back to column D
  range.setValues(newValues);
}

function patternColumnA() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Define the starting row and column
  var startRow = 4;
  var column = 1; // Column A corresponds to the 1st column
  var lastRow = sheet.getLastRow();
  
  if (lastRow < startRow) {
    return; // No data to move if fewer than 4 rows
  }
  
  // Get values from column A starting from row 4
  var range = sheet.getRange(startRow, column, lastRow - startRow + 1, 1);
  var values = range.getValues();
  
  var newValues = [];
  
  // Create the new pattern of text followed by a blank row
  for (var i = 0; i < values.length; i++) {
    var value = values[i][0];
    if (value !== '') {
      newValues.push([value]);
      newValues.push(['']); // Add a blank row after the text
    }
  }
  
  // Fill remaining rows with empty values if newValues has fewer rows
  var numRowsToFill = lastRow - startRow + 1;
  while (newValues.length < numRowsToFill) {
    newValues.push(['']);
  }
  
  // Trim newValues to match the exact number of rows needed
  if (newValues.length > numRowsToFill) {
    newValues = newValues.slice(0, numRowsToFill);
  }
  
  // Set the updated values back to column A
  range.setValues(newValues);
}

function patternColumnB() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Define the starting row and column
  var startRow = 4;
  var column = 2; // Column B corresponds to the 2nd column
  var lastRow = sheet.getLastRow();
  
  if (lastRow < startRow) {
    return; // No data to move if fewer than 4 rows
  }
  
  // Get values from column B starting from row 4
  var range = sheet.getRange(startRow, column, lastRow - startRow + 1, 1);
  var values = range.getValues();
  
  var newValues = [];
  
  // Create the new pattern of text followed by a blank row
  for (var i = 0; i < values.length; i++) {
    var value = values[i][0];
    if (value !== '') {
      newValues.push([value]);
      newValues.push(['']); // Add a blank row after the text
    }
  }
  
  // Fill remaining rows with empty values if newValues has fewer rows
  var numRowsToFill = lastRow - startRow + 1;
  while (newValues.length < numRowsToFill) {
    newValues.push(['']);
  }
  
  // Trim newValues to match the exact number of rows needed
  if (newValues.length > numRowsToFill) {
    newValues = newValues.slice(0, numRowsToFill);
  }
  
  // Set the updated values back to column B
  range.setValues(newValues);
}

function getChaptersAndCheckboxes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MASTER");
  if (!sheet) {
    Logger.log("Sheet 'MASTER' not found.");
    return;
  }

  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(24, 1, 1, sheet.getLastColumn());
  var headers = range.getValues()[0];

  var chapterCols = [];
  var checkboxesCols = []; // Assuming you have a way to determine checkbox columns

  // Identify all "CHAPTER" and "Checkbox" columns
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toUpperCase() === 'CHAPTER') {
      chapterCols.push(i + 1); // Store the column index of 'CHAPTER'
    } else if (headers[i].toUpperCase() === 'Translate') { // Example checkbox header
      checkboxesCols.push(i + 1); // Store the column index of 'Checkbox'
    }
  }

  if (chapterCols.length === 0) {
    Logger.log('No "CHAPTER" columns found.');
    return;
  }

  var chapters = [];
  var defaultChapter = "";

  // Fetch chapter names from the "CHAPTER" columns
  for (var row = 26; row <= lastRow; row++) {
    chapterCols.forEach(function(chapterCol) {
      var chapterText = sheet.getRange(row, chapterCol).getValue();
      if (chapterText) {
        if (chapters.indexOf(chapterText) === -1) {
          chapters.push(chapterText); // Add the chapter name if it's not already in the list
        }
      }
    });
  }

  // Determine the default chapter (find the chapter related to the current sheet)
  var currentSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  var chapterNumber = currentSheetName.replace(/[^0-9]/g, ''); // Extract number from sheet name
  defaultChapter = "Chapter " + chapterNumber;

  var checkboxes = []; // Fetch checkbox states if applicable

  // Populate checkboxes (assuming you have some logic for this)
  checkboxesCols.forEach(function(checkboxCol) {
    for (var row = 26; row <= lastRow; row++) {
      var checkboxValue = sheet.getRange(row, checkboxCol).getValue();
      // Example of processing checkbox values
      if (checkboxValue) {
        checkboxes.push({
          id: row + "-" + checkboxCol, // Unique ID example
          label: "Checkbox " + row, // Example label
          checked: checkboxValue === true // Example checked state
        });
      }
    }
  });

  return {
    chapters: chapters,
    defaultChapter: defaultChapter,
    checkboxes: checkboxes
  };
}



function openTranslationDialog() {
  var html = HtmlService.createHtmlOutputFromFile('updateTranslationStatusDialog')
    .setWidth(500)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}


function updateTranslationStatus(selectedChapter, updates) {
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MASTER');
  if (masterSheet) {
    var range = masterSheet.getDataRange();
    var values = range.getValues();
    var headers = values[0];
    var chapterColIndex = headers.indexOf('CHAPTER');
    var translatedColIndex = headers.indexOf('TRANSLATED');
    
    if (chapterColIndex !== -1 && translatedColIndex !== -1) {
      updates.forEach(function(update) {
        values.forEach(function(row, index) {
          if (index > 0) { // Skip header row
            if (row[chapterColIndex] === selectedChapter) {
              masterSheet.getRange(index + 1, translatedColIndex + 1).setValue(update.checked);
            }
          }
        });
      });
    }
  }
}



function createTrigger() {
  ScriptApp.newTrigger('updateReleaseDates')
    .timeBased()
    .everyDays(1) // Runs the script daily
    .atHour(0) // Runs at midnight
    .create();
}

function createHyperlinks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 2; // Starting from row 2 (H2)
  var column = 8; // Column H is the 8th column
  var lastRow = sheet.getLastRow();
  
  // Calculate the number of rows to process
  var numRows = lastRow - startRow + 1;
  
  if (numRows < 1) {
    // No data to process
    return;
  }
  
  // Get the range of URLs in Column H starting from H2
  var range = sheet.getRange(startRow, column, numRows, 1);
  var values = range.getValues();
  
  // Prepare an array to hold the HYPERLINK formulas
  var formulas = [];
  
  for (var i = 0; i < values.length; i++) {
    var url = values[i][0];
    
    if (url && typeof url === 'string' && url.trim() !== "") {
      // Construct the HYPERLINK formula
      var formula = `=HYPERLINK("${url}", "${url}")`;
      formulas.push([formula]);
    } else {
      // If the cell is empty or invalid, keep it blank
      formulas.push([""]);
    }
  }
  
  // Set the formulas in the range
  range.setFormulas(formulas);
}

function retainFirst300RowsBatch() {
  const ROW_LIMIT = 300;          
  const BATCH_SIZE = 5;           // Process only 5 sheets per batch
  const SHEET_PREFIX = 'c';       
  const RECIPIENT_EMAIL = 'your-email@example.com';
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  
  // Use script properties to track what's been done
  const scriptProperties = PropertiesService.getScriptProperties();
  let processedSheets = JSON.parse(scriptProperties.getProperty('processedSheets') || '[]');
  
  // Filter only sheets that start with 'c' and are not yet processed
  const sheetsToProcess = sheets.filter(sheet => {
    const sheetName = sheet.getName().trim().toLowerCase();
    return sheetName.startsWith(SHEET_PREFIX) && !processedSheets.includes(sheet.getName());
  }).slice(0, BATCH_SIZE);
  
  // If there are no sheets left to process...
  if (sheetsToProcess.length === 0) {
    sendCompletionEmail(RECIPIENT_EMAIL);
    scriptProperties.deleteProperty('processedSheets');
    SpreadsheetApp.getUi().alert('All relevant sheets processed.');
    return;
  }
  
  // Process each sheet in this batch
  sheetsToProcess.forEach(sheet => {
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    
    if (lastRow > ROW_LIMIT) {
      const rowsToDelete = lastRow - ROW_LIMIT;
      try {
        sheet.deleteRows(ROW_LIMIT + 1, rowsToDelete);
        Logger.log(`Processed: ${sheetName}, deleted ${rowsToDelete} rows.`);
      } catch (error) {
        Logger.log(`Error on ${sheetName}: ${error}`);
      }
    } else {
      Logger.log(`No action: ${sheetName} only has ${lastRow} rows.`);
    }
    
    // Mark this sheet as processed
    processedSheets.push(sheetName);
  });
  
  // Update which sheets are processed
  scriptProperties.setProperty('processedSheets', JSON.stringify(processedSheets));
  
  // Schedule next batch if we processed a full BATCH_SIZE
  if (sheetsToProcess.length === BATCH_SIZE) {
    ScriptApp.newTrigger('retainFirst300RowsBatch')
      .timeBased()
      .after(60 * 1000) // 1-minute delay
      .create();
  } else {
    // We reached the end
    sendCompletionEmail(RECIPIENT_EMAIL);
    scriptProperties.deleteProperty('processedSheets');
    SpreadsheetApp.getUi().alert('Completed processing all relevant sheets.');
  }
}

function sendCompletionEmail(recipient) {
  const subject = 'Google Sheets Processing - Completed';
  const body = 'All relevant sheets have been processed.';
  MailApp.sendEmail(recipient, subject, body);
}

/**
 * Optional: Function to reset processed sheets tracking.
 * Useful if you need to rerun the script from the beginning.
 */
function resetProcessedSheets() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('processedSheets');
  SpreadsheetApp.getUi().alert('Processed sheets tracking has been reset.');
}

/**
 * Deletes triggers associated with specific handler functions.
 * Useful for cleanup without affecting other triggers.
 */
function deleteSpecificTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const handlerFunctionsToDelete = ['retainFirst300RowsBatch']; // Add other handler function names as needed

  let deletedCount = 0;

  triggers.forEach(trigger => {
    const handlerFunction = trigger.getHandlerFunction();
    if (handlerFunctionsToDelete.includes(handlerFunction)) {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });

  SpreadsheetApp.getUi().alert(`Deleted ${deletedCount} trigger(s) associated with handler function(s): ${handlerFunctionsToDelete.join(', ')}.`);
}
