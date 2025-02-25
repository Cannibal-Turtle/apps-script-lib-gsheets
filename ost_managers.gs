function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Custom')
  menu.addItem('Go to Master\'s Sheet', 'goToMastersSheet');
  menu.addItem('Go to Next Sheet', 'goToNextSheet')
  menu.addItem('Go to Previous Sheet', 'goToPreviousSheet')
  menu.addItem('Renumber Sheets', 'checkAndStartRenumbering')
  menu.addSeparator();
  menu.addItem('Input Media', 'openInputDialog')
  menu.addItem('Sidebar', 'showSidebar'); // Existing menu item for the audio player
  // Add the menu to the Google Sheets UI
  menu.addToUi();
}

function showSidebar() {
  // HTML content for the sidebar, including a simple interface to play music
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Sidebar ♬')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function goToMastersSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = spreadsheet.getSheetByName('Master [OST 2022]'); // Replace with the name of your master sheet
  if (masterSheet) {
    spreadsheet.setActiveSheet(masterSheet);
  } else {
    SpreadsheetApp.getUi().alert('Master sheet not found.'); // Alert if the master sheet is not found
  }
}

function goToNextSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = spreadsheet.getActiveSheet();
  var currentSheetName = currentSheet.getName();
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var songData = JSON.parse(scriptProperties.getProperty("songData"));

  for (var i = 0; i < songData.length; i++) {
    if (songData[i].sheetName === currentSheetName) {
      var nextSheetName = (i < songData.length - 1) ? songData[i + 1].sheetName : songData[0].sheetName;
      spreadsheet.getSheetByName(nextSheetName).activate();
      break;
    }
  }
}

function goToPreviousSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = spreadsheet.getActiveSheet();
  var currentSheetName = currentSheet.getName();
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var songData = JSON.parse(scriptProperties.getProperty("songData"));

  for (var i = 0; i < songData.length; i++) {
    if (songData[i].sheetName === currentSheetName) {
      var prevSheetName = (i > 0) ? songData[i - 1].sheetName : songData[songData.length - 1].sheetName;
      spreadsheet.getSheetByName(prevSheetName).activate();
      break;
    }
  }
}


function openInputDialog() {
  var activeSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  
  // Assuming "Master [OST 2022]" is the name of the Master sheet
  if (activeSheetName === "Master [OST 2022]") {
    openDialogMasterSheet();
  } else {
    openDialogOtherSheets();
  }
}

function openDialogMasterSheet() {
  var html = HtmlService.createHtmlOutputFromFile('Dialog Master Sheet')
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Input Media Link ♬');
}

function openDialogOtherSheets() {
  var html = HtmlService.createHtmlOutputFromFile('Dialog Other Sheets')
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Input Media Link ♬');
}

function matchSheetNamesToURLs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Master [OST 2022]"); // Adjust if your master sheet has a different name
  const urlsRange = masterSheet.getRange("E2:E"); // Assumes URLs start from E2. Adjust if your range is different.
  const urls = urlsRange.getValues();
  let sheetNames = [];

  // Loop through each URL and find the corresponding sheet name
  urls.forEach(function(urlArray) {
    const url = urlArray[0];
    if (url) { // Check if URL is not empty
      const sheetId = extractSheetIdFromUrl(url);
      const sheet = findSheetById(ss, sheetId);
      if (sheet) {
        const sheetName = sheet.getName();
        // Exclude the Master sheet
        if (sheetName !== "Master [OST 2022]") {
          sheetNames.push([sheetName]);
        } else {
          sheetNames.push([""]); // Keep the array aligned with URLs
        }
      } else {
        sheetNames.push([""]); // In case no matching sheet is found
      }
    } else {
      sheetNames.push([""]); // For empty cells in the URL column
    }
  });

  // Write the matched sheet names to Column G, starting from G2
  masterSheet.getRange(2, 7, sheetNames.length, 1).setValues(sheetNames); // Column G is index 7
}


function romanizeHangul(input) {
  if (Array.isArray(input)) {
    return input.map(romanizeHangul);
  }

  if (typeof input !== 'string') {
    return "Invalid input. Please provide a string.";
  }

  var output = "";
  for (var i = 0; i < input.length; i++) {
    var unicode = input.charCodeAt(i);
    if (unicode >= 44032 && unicode <= 55203) {
      var syllableIndex = unicode - 44032;
      var jong = syllableIndex % 28;
      var jung = ((syllableIndex - jong) / 28) % 21;
      var cho = (((syllableIndex - jong) / 28) - jung) / 21;
      
      var choList = ["g", "gg", "n", "d", "dd", "r", "m", "b", "bb", "s", "ss", "", "j", "jj", "ch", "k", "t", "p", "h"];
      var jungList = ["a", "ae", "ya", "yae", "eo", "e", "yeo", "ye", "o", "wa", "wae", "oe", "yo", "u", "weo", "we", "wi", "yu", "eu", "ui", "i"];
      var jongList = ["", "g", "gg", "gs", "n", "nj", "nh", "d", "l", "lg", "lm", "lb", "ls", "lt", "lp", "lh", "m", "b", "bs", "s", "ss", "ng", "j", "ch", "k", "t", "p", "h"];
      
      output += choList[cho] + jungList[jung] + jongList[jong];
    } else {
      output += input.charAt(i);
    }
  }
  return output;
}

function onEdit(e) {
  var editedSheet = e.range.getSheet();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = spreadsheet.getSheetByName("Master [OST 2022]");

    // Ensure the edit is not in the "Master [OST 2022]" sheet and is in column D
  if (editedSheet.getName() !== "Master [OST 2022]" && e.range.getColumn() === 4) {
    compactColumnD(sheet);
    trimColumnD();
    numberTracksAcrossSheetsBasedOnMaster(spreadsheet);}

  // Direct synchronization: From "Master [OST 2022]" column H to sheet H2
  if (editedSheet.getName() === "Master [OST 2022]" && e.range.getColumn() === 8) { 
    var urlCell = masterSheet.getRange(e.range.getRow(), 5); // URL is in column E
    var sheetUrl = urlCell.getValue();
    var modifiedColumnHValue = modifyUrlForPreview(e.range.getValue()); // Apply modification function
    Logger.log("Direct Sync - URL: " + sheetUrl); // Logging the URL
    var sheetId = extractSheetIdFromUrl(sheetUrl);
    if (sheetId) {
      var targetSheet = findSheetById(spreadsheet, sheetId);
      if (targetSheet) {
        targetSheet.getRange("H2").setValue(modifiedColumnHValue);
        Logger.log("Direct Sync - Updated H2 in sheet ID: " + sheetId + " with value: " + modifiedColumnHValue);
      } else {
        Logger.log("Direct Sync - No target sheet found with ID: " + sheetId);
      }
    } else {
      Logger.log("Direct Sync - No valid sheet ID extracted from URL: " + sheetUrl);
    }
  } 
  // Reverse synchronization: From any sheet H2 to "Master [OST 2022]" column H
  else if (e.range.getA1Notation() === 'H2' && editedSheet.getName() !== "Master [OST 2022]") {
    var allEntries = masterSheet.getRange('E2:H' + masterSheet.getLastRow()).getValues();
    var editedSheetId = editedSheet.getSheetId().toString();
    var found = false;

    for (var i = 0; i < allEntries.length; i++) {
      var rowSheetId = extractSheetIdFromUrl(allEntries[i][0]);
      if (rowSheetId === editedSheetId) {
        var modifiedColumnHValue = modifyUrlForPreview(e.range.getValue()); // Apply modification function
        masterSheet.getRange(i + 2, 8).setValue(modifiedColumnHValue); // Update column H with modified value
        found = true;
        Logger.log("Reverse Sync - Updated Master [OST 2022] column H for Sheet ID: " + editedSheetId + " with value: " + modifiedColumnHValue);
        break;
      }
    }

    if (!found) {
      Logger.log("Reverse Sync - No matching sheet found for Sheet ID: " + editedSheetId);
    }
  }

  // New logic for Column I synchronization
  // Direct synchronization: From "Master [OST 2022]" column I to other sheets' column I
  if (editedSheet.getName() === "Master [OST 2022]" && e.range.getColumn() === 9) { // Column I is 9
    var urlCell = masterSheet.getRange(e.range.getRow(), 5); // Assuming URL is in column E
    var modifiedColumnIValue = convertToEmbedUrl(e.range.getValue()); // Apply conversion function
    Logger.log("Direct Sync - URL: " + sheetUrl);
    var sheetId = extractSheetIdFromUrl(urlCell.getValue()); // Reusing urlCell
    if (sheetId) {
      var targetSheet = findSheetById(spreadsheet, sheetId);
      if (targetSheet) {
        targetSheet.getRange("I2").setValue(modifiedColumnIValue); // Update Column I in target sheet
        Logger.log("Direct Sync - Updated I2 in sheet ID: " + sheetId + " with value: " + modifiedColumnIValue);
      } else {
        Logger.log("Direct Sync - No target sheet found with ID: " + sheetId);
      }
    } else {
      Logger.log("Direct Sync - No valid sheet ID extracted from URL: " + sheetUrl);
    }
  }
  // Reverse synchronization: From any sheet's column I to "Master [OST 2022]" column I
  else if (e.range.getA1Notation() === 'I2' && editedSheet.getName() !== "Master [OST 2022]") {
    var allEntries = masterSheet.getRange('E2:I' + masterSheet.getLastRow()).getValues();
    var editedSheetId = editedSheet.getSheetId().toString();
    var found = false;

    for (var i = 0; i < allEntries.length; i++) {
      var rowSheetId = extractSheetIdFromUrl(allEntries[i][0]);
      if (rowSheetId === editedSheetId) {
        var modifiedColumnIValue = convertToEmbedUrl(e.range.getValue()); // Apply conversion function
        masterSheet.getRange(i + 2, 9).setValue(modifiedColumnIValue); // Update Column I in master sheet
        found = true;
        Logger.log("Reverse Sync - Updated Master [OST 2022] column I for Sheet ID: " + editedSheetId + " with value: " + modifiedColumnIValue);
        break;
      }
    }

    if (!found) {
      Logger.log("Reverse Sync - No matching sheet found for Sheet ID: " + editedSheetId);
    }
  }

//Youtube URL//
  var editedCell = e.range;

  // Check if the edited cell is in column I (column 9) and is not empty
  if (editedCell.getColumn() == 9 && editedCell.getValue() != "") {
    var newValue = convertToEmbedUrl(editedCell.getValue());
    // Set the new value in the edited cell
    editedCell.setValue(newValue);
  }

// MP3 URL//
  var editedCell = e.range

  // Check if the edit is in column H and the cell is not empty
  if (editedCell.getColumn() === 8 && editedCell.getValue() !== "") {
    var originalValue = editedCell.getValue();
    
    // Call modifyUrlForPreview with the cell's value
    var newValue = modifyUrlForPreview(originalValue);

    // Update the cell with the new value if it's different from the original value
    if (newValue !== originalValue) {
      editedCell.setValue(newValue);
    }
  }

  // If the event object is not present, get the active spreadsheet
var sheet = e && e.source ? e.source.getSheetByName(e.range.getSheet().getName()) : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Check if the sheet is not named "Master [OST 2022]" and column D"
  if (editedSheet.getName() !== "Master [OST 2022]" && e.range.getColumn() === 4) {
  var dataRange = sheet.getRange('D1:D' + sheet.getLastRow()); // Adjusted to get dynamic range up to the last row
  var values = dataRange.getValues().flat(); // Flatten to a single-dimensional array for easier manipulation
  var filteredValues = values.filter(value => value !== ""); // Remove empty strings
  var numRows = filteredValues.length;

  // Clear the original data range
  dataRange.clearContent();

  // Write back the filtered values to Column D, starting from the first cell
  if (numRows > 0) {
    sheet.getRange('D1:D' + numRows).setValues(filteredValues.map(value => [value]));
  }
}
  var range = e ? e.range : sheet.getActiveCell();
 // Find all occurrences of the same hangul in column D
  if (range.getColumn() == 6 && range.getValue() !== "") {
    var translation = range.getValue();
    var hangul = sheet.getRange(e.range.getRow(), 4).getValue().trim();
    var duplicateRanges = sheet.getRange(1, 4, sheet.getLastRow(), 1)
      .createTextFinder(hangul)
      .findAll();

  // Split the hangul into individual words
    var hangulWords = hangul.split(" ");

  // Update the translations in the duplicate rows in column F
    duplicateRanges.forEach(function (duplicateRange) {
  // Get the value in column D for the current row
    var correspondingHangul = sheet.getRange(duplicateRange.getRow(), 4).getValue();
  // Split the corresponding hangul into individual words
    var correspondingHangulWords = correspondingHangul.split(" ");
  
  // Check if the corresponding value in column D matches exactly
  if (isExactMatch(hangulWords, correspondingHangulWords)) {
    sheet.getRange(duplicateRange.getRow(), 6).setValue(translation); // Assuming column F is the 6th column
    }
  });
} 
} 

// Function to check if two arrays of words match exactly
function isExactMatch(words1, words2) {
  if (words1.length !== words2.length) {
    return false;
  }
  for (var i = 0; i < words1.length; i++) {
    if (words1[i] !== words2[i]) {
      return false;
    }
  }
  return true;
}

function extractSheetIdFromUrl(url) {
  var match = url.match(/#gid=(\d+)/);
  return match ? match[1] : null;
}

function findSheetById(spreadsheet, sheetId) {
  console.log("Attempting to find sheet with ID: ", sheetId);
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId().toString() === sheetId) {
      return sheets[i];
    }
  }
  return null;
}

// Function to retrieve filename from Google Drive link
function retrieveFilename() {
  // Access the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // Get the range of cells containing Google Drive links (assuming they are in column H)
  var range = sheet.getRange('H2:H' + sheet.getLastRow());

  // Get the values of the range (Google Drive links)
  var links = range.getValues();

  // Iterate through each link
  links.forEach(function(link) {
    if (link[0]) { // Check if the cell is not empty
      var fileId = extractFileId(link[0]); // Extract file ID from the link
      var filename = fileId ? retrieveFilename(fileId) : 'Invalid link'; // Retrieve filename using file ID

      // Do something with the filename, such as logging it
      Logger.log('Filename for ' + link[0] + ' is ' + filename);
    }
  });
}

// Function to extract file ID from Google Drive link
function extractFileId(link) {
  var match = link.match(/\/d\/([^/]+)\//);
  if (match && match[1]) {
    return match[1];
  } else {
    return null;
  }
}

// Function to retrieve filename using file ID
function retrieveFilename(fileId) {
  var file = DriveApp.getFileById(fileId);
  return file.getName();
}

function getMediaLinks() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = spreadsheet.getSheetByName("Master [OST 2022]");
  var activeSheet = spreadsheet.getActiveSheet();
  var sheetName = activeSheet.getName();
  var links = [];

  // Check if the active sheet is the Master sheet
  if (sheetName === "Master [OST 2022]") {
    var lastRow = masterSheet.getLastRow();
    var range = masterSheet.getRange('B2:I' + lastRow);
    var values = range.getValues();

    // Fetch all song names and their corresponding media links
    values.forEach(function(row) {
      var songName = row[0]; // Column B: Song name
      var audioLink = row[6]; // Column H: Audio link
      var videoLink = row[7]; // Column I: Video link
      if (audioLink || videoLink) {
        var concatenatedLink = audioLink + '|' + videoLink;
        links.push({ name: songName, link: concatenatedLink });
      }
    });
  } else {
    // For non-Master sheets, match sheet ID to URLs in the Master sheet to find the corresponding song title
    var masterData = masterSheet.getRange('B2:E' + masterSheet.getLastRow()).getValues();
    var activeSheetId = activeSheet.getSheetId().toString();
    var songName = ""; // Initialize songName as an empty string

    for (var i = 0; i < masterData.length; i++) {
      var row = masterData[i];
      var url = row[3]; // Column E: URL
      var sheetIdFromUrl = extractSheetIdFromUrl(url);
      if (sheetIdFromUrl === activeSheetId) {
        songName = row[0]; // Column B: Song name
        break; // Stop the loop once the matching song title is found
      }
    }

    // Assuming audio and video links are in H2 and I2 of the active sheet
    var audioLink = activeSheet.getRange('H2').getValue();
    var videoLink = activeSheet.getRange('I2').getValue();
    if (audioLink || videoLink) {
      var concatenatedLink = audioLink + '|' + videoLink;
      links.push({ name: songName, link: concatenatedLink });
    }
  }

  return links;
}

// Function to convert a YouTube URL to the embed format
function convertToEmbedUrl(url) {
  // If the URL is in the format https://www.youtube.com/watch?v=o1I87SHTh9c
  if (url.includes('youtube.com/watch?v=')) {
    return url.replace('youtube.com/watch?v=', 'youtube.com/embed/');
  }
  // If the URL is in the format https://youtu.be/o1I87SHTh9c
  else if (url.includes('youtu.be/')) {
    return url.replace('youtu.be/', 'youtube.com/embed/');
  }
  // If the URL format is unrecognized, return the original URL
  else {
    return url;
  }
}

// Your existing openAudioPlayer function
const openAudioPlayer = () => {
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const cell = activeSheet.getActiveCell().getValue();
  var title = activeSheet.getName(); // Use sheet name as the title
  const html = `<iframe src="${cell}" width="400" height="80" frameborder="0" scrolling="no"></iframe>`;
  const dialog = HtmlService.createHtmlOutput(html).setTitle('Play').setWidth(420).setHeight(100);
  SpreadsheetApp.getUi().showModelessDialog(dialog, `Play Audio: ${title}`);
};

const openYouTubeVideo = () => {
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const cell = activeSheet.getActiveCell().getValue();
  var title = activeSheet.getName(); // Use sheet name as the title
  
  // Extract the VIDEO_ID from the YouTube URL
  const videoIdMatch = cell.match(/(?:https?:\/\/)?(?:www\.)?(?:youtube\.com\/(?:[^\/\n\s]+\/\S+\/|(?:v|e(?:mbed)?)\/|\S*?[?&]v=)|youtu\.be\/)([a-zA-Z0-9_-]{11})/);
  if (videoIdMatch && videoIdMatch[1]) {
    const videoId = videoIdMatch[1];
    const embedUrl = `https://www.youtube.com/embed/${videoId}`;
    const html = `<iframe src="${embedUrl}" width="560" height="315" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>`;
    const dialog = HtmlService.createHtmlOutput(html).setTitle('Watch Video').setWidth(580).setHeight(360);
    SpreadsheetApp.getUi().showModelessDialog(dialog, `Watch Video: ${title}`);
  } else {
    SpreadsheetApp.getUi().alert('Invalid YouTube URL');
  }
};

function getSheetData() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master [OST 2022]");
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
    var sheetData = [];

    for (var i = 1; i < data.length; i++) { // Start at 1 to skip header row
        var row = data[i];
        sheetData.push({
            songTrack: row[1], // Assuming song track is in the second column (B)
            row: i + 1, // +1 because array is 0-indexed but sheets are 1-indexed
            columnHValue: row[7], // Assuming column H values are in the 8th array position
            columnIValue: row[8] // Assuming column I values are in the 9th array position
        });
    }

    return sheetData;
}

function submitInfoForMaster(songTrack, columnHValue, columnIValue) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Master [OST 2022]");
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var modifiedColumnHValue = modifyUrlForPreview(columnHValue);
    var modifiedColumnIValue = convertToEmbedUrl(columnIValue);
  
    // Update the master sheet
    for (var i = 0; i < values.length; i++) {
        if (values[i][1] == songTrack) {
            sheet.getRange(i + 1, 8).setValue(modifiedColumnHValue);
            sheet.getRange(i + 1, 9).setValue(modifiedColumnIValue);
            break; // Exit after updating
        }
    }
  
    // Synchronize other sheets based on URLs in Column E
    values.forEach(function(row) {
        var url = row[4]; // Column E contains the URLs
        if (url) {
            var sheetId = extractSheetIdFromUrl(url);
            var targetSheet = findSheetById(spreadsheet, sheetId);
            if (targetSheet) {
                targetSheet.getRange("H2").setValue(modifiedColumnHValue); // Update Column H in target sheet
                targetSheet.getRange("I2").setValue(modifiedColumnIValue); // Update Column I in target sheet
            }
        }
    });
}

function submitInfoForOtherSheets(columnHValue, columnIValue) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var currentSheet = spreadsheet.getActiveSheet();
    var masterSheet = spreadsheet.getSheetByName("Master [OST 2022]");
    var sheetUrl = spreadsheet.getUrl() + "#gid=" + currentSheet.getSheetId();

    var modifiedColumnHValue = modifyUrlForPreview(columnHValue);
    var modifiedColumnIValue = convertToEmbedUrl(columnIValue);

    currentSheet.getRange('H2').setValue(modifiedColumnHValue);
    currentSheet.getRange('I2').setValue(modifiedColumnIValue);

    var dataRange = masterSheet.getDataRange();
    var values = dataRange.getValues();
    var found = false;

    for (var i = 1; i < values.length; i++) {
        var rowUrl = values[i][4]; // Assuming URLs are in column E
        if (rowUrl === sheetUrl) {
            masterSheet.getRange(i + 1, 8).setValue(modifiedColumnHValue); // Column H for audio link
            masterSheet.getRange(i + 1, 9).setValue(modifiedColumnIValue); // Column I for video link
            found = true;
            break;
        }
    }

    if (!found) {
        Logger.log('Sheet URL not found in master sheet');
    }
}

function modifyUrlForPreview(url) {
    // Check if the URL is empty or only contains whitespace, or does not include drive
  if (!url || url.trim() === '' || !url.includes('drive')) {
    return url; // Return the original URL without modification
  }
  
  // Normalize the URL by ensuring there's no trailing slash, which could interfere with replacements
  url = url.replace(/\/+$/, '');

  // Replace '/view?usp=sharing' or '/view?usp=drive_link' with '/preview', regardless of position
  url = url.replace(/\/view\?usp=sharing/, '/preview');
  url = url.replace(/\/view\?usp=drive_link/, '/preview');

  // Check if the URL ends with '/view' or '/view/' and replace it with '/preview'
  if (url.endsWith('/view/') || url.endsWith('/view')) {
    url = url.replace(/\/view\/?$/, '/preview');
  }

  // For URLs with 'open?id=', replace with 'file/d/' and append '/preview', avoiding double slashes
  else if (url.includes('open?id=')) {
    url = url.replace('open?id=', 'file/d/');
    // Append '/preview' if not already present and the URL doesn't end with '/', ensuring no double slashes
    if (!url.endsWith('/preview') && !url.endsWith('/')) {
      url += '/preview';
    }
  }

   // Append '/preview' if none of the above conditions applied and it's not already there
  else if (!url.endsWith('/preview')) {
    url += '/preview';
  }

  return url;
}

function getTitleTrack() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master [OST 2022]");
  var range = sheet.getRange("B2:B" + sheet.getLastRow());
  var values = range.getValues();
  var tracks = []; // Use a more descriptive variable name like 'tracks' instead of 'links'

  values.forEach(function(row) {
    // Check if row value does not contain the symbol "】"
    if (row[0] && !row[0].includes("】")) {
      tracks.push(row[0]);
    }
  });
  return tracks;
}


function fetchDramaNameForSong(columnBValue) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master [OST 2022]");
  var rows = sheet.getDataRange().getValues();
  var dramaName = "";

  for (var i = 0; i < rows.length; i++) {
    var currentBValue = rows[i][1].toString().trim();
    // If the current row's B column matches the selected song title
    if (currentBValue === columnBValue.trim()) {
      // Loop backwards to find the nearest preceding drama name, indicated by "】"
      for (var j = i; j >= 0; j--) {
        if (rows[j][1].includes("】")) {
          dramaName = rows[j][1];
          break; // Exit the loop once the drama name is found
        }
      }
      break; // Exit the loop once the matching song and its drama name are found
    }
  }
  return dramaName;
}

function getRowDataByColumnBValue(columnBValue) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master [OST 2022]");
    var rows = sheet.getDataRange().getValues();
    var result = { rowData: null };

    for (var i = 0; i < rows.length; i++) {
      var currentBValue = rows[i][1].toString().trim();
      if (currentBValue === columnBValue.trim()) {
        result.rowData = {
          columnC: rows[i][2],
          columnD: rows[i][3],
          columnE: rows[i][4],
          columnF: rows[i][5],
          columnH: rows[i][7],
          columnI: rows[i][8]
        };
        break; // Exit the loop once the song details are found
      }
    }

    Logger.log("Song details (excluding drama name): " + JSON.stringify(result));
    return result;
  } catch (e) {
    Logger.log("Error: " + e.toString());
    return null;
  }
}

function getDialogRowData(songTrack) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master [OST 2022]");
  var data = sheet.getRange("B2:I" + sheet.getLastRow()).getValues(); // Assuming H and I contain the links
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === songTrack) { // Column B matches songTrack
      return {audioLink: data[i][6], videoLink: data[i][7]}; // Columns H and I
    }
  }
  return {audioLink: "", videoLink: ""}; // Return empty strings if not found
}

// This function fetches the list of valid options for Column D from the data validation rules.
// Adjust the range as necessary to match where your data validation is set up.
function getColumnDOptions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master [OST 2022]");
  var range = sheet.getRange("D2:D"); // Adjust this range as needed
  var rules = range.getDataValidations();
  var options = [];

  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i][0]; // Assuming single-column range
    if (rule != null) {
      var criteria = rule.getCriteriaType();
      var args = rule.getCriteriaValues();
      if (criteria == SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        options = args[0];
        break; // Assuming all cells in the column have the same validation rule
      }
    }
  }
  return options;
}

function updateMasterSheetAndSync(columnBValue, columnDValue, columnEValue, columnFValue, columnHValue, columnIValue) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Master [OST 2022]");
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    // Fix variable names here: columnHValue and columnIValue instead of ColumnHValue and ColumnIValue
    var modifiedColumnHValue = modifyUrlForPreview(columnHValue);
    var modifiedColumnIValue = convertToEmbedUrl(columnIValue);
  
    // Update the master sheet
    for (var i = 0; i < values.length; i++) {
        if (values[i][1] == columnBValue) {
            sheet.getRange(i + 1, 4).setValue(columnDValue);
            sheet.getRange(i + 1, 5).setValue(columnEValue);
            sheet.getRange(i + 1, 6).setValue(columnFValue);
            sheet.getRange(i + 1, 8).setValue(modifiedColumnHValue);
            sheet.getRange(i + 1, 9).setValue(modifiedColumnIValue);
            break; // Exit after updating
        }
    }
  
    // Synchronize other sheets based on URLs in Column E
    values.forEach(function(row) {
        var url = row[4]; // Column E contains the URLs
        if (url) {
            var sheetId = extractSheetIdFromUrl(url);
            var targetSheet = findSheetById(spreadsheet, sheetId);
            if (targetSheet) {
                // Find the row in target sheet to update. This might require custom logic.
                targetSheet.getRange("H2").setValue(modifiedColumnHValue); // Update Column H in target sheet
                targetSheet.getRange("I2").setValue(modifiedColumnIValue); // Update Column I in target sheet
            }
        }
    });
}

function getSidebarPlaceholder() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = spreadsheet.getActiveSheet();
  var masterSheet = spreadsheet.getSheetByName("Master [OST 2022]");
  var isMasterSheet = activeSheet.getName() === "Master [OST 2022]";
  var selectedSong = isMasterSheet ? null : determineSelectedSongForSheet();
  var placeholder = isMasterSheet ? "Pick a song bruh" : (selectedSong || "Select a song");
  var data = getTitleTrack();

  if (!isMasterSheet) {
    var activeSheetUrl = activeSheet.getFormUrl(); // Assuming there's a way to get the unique URL or ID of the active sheet
    var urlsRange = masterSheet.getRange("E2:E" + masterSheet.getLastRow());
    var urls = urlsRange.getValues();

    for (var i = 0; i < urls.length; i++) {
      if (urls[i][0] === activeSheetUrl) {
        selectedSong = data[i];
        break;
      }
    }
  }

    // Logging to help debug the behavior
  Logger.log("Sheet Name: " + activeSheet.getName());
  Logger.log("Is Master Sheet: " + isMasterSheet);
  Logger.log("Selected Song: " + selectedSong);
  Logger.log("Placeholder: " + placeholder);

     return {
        data: data, // The data for the dropdown options
        placeholder: placeholder,
        selectedSong: selectedSong, // The song that should be selected by default
        isMasterSheet: isMasterSheet // Boolean indicating if the current sheet is the Master sheet
    };
}

function determineSelectedSongForSheet() {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var isMasterSheet = activeSheet.getName() === "Master [OST 2022]";

  if (isMasterSheet) {
    // Possibly use getMasterSheetData() for specific operations within the Master sheet
  } else {
    var activeSheetId = activeSheet.getSheetId().toString();
    var songSheetData = getSongSheetData();
    var selectedSong = songSheetData.find(item => item.sheetId === activeSheetId)?.songName;

    // Log for debugging
    Logger.log("Selected Song for sheet '" + activeSheet.getName() + "': " + selectedSong);

    return selectedSong;
  }
}

function isMasterSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheetName = spreadsheet.getActiveSheet().getName();
  return activeSheetName === "Master [OST 2022]";
}

function findSheetByIdNew(sheetId) {
  // Directly access the active spreadsheet without passing it as an argument
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  
  // Convert sheetId to a number if it's passed as a string
  var numericSheetId = parseInt(sheetId, 10);

  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === numericSheetId) {
      return sheets[i];
    }
  }
  return null; // Return null if no matching sheet is found
}

function getSongSheetData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = spreadsheet.getSheetByName("Master [OST 2022]");
  var data = [];
  var columnEValues = masterSheet.getRange('E2:E' + masterSheet.getLastRow()).getValues(); // URLs
  var columnBValues = masterSheet.getRange('B2:B' + masterSheet.getLastRow()).getValues(); // Song names

  columnEValues.forEach(function(eValue, i) {
    if (eValue[0] !== '') {
      var sheetId = extractSheetIdFromUrl(eValue[0]);
      if (sheetId) {
        data.push({ songName: columnBValues[i][0], sheetUrl: eValue[0], sheetId: sheetId }); // Include song name, URL, and sheetId
      }
    }
  });

  return data;
}


function matchingUrlforNumbering() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = spreadsheet.getSheetByName("Master [OST 2022]");
  var columnBRange = masterSheet.getRange('B2:B' + masterSheet.getLastRow());
  var columnERange = masterSheet.getRange('E2:E' + masterSheet.getLastRow());
  var columnGRange = masterSheet.getRange('G2:G' + masterSheet.getLastRow());
  var columnBValues = columnBRange.getValues();
  var columnEValues = columnERange.getValues();
  var columnGValues = columnGRange.getValues(); // Retrieve sheet names directly from column G

  var songData = [];
  var sequenceNumber = 0; // Start with 0, increment before assignment to start from 1

  columnEValues.forEach(function(eValue, index) {
    var url = eValue[0]; // Get the URL
    var title = columnBValues[index][0]; // Get the title
    var sheetName = columnGValues[index][0]; // Directly use the sheet name from column G
    var isDramaTitle = title.toLowerCase().includes("】"); // Check if title contains "】"

    if (url && !isDramaTitle) { // If there's a URL and it's not a drama title, proceed
      sequenceNumber++; // Increment sequence number for each valid URL entry
      var sheetId = extractSheetIdFromUrl(url); // Extract the sheet ID from the URL

      // Push entry with sequence number, song name from column B, and sheet ID, including the directly retrieved sheet name
      songData.push({
        sequenceNumber: sequenceNumber,
        sheetName: sheetName,
        songName: title,
        sheetId: sheetId,
      });
    }
  });
   // Store songData in Script Properties
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("songData", JSON.stringify(songData));

  Logger.log(songData);
  return songData;
}

// Function to store songData in Script Properties
function storeSongData() {
  var songData = matchingUrlforNumbering();
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("songData", JSON.stringify(songData));
}

// Function to retrieve songData from Script Properties
function retrieveSongData() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var songDataString = scriptProperties.getProperty("songData");
  if (songDataString) {
    return JSON.parse(songDataString);
  }
  return []; // Return an empty array if no data is found
}


function updateNumberingInfo(sheetId, firstNumber, lastNumber) {
  if (!sheetId) {
    console.error("No sheetId provided to updateNumberingInfo function.");
    return;
  }

  // Retrieve the stored songData from script properties
  var scriptProperties = PropertiesService.getScriptProperties();
  var songDataJson = scriptProperties.getProperty("songData");
  var songData = songDataJson ? JSON.parse(songDataJson) : [];
  
  // Find the sheet information by sheetId
  var sheetInfo = songData.find(item => item.sheetId == sheetId);
  if (!sheetInfo) {
    console.error("Sheet name not found for sheetId: " + sheetId);
    return;
  }

  // Use the found sheet name and sheetId to update the numbering info
  var numberingInfoJson = scriptProperties.getProperty("numberingInfo");
  var numberingInfo = numberingInfoJson ? JSON.parse(numberingInfoJson) : {};
  
  numberingInfo[sheetId] = {
    sheetName: sheetInfo.sheetName, // Use the sheet name from the found sheet info
    firstNumber: firstNumber,
    lastNumber: lastNumber
  };

  // Store the updated numbering info back into script properties
  scriptProperties.setProperty("numberingInfo", JSON.stringify(numberingInfo));

  console.log(`Numbering info updated for ${sheetInfo.sheetName} (Sheet ID: ${sheetId})`);
}

function getStoredNumberingInfo() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var numberingInfoString = scriptProperties.getProperty('numberingInfo');
  var numberingInfo = numberingInfoString ? JSON.parse(numberingInfoString) : {};
  return numberingInfo;
}

function numberTracksAcrossSheetsBasedOnMaster() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  const activeSheetName = activeSheet.getName();
  let numberingInfo = getStoredNumberingInfo();
  const songData = retrieveSongData(); // Assuming this retrieves your song data correctly
  const currentSheetData = songData.find(item => item.sheetName === activeSheetName);

  if (!currentSheetData) {
    console.log("Active sheet is not part of the sequence.");
    return;
  }

  const sheetId = currentSheetData.sheetId;
  let startingNumber = findStartingNumberFromContent(activeSheet, sheetId, songData);
  const lastRow = activeSheet.getLastRow();
  const columnDValues = activeSheet.getRange('D2:D' + lastRow).getValues();
  let numberingValues = [];
  
  columnDValues.forEach((cell, index) => {
    if (cell[0] !== '') {
      numberingValues.push([startingNumber++]);
    } else {
      numberingValues.push(['']); // Keep empty cells as is
    }
  });
  
  if (numberingValues.length > 0) {
    activeSheet.getRange(2, 3, numberingValues.length, 1).setValues(numberingValues);
  }

  // Assuming updateNumberingInfo correctly updates the numbering information
  updateNumberingInfo(sheetId, startingNumber - numberingValues.length, startingNumber - 1);

  // Update the stored numbering info to reflect this change
  PropertiesService.getScriptProperties().setProperty('numberingInfo', JSON.stringify(numberingInfo));

  console.log("Last number used:", startingNumber - 1);
}

function findStartingNumberFromContent(sheet, sheetId, songData) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Determine the current sheet's index in the sequence
  const currentIndex = songData.findIndex(data => data.sheetId === sheetId);

  // For the first sheet in the sequence, start numbering from 1
  if (currentIndex === 0) {
    return 1;
  }

  // For subsequent sheets, find the last number used in the previous sheet
  if (currentIndex > 0) {
    const previousSheet = spreadsheet.getSheetByName(songData[currentIndex - 1].sheetName);
    const lastRowOfPreviousSheet = findLastRowWithDataInColumnD(previousSheet);
    const lastNumberCell = previousSheet.getRange('C' + lastRowOfPreviousSheet).getValue();

    // Ensure the last number from the previous sheet is a number before returning the next number
    if (!isNaN(lastNumberCell) && lastNumberCell) {
      return parseInt(lastNumberCell) + 1;
    }
  }

  // Fallback if unable to determine from content, though in practice this should not happen
  Logger.log('Unable to determine starting number from content for sheet: ' + sheet.getName());
  return 1; // Default to 1 or consider a different fallback mechanism
}


function findStartingNumberForSheet(sheetId) {
  const numberingInfo = getStoredNumberingInfo();
  const songData = JSON.parse(PropertiesService.getScriptProperties().getProperty("songData") || "[]");
  
  // Find the index of the current sheet in the songData array
  const currentIndex = songData.findIndex(data => data.sheetId === sheetId);
  if (currentIndex <= 0) {
    // If it's the first sheet or not found, start from 1
    return 1;
  }
  
  // Find the previous sheet in the sequence
  const previousSheetId = songData[currentIndex - 1].sheetId;
  
  if (numberingInfo.hasOwnProperty(previousSheetId)) {
    const previousSheetInfo = numberingInfo[previousSheetId];
    // Return the next number after the last number of the previous sheet
    return previousSheetInfo.lastNumber + 1;
  }

  // Fallback if previous sheet info is missing
  Logger.log('No valid numbering info found for previous sheet in sequence. Starting from 1 or requires manual check.');
  return 1;
}


function findLastRowWithDataInColumnD(sheet) {
  // Use getLastRow to find the last row with any data in the sheet
  const lastDataRow = sheet.getLastRow();
  
  // If there's no data at all, return early to avoid unnecessary processing
  if (lastDataRow === 0) {
    return 2; // Assuming row 1 might have headers and data starts from row 2
  }
  
  // Adjust the range to only include up to the last row with data
  var columnDData = sheet.getRange('D2:D' + lastDataRow).getValues();
  
  // Iterate backwards through column D data to find the last non-empty cell
  for (var i = columnDData.length - 1; i >= 0; i--) {
    if (columnDData[i][0] !== "") {
      return i + 2; // +2 to adjust for zero-based indexing and starting from row 2
    }
  }
  
  // If we find no data in column D within the rows considered, return 2 by default
  return 2;
}


function checkAndStartRenumbering() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const songData = retrieveSongData();
  let numberingInfo = getStoredNumberingInfo();
  
  var needRenumbering = false;
  var details = {};

  for (let i = 0; i < songData.length; i++) {
    const sheetData = songData[i];
    const sheet = spreadsheet.getSheetByName(sheetData.sheetName);
    if (!sheet) continue;

    const expectedStartNumber = i === 0 ? 1 : (numberingInfo[songData[i - 1].sheetId].lastNumber + 1);
    const actualStartNumber = sheet.getRange('C2').getValue();

    if (actualStartNumber != expectedStartNumber) {
      needRenumbering = true;
      details = {
        songName: sheetData.songName,
        sequenceNumber: sheetData.sequenceNumber
      };
      break;
    }
  }

  if (needRenumbering) {
    openRenumberingDialog(details);
  } else {
    openNoRenumberingDialog();
  }
}


function openRenumberingDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Renumbering Dialog')
  var details = sheetData.songName; sheetData.sequenceNumber
  html.songName = details.songName; // Assign song name to the template
  html.sequenceNumber = details.sequenceNumber; // Assign sequence number to the template
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Renumbering Check ⋘ 𝑙𝑜𝑎𝑑𝑖𝑛𝑔 𝑑𝑎𝑡𝑎... ⋙');
}


function openNoRenumberingDialog() {
  var html = HtmlService.createTemplateFromFile('No Renumbering Dialog').evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, 'Renumbering Check ⋘ 𝑙𝑜𝑎𝑑𝑖𝑛𝑔 𝑑𝑎𝑡𝑎... ⋙');
}

function renumberFollowingSheetsBatch(startIndex = 0) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const songData = retrieveSongData(); // Assumes this returns an ordered array of sheets for processing, already sorted by sequence.
  let numberingInfo = getStoredNumberingInfo(); // Retrieve current numbering info

  songData.slice(startIndex).forEach((sheetData, index) => {
    const sheet = spreadsheet.getSheetByName(sheetData.sheetName);
    if (!sheet) {
      Logger.log('Sheet not found: ' + sheetData.sheetName);
      return; // Skip if the sheet doesn't exist.
    }

    const lastRow = findLastRowWithDataInColumnD(sheet); // Determine the last row with data in column D.
    let startNumber = findStartingNumberForSheet(sheetData.sheetId, numberingInfo, songData);
    let renumberData = []; // Initialize an array to hold the new numbering.

    for (let row = 2; row <= lastRow; row++) { // Start from row 2 assuming row 1 has headers.
      const cellValue = sheet.getRange('D' + row).getValue();
      if (cellValue !== '') { // Only number rows with data in column D.
        renumberData.push([startNumber++]);
      } else {
        renumberData.push(['']); // Keep empty rows unnumbered.
      }
    }

    // Batch update the numbering in column C
    if (renumberData.length > 0) {
      sheet.getRange(2, 3, renumberData.length, 1).setValues(renumberData);
    }

    // Always update numbering info for the current sheet
    numberingInfo[sheetData.sheetId] = { firstNumber: renumberData[0][0], lastNumber: startNumber - 1 };
    console.log(`Sheet "${sheetData.sheetName}" renumbered. Start number: ${renumberData[0][0]}, End number: ${startNumber - 1}`);
    Logger.log(`Sheet "${sheetData.sheetName}" renumbered. Start number: ${renumberData[0][0]}, End number: ${startNumber - 1}`);
  });

  // After renumbering, update the stored numbering info
  PropertiesService.getScriptProperties().setProperty('numberingInfo', JSON.stringify(numberingInfo));
  Logger.log('Batch renumbering process completed starting from index: ' + startIndex);
}

function compactColumnD(sheet) {
  if (!sheet) {
    Logger.log("Sheet is undefined.");
    return; // Exit the function if no sheet is provided.
  }

  // Log the name of the sheet being processed.
  Logger.log("Compacting columns D and E on sheet: " + sheet.getName());

  // Find the last row with content in column D.
  const lastRowD = findLastRowWithDataInColumnD(sheet);
  Logger.log("Last row with content in column D: " + lastRowD);

  // Fetch the ranges for column D and E up to the last row with content.
  const rangeD = sheet.getRange('D2:D' + lastRowD);
  const valuesD = rangeD.getValues();

  let compactedValuesD = [];
  // Loop through column D and E values and push non-empty cells to the compacted arrays.
  valuesD.forEach((value, index) => {
    if (value[0] !== '') {
      compactedValuesD.push([value[0]]);
    }
  });

  // Log the lengths of the compacted arrays for verification.
  Logger.log("Number of compacted rows for column D: " + compactedValuesD.length);

  // Clear original data ranges in columns D and E
  rangeD.clearContent();

  // Write back the compacted values to columns D and E, starting from the top.
  if (compactedValuesD.length > 0) {
    sheet.getRange('D2:D' + (1 + compactedValuesD.length)).setValues(compactedValuesD);
  }
}

function trimColumnD() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var column = "D";
  var startRow = 1;
  var range = sheet.getRange(column + startRow + ":" + column + sheet.getLastRow());
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++) {
    values[i][0] = values[i][0].toString().trim();
  }
  
  range.setValues(values);
}

function alignTranslationColumn(sheet) {
  if (!sheet) {
    Logger.log("Sheet is undefined.");
    return; // Exit the function if no sheet is provided.
  }

  Logger.log("Aligning translations in Column F with Column D on sheet: " + sheet.getName());

  // Determine the last row to process by finding the maximum last row of both columns D and F.
  const lastRowD = findLastRowWithDataInColumnD(sheet);
  const lastRowF = sheet.getLastRow(); // Assuming you want to process all rows that have data in Column F.
  const lastRow = Math.max(lastRowD, lastRowF);

  // Fetch the ranges and values for columns D and F.
  const rangeD = sheet.getRange('D2:D' + lastRow);
  const valuesD = rangeD.getValues();
  const rangeF = sheet.getRange('F2:F' + lastRow);
  const valuesF = rangeF.getValues();

  // Initialize arrays to hold the new aligned values for Column F.
  let alignedValuesF = [];
  let currentIndexF = 0; // Track the current index for non-empty cells in Column F.

  // Loop through the values in Column D to align Column F.
  valuesD.forEach((valueD, index) => {
    if (valueD[0] !== '') {
      // If there's content in Column D, align the translation from Column F.
      alignedValuesF.push(valuesF[currentIndexF]);
      currentIndexF++; // Move to the next translation in Column F.
    } else {
      // If Column D is empty at this row, add an empty cell for Column F alignment.
      alignedValuesF.push(['']);
    }
  });

  // Clear the original data range in Column F.
  rangeF.clearContent();

  // Write back the aligned values to Column F.
  if (alignedValuesF.length > 0) {
    sheet.getRange('F2:F' + alignedValuesF.length).setValues(alignedValuesF);
  }

  Logger.log("Translations in Column F aligned with Column D.");
}
