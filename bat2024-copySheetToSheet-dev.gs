const MAIN_FOLDER_ID = 
const SHEET_NAME = '2024-25 Roster';
const TEMPLATE_FILE_ID = 
const TEMPLATE_SHEET_NAME = '2023-24 Roster';
const ROSTERS_FOLDER_ID = 

function createFileFromTemplate() {
  const lock = LockService.getScriptLock();
  try {
    // Try to obtain a lock with a 5-minute timeout
    if (!lock.tryLock(300000)) {
      Logger.log('Could not obtain lock. Script is already running.');
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    Logger.log('Active sheet name: ' + sheet.getName());

    const dataRange = sheet.getRange('I2:J');
    const data = dataRange.getValues().filter(row => row[0].toString().trim() !== '');
    Logger.log('Filtered data from columns I and J: ' + JSON.stringify(data));

    let mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    let templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);

    const existingFiles = getExistingFiles(mainFolder);
    Logger.log('Existing files: ' + JSON.stringify(Object.keys(existingFiles)));

    data.forEach((row, index) => {
      const fileName = row[0].toString().trim() + ' Roster';
      Logger.log(`Processing file ${index + 1}/${data.length}: ${fileName}`);

      let newFile;
      let folder = mainFolder;
      if (existingFiles[fileName]) {
        newFile = existingFiles[fileName].file;
        folder = existingFiles[fileName].folder;
        Logger.log(`Existing file found: ${fileName} in folder: ${folder.getName()}`);
      } else {
        newFile = templateFile.makeCopy(fileName, folder);
        Logger.log(`New file created: ${fileName}`);
      }

      let newSs = SpreadsheetApp.open(newFile);
      let newSheet = newSs.getSheetByName(SHEET_NAME);

      if (!newSheet) {
        newSheet = createNewSheet(newSs);
        if (!newSheet) {
          Logger.log(`Error creating new sheet in: ${fileName}`);
          return;
        }
      }

      copyValuesToNewSheet(newSs, row[0].toString().trim());
    });

    Logger.log('Script completed successfully.');
  } catch (e) {
    Logger.log('Error: ' + e.toString());
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function getExistingFiles(folder) {
  const files = {};

  function searchFolder(folder) {
    const fileIterator = folder.getFiles();
    while (fileIterator.hasNext()) {
      const file = fileIterator.next();
      files[file.getName()] = { file: file, folder: folder };
    }

    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      searchFolder(subFolders.next());
    }
  }

  searchFolder(folder);
  return files;
}

function copyValuesToNewSheet(newSs, fileName) {
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sourceSheet.getLastRow();

  if (lastRow < 2) {
    Logger.log('No data to copy. Exiting function.');
    return;
  }

  const sourceValues = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
  const destinationSheet = newSs.getSheetByName(SHEET_NAME);

  if (!destinationSheet) {
    Logger.log('Destination sheet ' + SHEET_NAME + ' not found in new spreadsheet.');
    return;
  }

  // Get current data and check for summation row
  const destLastRow = destinationSheet.getLastRow();
  let existingValuesMap = new Map();
  let summationRow = null;

  if (destLastRow >= 6) {
    const existingValues = destinationSheet.getRange(6, 1, destLastRow - 5, destinationSheet.getLastColumn()).getValues();
    
    // Check if last row is a summation row
    const lastRowValues = existingValues[existingValues.length - 1];
    const isSummationRow = lastRowValues[0].toString().toLowerCase().includes('total');
    
    if (isSummationRow) {
      summationRow = lastRowValues;
      existingValues.pop(); // Remove summation row from processing
    }

    existingValues.forEach(row => {
      const key = `${row[3]}_${row[6]}`; // Using columns D and G as unique identifier
      existingValuesMap.set(key, true);
    });
  }

  Logger.log(`Copying data for file: ${fileName}`);
  Logger.log(`Source values length: ${sourceValues.length}`);

  const newValues = sourceValues
    .filter(row => row[8].toString().trim() === fileName) // Using column I (index 8) for filtering
    .filter(row => {
      const key = `${row[3]}_${row[6]}`; // Using columns D and G as unique identifier
      if (!existingValuesMap.has(key)) {
        existingValuesMap.set(key, true); // Mark as processed
        return true;
      }
      return false;
    })
    .map(row => [row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]]); // Including columns A to H

  Logger.log(`New values to be added: ${newValues.length}`);

  if (newValues.length > 0) {
    // If there was a summation row, remove it first
    if (summationRow) {
      destinationSheet.deleteRow(destLastRow);
    }

    // Add new data
    const startRow = Math.max(6, destLastRow + (summationRow ? 0 : 1));
    destinationSheet.getRange(startRow, 1, newValues.length, 8).setValues(newValues);
    
    // Update and append summation row
    const newLastRow = startRow + newValues.length;
    appendSummationRow(destinationSheet, newLastRow);
    
    Logger.log(`${newValues.length} new rows added to the '${SHEET_NAME}' sheet with updated summation.`);
  } else {
    Logger.log(`No new values to add to the '${SHEET_NAME}' sheet.`);
  }
}

function createNewSheet(spreadsheet) {
  try {
    const newSheet = spreadsheet.insertSheet(SHEET_NAME);
    // Add any necessary formatting or headers to the new sheet here
    return newSheet;
  } catch (e) {
    Logger.log('Error creating new sheet: ' + e.toString());
    return null;
  }
}

function appendSummationRow(sheet, lastDataRow) {
  const dataRange = sheet.getRange(6, 1, lastDataRow - 5, 15); // Extend range to column O
  const newLastRow = lastDataRow + 1;
  
  const summationRow = sheet.getRange(newLastRow, 1, 1, 15); // Extend range to column O
  
  summationRow.getCell(1, 1).setValue('Total');
  
  // Add SUM formula only for column O (index 15)
  const colO = `=SUM(${sheet.getRange(6, 15, lastDataRow - 5, 1).getA1Notation()})`;
  summationRow.getCell(1, 15).setFormula(colO);
  
  summationRow
    .setFontWeight('bold')
    .setBackground('#f3f3f3')
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}

// Optional: Add a function to manually trigger the script
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Roster Management')
    .addItem('Update Rosters', 'createFileFromTemplate')
    .addToUi();
}
