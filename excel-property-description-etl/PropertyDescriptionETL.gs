// =============================
// FINAL CLEAN RELIABLE SCRIPT
// =============================

var globalFolderId = "";

// ========== ON OPEN MENU ==========
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Data Processing')
    .addItem('Start Processing', 'confirmProcessing')
    .addItem('Continue from Excel Link', 'continueFromExcelLink') // ✅ Added back
    .addItem('Start Time-based Trigger', 'confirmAndSetTrigger')
    .addItem('Stop Time-based Trigger', 'stopTimeBasedTrigger')
    .addSeparator()
    .addItem('Start CM Matched Processing', 'confirmProcessingMatchingCommunityMinerals')
    .addItem('Start CM Time-based Trigger', 'confirmAndSetTriggerForCommunityMinerals')
    .addItem('Stop CM Time-based Trigger', 'stopTimeBasedTriggerForCommunityMinerals')
    .addSeparator()
    .addItem('Remove All Log Sheets', 'removeAllLogSheets')
    .addToUi();
}


// ========== TIME MANAGEMENT ==========
function isTimeExceeding(startTime) {
  const now = Date.now();
  const elapsed = now - startTime;
  const timeLimit = 5 * 60 * 1000; // 5 min
  const buffer = 30 * 1000;         // 30 sec
  const remaining = timeLimit - elapsed;
  if (remaining <= buffer) {
    logToSheet(`⚠️ Aborting to avoid timeout. Remaining ms: ${remaining}`);
    return true;
  }
  return false;
}

// ========== LOGGING ==========
var logSheet = null;
function logToSheet(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!logSheet) {
    const timestamp = new Date().toISOString().replace(/[^0-9]/g, "").slice(0, 14);
    try {
      logSheet = ss.insertSheet("Logs_" + timestamp);
    } catch (e) {
      logSheet = ss.insertSheet("Logs_" + timestamp + "_" + Math.floor(Math.random() * 1000));
    }
    logSheet.appendRow(["Timestamp", "Log Message"]);

    // Redefine to append directly after initialization
    logToSheet = function (msg) {
      logSheet.appendRow([new Date(), msg]);
    };
  }
  logSheet.appendRow([new Date(), message]);
}

// ========== DATA HANDLING ==========
function getProcessedFileDataFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const processedFilesData = {};

  for (let i = 1; i < data.length; i++) {
    const fileName = data[i][1];
    const description = data[i][0];
    if (!fileName || !description) continue;

    const cleanedFileName = fileName
      .replace(/^Copy of /i, "")
      .replace(/\.(xlsx|xls|pdf)$/i, "")
      .trim();

    if (!cleanedFileName) continue;

    if (!processedFilesData[cleanedFileName]) processedFilesData[cleanedFileName] = new Set();
    processedFilesData[cleanedFileName].add(description);
  }
  return processedFilesData;
}

function writeUniqueDataToSheet(uniqueDescriptions, fileName) {
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Data");
  if (dataSheet.getLastRow() === 0) dataSheet.appendRow(['Unique Property Descriptions', 'Processed File']);
  if (uniqueDescriptions.length > 0) {
    const rows = uniqueDescriptions.map(desc => [desc, fileName]);
    dataSheet.getRange(dataSheet.getLastRow() + 1, 1, rows.length, 2).setValues(rows);
    logToSheet(`✅ Written ${uniqueDescriptions.length} entries from ${fileName}.`);
  } else {
    logToSheet(`ℹ️ No entries to write for ${fileName}.`);
  }
}

// ========== GENERIC PIPELINE ==========
function confirmProcessing() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter Folder ID', 'Input Google Drive Folder ID:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    globalFolderId = response.getResponseText().trim();
    PropertiesService.getScriptProperties().setProperty('folderId', globalFolderId);
    processFolderGeneric(globalFolderId);
  } else {
    ui.alert('Processing cancelled.');
  }
}

function processFolderGeneric(folderId) {
  const startTime = Date.now();
  logToSheet(`🚀 Starting generic processing for folder: ${folderId}`);

  const processedFilesData = getProcessedFileDataFromSheet();
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const fileList = [];
  while (files.hasNext()) fileList.push(files.next());
  fileList.sort((a, b) => a.getName().localeCompare(b.getName()));

  let processedCount = 0;
  const MAX_FILES_PER_RUN = 10;

  for (let file of fileList) {
    if (isTimeExceeding(startTime)) break;

    const fileName = file.getName();
    if (processedFilesData[fileName]) continue;

    try {
      processAndWriteExcelFile(file, processedFilesData, startTime);
      processedFilesData[fileName] = true;
      processedCount++;
    } catch (e) {
      logToSheet(`❌ Error processing ${fileName}: ${e.message}`);
    }

    if (processedCount >= MAX_FILES_PER_RUN) {
      logToSheet(`⏸️ Max files reached (${MAX_FILES_PER_RUN}). Stopping run.`);
      break;
    }
  }

  logToSheet(`✅ Generic processing complete. Files processed: ${processedCount}`);
}

function continueFromExcelLink() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Paste Google Drive link or file ID of the Excel file you wish to process:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Operation cancelled.');
    return;
  }

  let urlOrId = response.getResponseText().trim();
  if (!urlOrId) {
    ui.alert('❌ No link or ID entered.');
    return;
  }

  let fileId = urlOrId;
  try {
    // Extract ID if a full URL is pasted
    const match = urlOrId.match(/[-\w]{25,}/);
    if (match) fileId = match[0];
  } catch (e) {
    ui.alert('❌ Could not parse the file ID from your input.');
    return;
  }

  try {
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    logToSheet(`📄 Found file: ${fileName} (${fileId}). Starting conversion to Google Sheets...`);

    // Convert Excel to Google Sheets
    const blob = file.getBlob();
    const tempFileResource = {
      title: fileName.replace(/\.xlsx$/i, '').replace(/\.xls$/i, ''),
      mimeType: MimeType.GOOGLE_SHEETS
    };

    let newSheetFile;
    let retries = 5;
    while (retries > 0) {
      try {
        newSheetFile = Drive.Files.insert(tempFileResource, blob);
        break;
      } catch (e) {
        retries--;
        Utilities.sleep(2000);
        if (retries === 0) {
          logToSheet(`❌ Failed to convert ${fileName}. ${e.message}`);
          ui.alert(`❌ Failed to convert ${fileName}. ${e.message}`);
          return;
        }
      }
    }

    logToSheet(`✅ Conversion complete: ${newSheetFile.id}. Opening as Google Sheet...`);
    const spreadsheet = SpreadsheetApp.openById(newSheetFile.id);
    const sheet = spreadsheet.getSheets()[0];

    // Pull data for processing
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let idx = headers.indexOf("Property Description");
    if (idx === -1) idx = headers.indexOf("Property Name");
    if (idx === -1) idx = headers.indexOf("CC Description");
    if (idx === -1) idx = headers.indexOf("Property name");

    if (idx === -1) {
      logToSheet(`ℹ️ No relevant column found in ${fileName}. Skipping.`);
      DriveApp.getFileById(newSheetFile.id).setTrashed(true);
      ui.alert(`ℹ️ No relevant column found in ${fileName}. Skipped.`);
      return;
    }

    const uniqueDescriptions = new Set();
    for (let i = 1; i < data.length; i++) {
      const desc = data[i][idx];
      if (desc) uniqueDescriptions.add(desc.trim());
    }

    if (uniqueDescriptions.size > 0) {
      writeUniqueDataToSheet(Array.from(uniqueDescriptions), fileName);
    } else {
      logToSheet(`ℹ️ No unique entries found in ${fileName}.`);
      ui.alert(`ℹ️ No unique entries found in ${fileName}.`);
    }

    DriveApp.getFileById(newSheetFile.id).setTrashed(true);
    logToSheet(`✅ Finished processing and trashed temporary Google Sheet for ${fileName}.`);
    ui.alert(`✅ Successfully processed: ${fileName}.`);

  } catch (e) {
    const errorMsg = `❌ Error processing file: ${e.message}`;
    logToSheet(errorMsg);
    ui.alert(errorMsg);
  }
}

// ========== CM PIPELINE ==========
function confirmProcessingMatchingCommunityMinerals() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter Folder ID', 'Input Google Drive Folder ID for CM Processing:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const folderId = response.getResponseText().trim();
    const cmSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Community Minerals_DATA");
    if (!cmSheet) {
      ui.alert("❌ 'Community Minerals_DATA' sheet not found.");
      return;
    }

    // ❌ REMOVE checkpoint check and popup completely
    // ✅ No need to clear Column B checkpoints since we now use internal continuation

    PropertiesService.getScriptProperties().setProperty('folderId', folderId);
    processFolderWithCommunityMineralsMatch(folderId, false);
  } else {
    ui.alert('Processing cancelled.');
  }
}

function confirmAndSetTriggerForCommunityMinerals() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter Folder ID', 'Input Google Drive Folder ID for CM Time-based Processing:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('❌ Operation cancelled.');
    return;
  }

  const folderId = response.getResponseText().trim();
  if (!folderId) {
    ui.alert('❌ No folder ID entered. Operation cancelled.');
    return;
  }

  // Save the folder ID and log it
  PropertiesService.getScriptProperties().setProperty('folderId', folderId);
  logToSheet(`✅ Folder ID saved for CM processing: ${folderId}`);

  // Run the first batch immediately
  logToSheet("⏳ Starting immediate CM run before setting time-based trigger.");
  processFolderWithCommunityMineralsMatch(folderId, true);

  // Create the 5-minute repeating trigger
  createTriggerForCommunityMinerals();

  ui.alert("✅ CM time-based trigger created and initial run completed.");
}

// ========== CM CORE FUNCTION ==========
function processFolderWithCommunityMineralsMatch(folderId, isTrigger) {
  const startTime = Date.now();
  logToSheet(`🚀 Starting CM processing. Trigger run: ${isTrigger}`);

  if (isTimeExceeding(startTime)) return;

  const processedFilesData = getProcessedFileDataFromSheet();
  logToSheet(`✅ Retrieved processed file data.`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cmSheet = ss.getSheetByName("Community Minerals_DATA");
  if (!cmSheet) {
    logToSheet("❌ 'Community Minerals_DATA' sheet not found.");
    return;
  }
  logToSheet(`✅ Found CM Sheet.`);

  const cmData = cmSheet.getRange(2, 1, cmSheet.getLastRow() - 1, 1).getValues();
  const cmFileNamesSet = new Set(cmData.map(row =>
    row[0]
      ?.replace(/^Copy of /i, "")
      ?.replace(/\.(xlsx|xls|pdf)$/i, "")
      ?.trim()
  ).filter(Boolean));
  logToSheet(`✅ Loaded ${cmFileNamesSet.size} unique CM file names for matching.`);

  const rootFolder = DriveApp.getFolderById(folderId);
  logToSheet(`✅ Opened root folder: ${rootFolder.getName()} (${rootFolder.getId()})`);

  const topLevelFolders = [];
  const folderIterator = rootFolder.getFolders();
  while (folderIterator.hasNext()) {
    const folder = folderIterator.next();
    topLevelFolders.push({
      folder: folder,
      name: folder.getName(),
      id: folder.getId()
    });
  }
  topLevelFolders.sort((a, b) => a.name.localeCompare(b.name));

  // If no folders inside, process files directly
  if (topLevelFolders.length === 0) {
    logToSheet(`⚠️ No folders found. Processing files directly in "${rootFolder.getName()}"`);
    processFilesInFolderForCM(rootFolder, cmFileNamesSet, processedFilesData, startTime, 10);
    logToSheet(`✅ Run complete (root-only). Trigger active: ${isTrigger}`);
    return;
  }

  let lastProcessedFolderId = isTrigger ? PropertiesService.getScriptProperties().getProperty('lastProcessedFolderId_CM') : null;
  let resumeIndex = 0;

  if (isTrigger && lastProcessedFolderId) {
    const index = topLevelFolders.findIndex(f => f.id === lastProcessedFolderId);
    if (index !== -1) {
      resumeIndex = index;
      logToSheet(`📌 Resuming from folder: ${topLevelFolders[resumeIndex].name}`);
    } else {
      logToSheet(`⚠️ Last processed folder ID not found. Starting from top.`);
    }
  } else {
    logToSheet(`🧭 No previous folder found. Starting from top.`);
  }

  let processedCount = 0;
  const MAX_FILES_PER_RUN = 10;

  for (let i = resumeIndex; i < topLevelFolders.length; i++) {
    if (isTimeExceeding(startTime)) break;

    const folder = topLevelFolders[i].folder;
    const folderName = folder.getName();
    const folderIdCurrent = folder.getId();

    logToSheet(`📁 Processing folder: ${folderName}`);
    PropertiesService.getScriptProperties().setProperty('lastProcessedFolderId_CM', folderIdCurrent);

    const processedInFolder = processFilesInFolderForCM(folder, cmFileNamesSet, processedFilesData, startTime, MAX_FILES_PER_RUN - processedCount);
    processedCount += processedInFolder;

    if (processedCount >= MAX_FILES_PER_RUN) {
      logToSheet(`⏸️ Max files (${MAX_FILES_PER_RUN}) reached. Pausing.`);
      return;
    }
  }
  logToSheet(`✅ CM processing complete. Files processed: ${processedCount}`);
}

function processFilesInFolderForCM(folder, cmFileNamesSet, processedFilesData, startTime, maxFiles = 10) {
  let processedInFolder = 0;
  let matchesInFolder = 0;

  const files = getAllXLSXFilesRecursively(folder);

  for (const file of files) {
    if (isTimeExceeding(startTime)) return processedInFolder;
    if (processedInFolder >= maxFiles) return processedInFolder;

    const fileName = file.getName();
    const normalizedName = fileName
      .replace(/^Copy of /i, "")
      .replace(/\.xlsx$/i, "")
      .replace(/\.xls$/i, "")
      .trim();

    const isInCMList = cmFileNamesSet.has(normalizedName);
    const isAlreadyProcessed = processedFilesData[normalizedName] ? true : false;

    logToSheet(`🔍 Checking "${fileName}" | In CM List: ${isInCMList} | Already Processed: ${isAlreadyProcessed}`);

    if (!isInCMList || isAlreadyProcessed) continue;

    try {
      processAndWriteExcelFile(file, processedFilesData, startTime);
      processedFilesData[normalizedName] = true;
      matchesInFolder++;
      processedInFolder++;
    } catch (e) {
      logToSheet(`❌ Error processing ${fileName}: ${e.message}`);
    }
  }

  if (matchesInFolder === 0) {
    logToSheet(`📁 Completed folder: ${folder.getName()} - No matches found.`);
  } else {
    logToSheet(`📁 Completed folder: ${folder.getName()} - ${matchesInFolder} matching files processed and written.`);
  }

  return processedInFolder;
}

function getAllXLSXFilesRecursively(folder) {
  const xlsxFiles = [];

  // Get files in current folder
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName().toLowerCase();
    if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      xlsxFiles.push(file);
    }
  }

  // Recurse into subfolders
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    xlsxFiles.push(...getAllXLSXFilesRecursively(subfolder));
  }

  return xlsxFiles;
}

// ========== TRIGGERS ==========
function createTriggerForCommunityMinerals() {
  deleteTriggerForCommunityMinerals();
  ScriptApp.newTrigger('triggerCMScheduled').timeBased().everyMinutes(5).create();
  logToSheet("✅ CM time-based trigger created.");
}

function deleteTriggerForCommunityMinerals() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'triggerCMScheduled') ScriptApp.deleteTrigger(t);
  });
}

function triggerCMScheduled() {
  const folderId = PropertiesService.getScriptProperties().getProperty('folderId');
  if (!folderId) {
    logToSheet("⚠️ No folder ID found for CM scheduled run. Skipping.");
    return;
  }
  processFolderWithCommunityMineralsMatch(folderId, true);
}

function stopTimeBasedTriggerForCommunityMinerals() {
  deleteTriggerForCommunityMinerals();
  SpreadsheetApp.getUi().alert('✅ CM time-based trigger stopped.');
}

// ========== EXCEL FILE PROCESSING ==========
function processAndWriteExcelFile(file, processedFilesData, startTime) {
  startTime = startTime || Date.now();
  if (isTimeExceeding(startTime)) return;

  logToSheet(`📄 Processing file: ${file.getName()}`);
  const blob = file.getBlob();
  const tempFile = {
    title: file.getName().replace(/\.xlsx$/i, '').replace(/\.xls$/i, ''),
    mimeType: MimeType.GOOGLE_SHEETS
  };

  let newSheetFile;
  let retries = 5;
  while (retries > 0) {
    try {
      newSheetFile = Drive.Files.insert(tempFile, blob);
      break;
    } catch (e) {
      retries--;
      Utilities.sleep(2000);
      if (retries === 0) {
        logToSheet(`❌ Failed to convert ${file.getName()}.`);
        return;
      }
    }
  }

  const spreadsheet = SpreadsheetApp.openById(newSheetFile.id);
  const sheet = spreadsheet.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  let idx = headers.indexOf("Property Description");
  if (idx === -1) idx = headers.indexOf("Property Name");
  if (idx === -1) idx = headers.indexOf("CC Description");
  if (idx === -1) idx = headers.indexOf("Property name");

  if (idx === -1) {
    logToSheet(`ℹ️ No relevant column found in ${file.getName()}. Skipping.`);
    DriveApp.getFileById(newSheetFile.id).setTrashed(true);
    return;
  }

  const uniqueDescriptions = new Set();
  for (let i = 1; i < data.length; i++) {
    const desc = data[i][idx];
    if (desc) uniqueDescriptions.add(desc.trim());
  }

  if (uniqueDescriptions.size > 0) {
    writeUniqueDataToSheet(Array.from(uniqueDescriptions), file.getName());
  } else {
    logToSheet(`ℹ️ No unique entries found in ${file.getName()}.`);
  }

  DriveApp.getFileById(newSheetFile.id).setTrashed(true);
  logToSheet(`✅ Finished processing ${file.getName()}.`);
}

// ========== CLEANUP ==========
function removeAllLogSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(s => {
    if (/^Logs_\d+$/.test(s.getName())) ss.deleteSheet(s);
  });
  SpreadsheetApp.getUi().alert("✅ All log sheets removed.");
}
