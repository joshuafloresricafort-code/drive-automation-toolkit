function onOpen() {
  // Create a custom menu in Google Sheets
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('File Management')
    .addItem('Remove "Copy of" from all files', 'removeCopyOfFromAllFiles')
    .addToUi();
}

function removeCopyOfFromAllFiles() {
  // Prompt user for the folder ID
  var folderId = SpreadsheetApp.getUi().prompt('Enter the Google Drive folder ID:').getResponseText();
  
  if (folderId !== '') {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles(); // Get all files, regardless of type
    
    while (files.hasNext()) {
      var file = files.next();
      var originalName = file.getName();
      
      // Check if the file name starts with 'Copy of'
      if (originalName.startsWith('Copy of ')) {
        var newName = originalName.replace('Copy of ', ''); // Remove 'Copy of '
        file.setName(newName); // Rename the file
        Logger.log('Renamed: ' + originalName + ' -> ' + newName);
      }
    }
    
    SpreadsheetApp.getUi().alert('Renaming complete.');
  } else {
    SpreadsheetApp.getUi().alert('No folder ID provided.');
  }
}
