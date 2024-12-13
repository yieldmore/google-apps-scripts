/*
  What this does:
    Lists all files and sub-folders from a folder in Google Drive.
    Assumes activeSheet >> parent is the folder to be scanned.

  Adapted from Code written by @hubgit https://gist.github.com/hubgit/3755293
  Updated since DocsList is deprecated  https://ctrlq.org/code/19854-list-files-in-google-drive-folder
  Bluexm: added recursion on subfolders - SO: https://webapps.stackexchange.com/a/142584

  Imran <team@yieldmore.org> in Dec 2024 has
   * introduced "onlyFolders" for own use
   * done a full cleanup of all variables
   * uses ' » ' as folder separator
   * added removeEmptyRows/Columns from Trey (SO)
   * added .getParents()[0].getName() per SO: https://stackoverflow.com/a/17618407
  TODO: publish to amadeus blog:
*/

var sheet = SpreadsheetApp.getActiveSheet();
var onlyFolders = sheet.getSheetName() == "Folders Only"

var parentFolder = Drive.Drives.get(sheet.getParent().getId())
var topFolderName = parentFolder.getName();
//var topFolderName = 'All YieldMore.org and Growth / Work'

// entry function 
function ScanFolder() {
  ScanFoldeRecursively(topFolderName, '')
  removeEmptyColumns()
  removeEmptyRows()
}

function ScanFoldeRecursively(folderName, relativeFolderName) {
  if (topFolderName == folderName) {
    // clear any existing content
    sheet.clearContents();
    // append a header row
    sheet.appendRow([
      "#Folder",
      "Name",
      //"Date Last Updated",
      //"Size",
      "URL",
      //"ID",
      "Description",
      "Type"
    ]);

    Logger.log("REBUILDING SHEET: " + sheet.getSheetName())
  }

  Logger.log("SCANNING: " + folderName) //DEBUG

  var folderId = DriveApp.getFoldersByName(folderName, { }).next().getId()
  var folder = Drive.Drives.get(folderId, { supportsAllDrives: true })

  // files is a File Iterator
  var files = onlyFolders ? false : folder.getFiles()

  var parentFolderPrefix = relativeFolderName == topFolderName ? '' : relativeFolderName + ' » '

  // loop through files in the folder
  while (files && files.hasNext()) {
    var item = files.next();

    var data = [
      parentFolderPrefix + folderName,
      item.getName(),
      //item.getLastUpdated(),
      //item.getSize(),
      item.getUrl(),
      //item.getId(),
      item.getDescription(),
      'FILE', //item.getBlob().getContentType(),
    ];

    sheet.appendRow(data);
  } // Completes listing of the files in the named folder

  var subFolders = onlyFolders ? folder.getFolders() : false //Drive.Drives.list({ driveId: folder.id, supportsAllDrives: true }) : false

  // now start a loop on the subFolders list
  while (onlyFolders && subFolders.hasNext()) {
    var item = subFolders.next();
    Logger.log("Subfolder name:" + item.getName()); //DEBUG

    if (!onlyFolders) Logger.log("FILES IN THIS FOLDER"); //DEBUG

    if (onlyFolders) {
      var data = [
        relativeFolderName + ' » ' + folderName,
        item.getName(),
        //item.getLastUpdated(),
        //item.getSize(),
        item.getUrl(),
        //item.getId(),
        item.getDescription(),
        'Google Folder',
      ];
      //Logger.log("data = " + data); //DEBUG
      sheet.appendRow(data);
    }

    ScanFoldeRecursively(item.getName(),
      (topFolderName == folderName ? '' : relativeFolderName + ' » ') + folderName)
  }
}

//NOT USED: https://arisazhar.com/remove-empty-rows-in-spreadsheet-instantly/

//FROM: https://stackoverflow.com/a/34781833
//Remove All Empty Columns in the Entire Workbook
function removeEmptyColumns() {
  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();
  for (var s in allsheets){
    var sheet=allsheets[s]
    var maxColumns = sheet.getMaxColumns(); 
    var lastColumn = sheet.getLastColumn();
    if (maxColumns-lastColumn != 0){
      sheet.deleteColumns(lastColumn+1, maxColumns-lastColumn);
    }
  }
}

//Remove All Empty Rows in the Entire Workbook
function removeEmptyRows() {
  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();
  for (var s in allsheets){
    var sheet=allsheets[s]
    var maxRows = sheet.getMaxRows(); 
    var lastRow = sheet.getLastRow();
    if (maxRows-lastRow != 0){
      sheet.deleteRows(lastRow+1, maxRows-lastRow);
    }
  }
}
