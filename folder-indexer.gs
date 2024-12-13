/*
  What this does:
    Lists all files and sub-folders from a folder in Google Drive.
    Assumes activeSheet >> parent is the [Shared] Folder to be scanned.

  Adapted from Code written by @hubgit https://gist.github.com/hubgit/3755293
  Updated since DocsList is deprecated  https://ctrlq.org/code/19854-list-files-in-google-drive-folder
  Bluexm: added recursion on subfolders - SO: https://webapps.stackexchange.com/a/142584

  Imran <team@yieldmore.org> in Dec 2024 has
   * introduced "onlyFolders" for own use
   * done a full cleanup of all variables
   * uses ' » ' as folder separator
   * added removeEmptyRows/Columns from Trey (SO)
   * added .getParents()[0].getName() per SO: https://stackoverflow.com/a/17618407
   * Merged Files & Fols into one sheet and added level, indent

  Get this From:
    https://github.com/yieldmore/google-apps-scripts/blob/master/folder-indexer.gs

  TODO:
    * Put in a date sheet and move that sheet to the beginning and set it as active.
    * Debate whether to list folders / files first.
*/

var sheet = SpreadsheetApp.getActiveSheet()
var sheetFile = DriveApp.getFileById(sheet.getParent().getId())
Logger.log('Detected Sheet: ' + sheetFile.getName())

var sharedDriveId = sheetFile.getParents().next().getId()
var sharedDrive = Drive.Drives.get(sharedDriveId, { supportsAllDrives: true })
Logger.log('Scanning Shared Drive: ' + sharedDrive.getName())

var topFolderName = sharedDrive.name

function ScanFolder() {
  ScanFoldeRecursively(topFolderName, '', 0, '')
  removeEmptyColumns()
  removeEmptyRows()
}

function ScanFoldeRecursively(folderName, relativeFolderName, level, indent) {
  var isTopFolder = topFolderName == folderName
  if (isTopFolder) {
    sheet.clearContents()

    sheet.appendRow([
      "Indent",
      "Level",
      "Folder",
      "File",
      "Description",
      "Date Last Updated",
      "Size",
      "URL",
      //"ID",
      "Type",
      "Folder",
    ])

    Logger.log("REBUILDING SHEET In: " + sheetFile.getName())
  }

  Logger.log("SCANNING: " + folderName) //DEBUG

  var folder = isTopFolder
    ? DriveApp.getFolderById(sharedDrive.id)
    : DriveApp.getFoldersByName(folderName).next()

  var files = getFilesOf(isTopFolder, folder)

  var parentFolderPrefix = relativeFolderName == topFolderName ? '' : relativeFolderName + ' » '

  var fileIndex = 0
  while (fileIndex < files.length) {
    var item = files[fileIndex]
    fileIndex += 1

    var data = [
      level,
      "--",
      indent + item.getName(),
      item.getDescription(),
      isTopFolder ? item.ModifiedTimeRaw : item.getLastUpdated(),
      item.getSize(),
      isTopFolder ? item.WebViewLink : item.getUrl(),
      //item.getId(),
      'FILE', //item.getBlob().getContentType(),
      parentFolderPrefix + folderName,
    ]

    sheet.appendRow(data)
  }

  var subFolders = folder.getFolders()

  while (subFolders.hasNext()) {
    var item = subFolders.next()
    Logger.log('Adding Folder: ' + item.getName())

    var data = [
      level,
      indent + item.getName(),
      "--",
      item.getDescription(),
      isTopFolder ? item.ModifiedTimeRaw : item.getLastUpdated(),
      item.getSize(),
      item.getUrl(),
      //item.getId(),
      'Google Folder',
      relativeFolderName + ' » ' + folderName,
    ]

    sheet.appendRow(data)

    var relativeFolderParam = folderName + (topFolderName == folderName ? '' : ' « ' + relativeFolderName)
    ScanFoldeRecursively(item.getName(), relativeFolderParam, level + 1, indent + '  ')
  }
}

function getFilesOf(top, folder) {
  if (top) {
    return Drive.Files.list({driveId: sharedDriveId, corpora: "drive",
      includeItemsFromAllDrives: true, supportsAllDrives: true}).files
  }
  
  var result = []
  var files = folder.getFiles()

  while (files.hasNext())
    result.push(files.next())

  return result;
}

//NOT USED: https://arisazhar.com/remove-empty-rows-in-spreadsheet-instantly/

//FROM: https://stackoverflow.com/a/34781833
//Remove All Empty Columns in the Entire Workbook
function removeEmptyColumns() {
  var ss = SpreadsheetApp.getActive()
  var allsheets = ss.getSheets()
  for (var s in allsheets) {
    var sheet=allsheets[s]
    var maxColumns = sheet.getMaxColumns()
    var lastColumn = sheet.getLastColumn()
    if (maxColumns - lastColumn != 0) {
      sheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn)
    }
  }
}

//Remove All Empty Rows in the Entire Workbook
function removeEmptyRows() {
  var ss = SpreadsheetApp.getActive()
  var allsheets = ss.getSheets()

  for (var s in allsheets){
    var sheet = allsheets[s]
    var maxRows = sheet.getMaxRows()
    var lastRow = sheet.getLastRow()
    if (maxRows - lastRow != 0) {
      sheet.deleteRows(lastRow + 1, maxRows - lastRow)
    }
  }
}
