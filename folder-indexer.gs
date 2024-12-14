/*
  What this does:
    Lists all files and sub-folders from a folder in Google Drive.
    Looks for a "Drives to Scan" sheet to scan in multiple places, else scans activeSheet >> parent.

  View Code / History:
    https://github.com/yieldmore/google-apps-scripts/blob/master/folder-indexer.gs

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
   * Decided whether to list subfolders first, not files.
   * Support for Drive Folder / Shared Drive (14th Dec)
   * Uses "Drives to Scan" which tells the name/type, and, if missing, will search SharedDrives to detect the type

  TODO:
   * Set status of row and update the count
   * Move output to a teams.amadeusweb.com and use datatables / an auth database to show it.
*/

var sheet = SpreadsheetApp.getActiveSheet(),
    sheetFile = SpreadsheetApp.getActiveSpreadsheet(),
    namesSheet = sheetFile.getSheetByName('Drives to Scan')

var isSharedDrive = false,
    isTopFolder = false,
    topFolderName = '',
    topFolderId = ''

function ScanAllDrives() {
  if (sheet == null) {
    Logger.log('No Active Sheet')
    return
  }

  Logger.log('Detected Sheet: ' + sheetFile.getName())


  var parents = []

  if (namesSheet != null) {
    rows = namesSheet.getRange(2, 1, namesSheet.getLastRow() - 1, 2).getValues()
    rows.forEach(function(item) { parents.push({name: item[0], type: item[1]}) })
    sheetFile.setActiveSheet(namesSheet)
  } else {
    var parentType = 'folder',
      parentName = DriveApp.getFileById(sheetFile.getId()).getParents().next().getName()

    try {
      var sharedDrive = Drive.Drives.list
      ({q: 'name = "' + parentName + '"', supportsAllDrives: true })
      .drives.pop()
      topFolderId = sharedDrive.id
      parentType = 'shared'
    } catch {}
    parents.push({name: parentName, type: parentType})
  }

  parents.forEach(ScanDrive)
}

function ScanDrive(parent, nameIndex) {
  Logger.log(parent)
  setCount('...', nameIndex)

  isSharedDrive = parent.type == 'shared'
  topFolderName = parent.name

  Logger.log('Scanning ' + (isSharedDrive ? 'Shared' : 'Folder in ') + ' Drive: ' + topFolderName)

  sheet = sheetFile.getSheetByName(parent.name)
  if (sheet == null) {
    sheet = sheetFile.insertSheet(parent.name)
  }

  ScanFolder(topFolderName, '', 0, '')
  removeEmptyColumns()
  removeEmptyRows()
  setCount(sheet.getLastRow() - 1, nameIndex)
}

function ScanFolder(folderName, relativeFolderName, level, indent) {
  isTopFolder = topFolderName == folderName
  if (isTopFolder) {
    sheet.clearContents()

    sheet.appendRow([
      "Level",
      "Type",
      "Name",
      "Description",
      "Date Last Updated",
      "Size",
      "URL",
      "Folder",
    ])

    Logger.log("REBUILDING SHEET In: " + sheetFile.getName())
  }

  Logger.log("SCANNING: " + folderName)

  var relativeFolder = folderName + (isTopFolder ? '' : ' « ' + relativeFolderName)

  var folder = DriveApp.getFoldersByName(folderName).next()

  var subFolders = getFoldersOf(folder)

  var folderIndex = 0
  while (folderIndex < subFolders.length) {
    var item = subFolders[folderIndex]
    folderIndex += 1

    var data = [
      level,
      'Folder',
      indent + '/ ' + item.getName(),
      item.getDescription(),
      isTopFolder && isSharedDrive ? item.ModifiedTimeRaw : item.getLastUpdated(),
      item.getSize(),
      item.getUrl(),
      relativeFolder,
    ]

    sheet.appendRow(data)

    ScanFolder(item.getName(), folderName + ' « ' + relativeFolder, level + 1, indent + '  ')
  }

  var files = getFilesOf(folder)

  var fileIndex = 0
  while (fileIndex < files.length) {
    var item = files[fileIndex]
    fileIndex += 1

    var data = [
      level,
      'File',
      indent + '/ ' + item.getName(),
      item.getDescription(),
      isTopFolder && isSharedDrive ? item.ModifiedTimeRaw : item.getLastUpdated(),
      item.getSize(),
      isTopFolder && isSharedDrive ? item.WebViewLink : item.getUrl(),
      relativeFolder,
    ]

    sheet.appendRow(data)
  }

}

function sortAscending(item1, item2) {
  var a = item1.getName(), b = item2.getName()
  return a > b ? 1 : (a < b ? -1 : 0)
}

function getFoldersOf(folder) {
  var result = []

  var subFolders = folder.getFolders();
  while (subFolders.hasNext())
    result.push(subFolders.next())
  
  result.sort(sortAscending)
  return result
}

function getFilesOf(folder) {
  if (isTopFolder && isSharedDrive) {
    return Drive.Files.list({driveId: topFolderId, corpora: "drive",
      includeItemsFromAllDrives: true, supportsAllDrives: true}).files
  }
  
  var result = []
  var files = folder.getFiles()

  while (files.hasNext())
    result.push(files.next())

  result.sort(sortAscending)
  return result;
}

function setCount(to, row) {
  namesSheet.getRange(row + 2, 4).setValue(to)
}

//NOT USED: https://arisazhar.com/remove-empty-rows-in-spreadsheet-instantly/

//FROM: https://stackoverflow.com/a/34781833
//UPDATE: Use Global Sheet Variable
//Remove All Empty Columns in the Current Sheet
function removeEmptyColumns() {
  var maxColumns = sheet.getMaxColumns()
  var lastColumn = sheet.getLastColumn()
  if (maxColumns - lastColumn != 0) {
    sheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn)
  }
}

//Remove All Empty Rows in the Current Sheet
function removeEmptyRows() {
  var maxRows = sheet.getMaxRows()
  var lastRow = sheet.getLastRow()
  if (maxRows - lastRow != 0) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow)
  }
}
