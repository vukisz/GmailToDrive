function deleteDuplicatesInDriveFolder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSettings = ss.getSheetByName(SETTINGS.sheetNameSettings);
  var settingsJson = readSettingsToJson(sheetSettings)
  var sheet = ss.getSheetByName(settingsJson.SheetTransName);
  //var sheet = ss.getSheetByName('Trans2');
  var firstRowRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var firstRowValues = firstRowRange.getValues();
  var dupsCleanedColIndex = firstRowValues[0].indexOf("DupsCleaned") + 1;
  var transInJson = readTransToJson(sheet, dupsCleanedColIndex)


  for (var key in transInJson)  // OK in V8
  {
    var lines = transInJson[key];

    Logger.log('Processing folder: %s (ID: %s)', DriveApp.getFolderById(lines['FolderId']).getName(), lines['FolderId']);
    removeDupsInFolder(lines['FolderId']);

    sheet.getRange(lines.rowIndex, dupsCleanedColIndex).setValue('TRUE')

  }

  console.info('done')
}

function removeDupsInFolder(_folderId) {
  //var query = "'1q3godCvVoyqQCur3mSZmd82Bo1yRV8hS' in parents";
  var query = "'" + _folderId + "' in parents and trashed = false";

  //var query = "title contains '!MessageBody.pdf' and trashed = false"
  var totalDups = 0
  var files;

  var pageToken;



  //This part need Drive API
  do {
    files = Drive.Files.list({
      q: query,
      maxResults: 100,
      pageToken: pageToken,
      orderBy: "title asc",
    });
    var jsonMap = Object();

    if (files.items && files.items.length > 0) {
      for (var i = 0; i < files.items.length; i++) {
        var file = files.items[i]
        var name = file.title
        var md5Checksum = file.md5Checksum

        //if (name.endsWith("!MessageBody.html")) continue;

        //if (name.toLowerCase().endsWith(".jpg") || name.toLowerCase().endsWith(".jpeg") || name.toLowerCase().endsWith(".png") || name.toLowerCase().endsWith(".gif") || name.toLowerCase().endsWith(".pdf")) {
        if (jsonMap[md5Checksum] != null) {
          totalDups++;
          console.info('Dup found', name, jsonMap[md5Checksum])
          Drive.Files.trash(file.id)
        } else {
          jsonMap[md5Checksum] = name;
        }
        //}

      }
    } else {
      Logger.log('No files found.');
    }
    pageToken = files.nextPageToken;
  } while (pageToken);
  console.info('Total: ' + totalDups);
}
