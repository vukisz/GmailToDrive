SETTINGS = {
  sheetNameSettings: "SettingsAndRunningValues"
};


function run() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSettings = ss.getSheetByName(SETTINGS.sheetNameSettings);
  var settingsJson = readSettingsToJson(sheetSettings)
  var sheet = ss.getSheetByName(settingsJson.SheetTransName);

  var dateOfScriptRun = Utilities.formatDate(new Date(), 'Europe/Vilnius', 'yyyy-MM-dd HH:mm:ss')

  var query = '';
  var now = new Date();


  if (settingsJson.donotExecute) {
    alertWithTryCatch('donotExecute is set to true')
  }


  if (settingsJson.labelNameToDoSearchIn != '')
    query = 'label:' + settingsJson.labelNameToDoSearchIn + ' ';//Important, as it is removing this label when thread is being processed

  var threads = GmailApp.search(query);

  var file = DriveApp.getFileById(ss.getId());
  var folders = file.getParents();
  while (folders.hasNext()) {
    destFolder = DriveApp.getFolderById(folders.next().getId())
    //Logger.log('folder name = '+folders.next().getName());
  }

  if (GmailApp.getUserLabelByName(settingsJson.labelNameToDoSearchIn) == null) {
    alertWithTryCatch('labelNameToDoSearchIn Label ' + settingsJson.labelNameToDoSearchIn + ' does not exists. Exiting')
    return;
  }
  if (GmailApp.getUserLabelByName(settingsJson.assignLabelForArchived) == null) {
    alertWithTryCatch('assignLabelForArchived Label ' + settingsJson.assignLabelForArchived + ' does not exists. Exiting')
    return;
  }
  if (threads.length == 0) {

    alertWithTryCatch('All messages have been processed.')
    return;
  }



  for (var i in threads) {
    var mesgs = threads[i].getMessages();
    for (var j in mesgs) {

      var size = mesgs[j].getRawContent().length

      if (size > settingsJson.emailSizeInBytesToArchive) {
        var dateOfFirstMessage = Utilities.formatDate(mesgs[0].getDate(), settingsJson.TimeZone, settingsJson.DateTimeFormat)
        var dateOfCurMessage = Utilities.formatDate(mesgs[j].getDate(), settingsJson.TimeZone, settingsJson.DateTimeFormat);
        var subfolderName = dateOfFirstMessage + ' ' + mesgs[0].getSubject()
        var folderFinal
        var messageBodyFile
        var attachmentsCount = 0

        if (destFolder.getFoldersByName(subfolderName).hasNext()) {
          folderFinal = destFolder.getFoldersByName(subfolderName).next()
        }
        else if (settingsJson.saveAsSeparateEmlFiles || settingsJson.saveAsSeparateAttachments) {
          folderFinal = destFolder.createFolder(subfolderName)
        }

        if (settingsJson.saveAsSeparateEmlFiles) {// case if saving attachments as whole .eml messages (would not help finding attachment dups ;-) )        
          //Quite "heavy"
          //var size = Utilities.newBlob(mesgs[j].getRawContent()).getBytes().length;            
          var messageBlob = Utilities.newBlob(mesgs[j].getRawContent());
          var file = DriveApp.createFile(messageBlob.copyBlob().setName(dateOfCurMessage + ' ' + mesgs[0].getSubject() + '.eml'));

          if (destFolder.getFoldersByName(subfolderName).hasNext())
            destFolder.getFoldersByName(subfolderName).next().addFile(file)
          else
            destFolder.createFolder(subfolderName).addFile(file)
        }

        if (settingsJson.saveAsSeparateAttachments) { // save attachments as separate files        
          //Inspiration from https://stackoverflow.com/questions/60551975/copying-gmail-attachments-to-google-drive-using-apps-script/60552276#60552276
          var attA = mesgs[j].getAttachments();



          messageBodyFile = folderFinal.createFile(dateOfCurMessage + '_!MessageBody.pdf', mesgs[j].getBody(), MimeType.PDF)
          attA.forEach(function (a) {
            attachmentsCount++
            folderFinal.createFile(a.copyBlob()).setName(dateOfCurMessage + '_' + a.getName());
          });

        }

        if (settingsJson.logToSpreadsheet) { //log to spreadsheet
          sheet.appendRow([
            dateOfScriptRun,
            dateOfCurMessage,
            mesgs[j].getFrom(),
            mesgs[j].getTo(),
            mesgs[j].getCc(),
            mesgs[j].getSubject(),
            size,
            attachmentsCount,
            '=hyperlink("' + threads[i].getPermalink() + '", "View")',
            folderFinal != null ? '=hyperlink("https://drive.google.com/drive/u/0/folders/' + folderFinal.getId() + '", "View Folder")' : '-',
            //'=hyperlink("https://drive.google.com/uc?export=view&id=' + messageBodyFile.getId() + '", "View message body PDF")'
            messageBodyFile != null ? '=hyperlink("https://drive.google.com/file/d/' + messageBodyFile.getId() + '", "View message body PDF")' : '-'
          ]);
        }


        console.info(mesgs[j].getSubject().toString() + " length:" + size);

        if (settingsJson.duplicateMessageWithoutAttachments) {
          duplicateMessageWithoutAttachments(threads[i].getId(), mesgs[j].getId())
        }

        if (settingsJson.moveMessageToTrash) {
          mesgs[j].moveToTrash();
        }

        threads[i].addLabel(GmailApp.getUserLabelByName(settingsJson.assignLabelForArchived))
        threads[i].removeLabel(GmailApp.getUserLabelByName(settingsJson.labelNameToDoSearchIn))// Important! as if not removed script resume would start again with same mesages

      }
    }
  }

}


function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menu = [
    { name: "Run GmailToDrive", functionName: "run" }
  ];
  ss.addMenu("➫ GmailToDrive", menu);
  ss.toast("Click the Gmail Menu to continue..");
}

function duplicateMessageWithoutAttachments(_threadId, _emailId) {
  // Taken from https://stackoverflow.com/questions/46434390/remove-an-attachment-of-a-gmail-email-with-google-apps-script
  // Get the `raw` email
  var email = GmailApp.getMessageById(_emailId).getRawContent();

  // Find the end boundary of html or plain-text email
  var re_html = /(-*\w*)(\r)*(\n)*(?=Content-Type: text\/html;)/.exec(email);
  var re = re_html || /(-*\w*)(\r)*(\n)*(?=Content-Type: text\/plain;)/.exec(email);

  if (re == null) {
    var re_html = /(-*\w*)(\r)*(\n)*(?=Content-type: text\/html;)/.exec(email);
    var re = re_html || /(-*\w*)(\r)*(\n)*(?=Content-type: text\/plain;)/.exec(email);
  }

  // Find the index of the end of message boundary
  var start = re[1].length + re.index;
  var boundary = email.indexOf(re[1], start);

  // Remove the attachments & Encode the attachment-free RFC 2822 formatted email string
  var base64_encoded_email = Utilities.base64EncodeWebSafe(email.substr(0, boundary));
  // Set the base64Encoded string to the `raw` required property
  var resource = { 'raw': base64_encoded_email, "threadId": _threadId }

  // Re-insert the email into the user gmail account with the insert time
  /* var response = Gmail.Users.Messages.insert(resource, 'me'); */

  // Re-insert the email with the original date/time 
  var response = Gmail.Users.Messages.insert(resource, 'me',
    null, { 'internalDateSource': 'dateHeader' });

  Logger.log("The inserted email id is: %s", response.id)


}


function alertWithTryCatch(_message) {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert(_message)
  } catch (e) {
    // Logs an ERROR message.
    console.warn('SpreadsheetApp.getUi() yielded an error, but was handled: ' + e);
    console.info(_message)
  }
}
