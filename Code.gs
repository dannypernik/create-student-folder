function NewSatFolder(nameOnReport=true) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var sourceFolder = file.getParents().next();
  var sourceFolderId = sourceFolder.getId();
  var parentFolderId = sourceFolder.getParents().next().getId();

  var ui = SpreadsheetApp.getUi();
  var studentName = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL).getResponseText();

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();
  
  if (nameOnReport) {
    nameOnReport = studentName;
  }

  copyFolder(sourceFolderId, newFolderId, studentName, 'sat');
  linkSheets(newFolderId, nameOnReport);

  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + studentName + '\'s folder</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'SAT folder created successfully');
}

function NewActFolder(nameOnReport=true) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var sourceFolder = file.getParents().next();
  var sourceFolderId = sourceFolder.getId();
  var parentFolderId = sourceFolder.getParents().next().getId();

  var ui = SpreadsheetApp.getUi();
  var studentName = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL).getResponseText();

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();

  Logger.log('nameOnReport: ' + nameOnReport);
  if (nameOnReport) {
    nameOnReport = studentName;
  }
  Logger.log('nameOnReport: ' + nameOnReport);


  copyFolder(sourceFolderId, newFolderId, studentName, 'act');
  linkSheets(newFolderId, nameOnReport);

  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + studentName + '\'s folder</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ACT folder created successfully');
}

function NewTestPrepFolder(nameOnReport=true) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var sourceFolder = file.getParents().next();
  var sourceFolderId = sourceFolder.getId();
  var parentFolderId = sourceFolder.getParents().next().getId();

  var ui = SpreadsheetApp.getUi();
  var studentName = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL).getResponseText();

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();
  Logger.log('nameOnReport: ' + nameOnReport);

  if (nameOnReport) {
    nameOnReport = studentName;
  }
  Logger.log('nameOnReport: ' + nameOnReport);

  copyFolder(sourceFolderId, newFolderId, studentName, 'all');
  linkSheets(newFolderId, nameOnReport);

  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + studentName + '\'s folder</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Test prep folder created successfully');
}

function copyFolder(sourceFolderId = '1yqQx_qLsgqoNiDoKR9b63mLLeOiCoTwo', newFolderId = '1_qQNYnGPFAePo8UE5NfX72irNtZGF5kF', studentName = '_Aaron S', folderType = 'sat') {

  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const newFolder = DriveApp.getFolderById(newFolderId);

  var sourceSubFolders = sourceFolder.getFolders();
  var files = sourceFolder.getFiles();

  if (folderType.toLowerCase() === 'sat') {
    var testType = 'SAT';
  }
  else if (folderType.toLowerCase() === 'act') {
    var testType = 'ACT';
  }
  else {
    var testType = 'Test';
  }

  while (files.hasNext()) {
    var file = files.next();
    let prefixFiles = ['Tutoring notes', 'ACT review sheet', 'SAT review sheet'];
    var fileName = file.getName();
    Logger.log(fileName);

    if (prefixFiles.includes(fileName)) {
      fileName = studentName + " " + fileName;
    }
    else if (fileName.toLowerCase().includes('template')) {
      rootName = fileName.slice(0, fileName.indexOf('-') + 2);
      fileName = rootName + studentName;
    }

    var newFile = file.makeCopy(fileName, newFolder);
    var newFileName = newFile.getName().toLowerCase();

    if (newFileName.includes('tutoring notes')) {
      var ssId = newFile.getId();
      var ss = SpreadsheetApp.openById(ssId);
      var sheet = ss.getSheetByName('Session notes');
      shId = sheet.getSheetId();
      sheet.getRange('G3').setValue('=hyperlink("https://docs.google.com/spreadsheets/d/' + ssId + '/edit?gid=' + shId + '#gid=' + shId + '&range=B"&match(G2,B1:B,0)-1,"Go to latest session")');
    }

    if (newFileName.includes('admin notes')) {
      DocumentApp.openById(newFile.getId()).getBody().replaceText('StudentName', studentName);
    }

    if (testType === 'SAT' && fileName.toLowerCase().includes('act') && fileName.toLowerCase().includes('answer analysis')) {
      newFile.setTrashed(true);
    }
    else if (testType === 'ACT' && fileName.toLowerCase().includes('sat') && fileName.toLowerCase().includes('answer analysis')) {
      newFile.setTrashed(true);
    }

    if (newFolder.getName().includes(folderType.toUpperCase()) && !newFolder.getName().includes(studentName)) {
      newFile.moveTo(newFolder.getParents().next());
      Logger.log("new location: " + newFile.getParents().next().getId());
      if (isEmptyFolder(newFolder.getId())) {
        newFolder.setTrashed(true);
        Logger.log(newFolder.getName() + " trashed")
      }
    }
  }

  while (sourceSubFolders.hasNext()) {
    var sourceSubFolder = sourceSubFolders.next();
    var folderName = sourceSubFolder.getName();
    Logger.log(folderName + ' ' + newFolder);

    if (folderName === 'Student') {
      var targetFolder = newFolder.createFolder(studentName + " " + testType + " prep");
    }
    else if (newFolder.getName().includes(folderType.toUpperCase()) && newFolder.getName() !== studentName + " " + testType + " prep") {
      var targetFolder = newFolder.getParents().next().createFolder(folderName);
      Logger.log(sourceSubFolder.getId() + " moved");
    }
    else {
      var targetFolder = newFolder.createFolder(folderName);
    }

    if (targetFolder.getName().includes('ACT') && folderType.toLowerCase() === 'sat') {
      targetFolder.setTrashed(true);
      Logger.log(targetFolder.getName() + " trashed");
    }
    else if (targetFolder.getName().includes('SAT') && folderType.toLowerCase() === 'act') {
      targetFolder.setTrashed(true);
      Logger.log(targetFolder.getName() + " trashed");
    }
    else {
      copyFolder(sourceSubFolder.getId(), targetFolder.getId(), studentName, folderType);
    }
  }

  //SpreadsheetApp.getUi().alert("copyFolder() ended");
}

var satSheetIds = {
  'admin': null,
  'student': null,
  'studentData': null,
  'adminData': null
}

var satSheetDataUrls = {
  'admin': null,
  'student': null
}

var actSheetIds = {
  'admin': null,
  'student': null,
  'studentData': null,
  'adminData': null
}

var actSheetDataUrls = {
  'admin': null,
  'student': null
}

function linkSheets(folderId, nameOnReport=false) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var subFolders = DriveApp.getFolderById(folderId).getFolders();

  while (files.hasNext()) {
    file = files.next();
    fileName = file.getName();
    if (fileName.includes('SAT')) {
      if (fileName.toLowerCase().includes('student answer sheet')) {
        satSheetIds.student = file.getId();
        DriveApp.getFileById(satSheetIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      }
      else if (fileName.toLowerCase().includes('answer analysis')) {
        satSheetIds.admin = file.getId();

        var ss = SpreadsheetApp.openById(file.getId());
        if (nameOnReport) {
          for (i in ss.getSheets()) {
            var s = ss.getSheets()[i];
            if (s.getName().toLowerCase().includes('analysis') || s.getName().toLowerCase().includes('opportunity')) {
              s.getRange('D5').setValue('for ' + nameOnReport)
            }
          }
        }
      }
    }

    if (fileName.includes('ACT')) {
      if (fileName.toLowerCase().includes('student answer sheet')) {
        actSheetIds.student = file.getId();
        DriveApp.getFileById(actSheetIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      }
      else if (fileName.toLowerCase().includes('answer analysis')) {
        actSheetIds.admin = file.getId();
      }
    }
  }

  if (satSheetIds.student && satSheetIds.admin) {
    Logger.log(satSheetIds.student);
    Logger.log(satSheetIds.admin);
    SpreadsheetApp.openById(satSheetIds.admin).getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);
    SpreadsheetApp.openById(satSheetIds.student).getSheetByName('Question bank data').getRange('I2').setValue('=iferror(importrange("' + satSheetIds.admin + '","Question bank data!I2:I"),"")');
    SpreadsheetApp.openById(satSheetIds.student).getSheets()[0].getRange('D1').setValue('=importrange("' + satSheetIds.admin + '","Question bank data!V1")');
  }
  Logger.log('actSheetIds.student: ' + actSheetIds.student);
  Logger.log('actSheetIds.admin: ' + actSheetIds.admin);
  if (actSheetIds.student && actSheetIds.admin) {
    SpreadsheetApp.openById(actSheetIds.admin).getSheetByName('Student responses').getRange('B1').setValue(actSheetIds.student);
  }

  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    linkSheets(subFolder.getId(), nameOnReport);
  }
}

function isEmptyFolder(folderId) {
  const folders = DriveApp.getFolderById(folderId).getFolders();
  const files = DriveApp.getFolderById(folderId).getFiles();

  if (folders.hasNext() || files.hasNext()) {
    return false;
  }
  else {
    return true;
  }
}

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts').addItem('New SAT student', 'NewSatFolder').addItem('New ACT student', 'NewActFolder').addItem('New Test prep student', 'NewTestPrepFolder').addToUi();
}