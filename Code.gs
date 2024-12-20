function NewSatFolder(nameOnReport=true) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var sourceFolder = file.getParents().next();
  var sourceFolderId = sourceFolder.getId();
  var parentFolderId = sourceFolder.getParents().next().getId();

  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);
  if(prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  else {
    var studentName = prompt.getResponseText();
  }

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
  var prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);

  if(prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  else {
    var studentName = prompt.getResponseText();
  }

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();

  if (nameOnReport) {
    nameOnReport = studentName;
  }

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
  var prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);

  if(prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  else {
    var studentName = prompt.getResponseText();
  }

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();

  if (nameOnReport) {
    nameOnReport = studentName;
  }

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
  const newFolderName = newFolder.getName();

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

    if (newFolderName.includes(folderType.toUpperCase()) && !newFolderName.includes(studentName)) {
      newFile.moveTo(newFolder.getParents().next());
      Logger.log("new location for " + newFileName + ": " + newFile.getParents().next().getId());
      if (isEmptyFolder(newFolder.getId())) {
        newFolder.setTrashed(true);
        Logger.log(newFolderName + " trashed")
      }
    }
  }

  while (sourceSubFolders.hasNext()) {
    var sourceSubFolder = sourceSubFolders.next();
    var folderName = sourceSubFolder.getName();

    if (folderName === 'Student') {
      var targetFolder = newFolder.createFolder(studentName + " " + testType + " prep");
    }
    else if (newFolderName.includes(folderType.toUpperCase()) && newFolderName !== studentName + " " + testType + " prep") {
      var targetFolder = newFolder.getParents().next().createFolder(folderName);
      Logger.log(sourceSubFolder.getName() + " moved");
    }
    else {
      var targetFolder = newFolder.createFolder(folderName);
    }

    targetFolderName = targetFolder.getName();

    if (targetFolderName.includes('ACT') && folderType.toLowerCase() === 'sat') {
      targetFolder.setTrashed(true);
      Logger.log(targetFolderName + " trashed");
    }
    else if (targetFolderName.includes('SAT') && folderType.toLowerCase() === 'act') {
      targetFolder.setTrashed(true);
      Logger.log(targetFolderName + " trashed");
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
        for (i in ss.getSheets()) {
          var s = ss.getSheets()[i];

          if (s.getName().toLowerCase().includes('analysis') || s.getName().toLowerCase().includes('opportunity')) {
            if (nameOnReport) {
              s.getRange('D4').setValue('for ' + nameOnReport)
            }
          }
          else {
            var protections = s.getProtections(SpreadsheetApp.ProtectionType.SHEET);
            for(var p=0; p< protections.length; p++) {
              protections[p].setUnprotectedRanges([s.getRange('C5:C'),s.getRange('G5:G'),s.getRange('K5:K')]);
            }
          }
        }
        let revBackend = ss.getSheetByName('Rev sheet backend');
        revBackend.getRange('K2').setValue(nameOnReport);
        revBackend.protect().setWarningOnly(true);
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
    let satAdminSheet = SpreadsheetApp.openById(satSheetIds.admin);
    let satStudentSheet = SpreadsheetApp.openById(satSheetIds.student);
    satAdminSheet.getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);
    
    let revDataId = satAdminSheet.getSheetByName('Rev sheet backend').getValue('D2');

    let adminRevSheet = satAdminSheet.getSheetByName('Rev sheets');
    adminRevSheet.getRange('B5').setValue('=importrange("' + revDataId + '", "' + nameOnReport + '!B5:C")');
    adminRevSheet.getRange('G5').setValue('=importrange("' + revDataId + '", "' + nameOnReport + '!E5:F")');

    let studentRevSheet = satStudentSheet.getSheetByName('Rev sheets');
    studentRevSheet.getRange('B5').setValue('=importrange("' + revDataId + '", "' + nameOnReport + '!B5:C")');
    studentRevSheet.getRange('F5').setValue('=importrange("' + revDataId + '", "' + nameOnReport + '!E5:F")');

    SpreadsheetApp.openById(satSheetIds.admin).getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);
    
    // SpreadsheetApp.openById(satSheetIds.student).getSheetByName('Question bank data').getRange('U2').setValue(satSheetIds.admin);
    // SpreadsheetApp.openById(satSheetIds.student).getSheetByName('Question bank data').getRange('I2').setValue('=iferror(importrange("' + satSheetIds.admin + '","Question bank data!I2:I"),"")');
    // SpreadsheetApp.openById(satSheetIds.student).getSheets()[0].getRange('D1').setValue('=importrange("' + satSheetIds.admin + '","Question bank data!V1")');
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

function createRwRevSheet() {
  createRevSheet('RW', 0);
}

function createMathRevSheet() {
  createRevSheet('Math', 1)
}

function createRevSheet(sub, subIndex) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let revBackend = ss.getSheetByName('Rev sheet backend');
  let revSheet = ss.getSheetByName(sub + ' Rev sheet');
  let revResponseSheet = ss.getSheetByName('Rev sheets')
  let subBackendOffset = subIndex * 4;
  let folderIdRange = revBackend.getRange(2, 3 + subBackendOffset);
  let revSheetSubjectFolderId = folderIdRange.getValue();
  let satFolder = null;
  let studentName = revBackend.getRange('K2');
  let revDataSs = SpreadsheetApp.openById(revBackend.getRange('D2'));
  let revData = revDataSs.getSheetByName(studentName);

  if(!revData) {
    revDataSs.getSheetByName('Template').copyTo(revDataSs).setName(studentName);
  }

  if (!revBackend.getRange(2, 1 + subBackendOffset).getValue()) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Error: No missed questions available for ' + revResponseSheet.getRange(1, 3 + subIndex * 5).getValue());
    return;
  }
  
  var maxQuestionRange = revBackend.getRange('L2');
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Max # of questions - leave blank to use prior value of ' + maxQuestionRange.getValue(), ui.ButtonSet.OK_CANCEL);
  if(prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  else if (prompt.getResponseText() !== '') {
    maxQuestionRange.setValue(prompt.getResponseText());
  }

  try {
    DriveApp.getFolderById(revSheetSubjectFolderId);
  }
  catch {
    revSheetSubjectFolderId = ''
    folderIdRange.setValue(revSheetSubjectFolderId);
    Logger.log('blank/invalid folder ID in ' + folderIdRange);
  }
  
  if (revSheetSubjectFolderId === '') {
    var adminFolder = DriveApp.getFileById(ss.getId()).getParents().next();
    var subfolders = adminFolder.getFolders();

    if(sub === 'RW') {
      var subject = 'Reading & Writing';
    }
    else {
      var subject = 'Math';
    }

    if (subfolders.hasNext()) {
      var revSheetParentFolder = subfolders.next();
      var nextSubfolders = revSheetParentFolder.getFolders();
      var revSheetFolder = null;

      while (nextSubfolders.hasNext()) {
        var nextSubfolder = nextSubfolders.next();
        if (nextSubfolder.getName().toLowerCase().includes('rev')) {
          revSheetFolder = nextSubfolder;
        }
        else if (nextSubfolder.getName().includes('SAT')) {
          satFolder = subfolder;
        }
      }

      if (revSheetFolder) {
        getRevSubjectFolderId(revSheetFolder);
      }
      else if (satFolder) {
        let subfolders = satFolder.getFolders();
        while (subfolders.hasNext()) {
          let subfolder = subfolders.next();

          if (subfolder.getName().toLowerCase().includes('rev')) {
            revSheetFolder = subfolder;
            getRevSubjectFolder(revSheetFolder);
          }
        }
      }
      else {
        revSheetSubjectFolderId = revSheetParentFolder.createFolder('Rev sheets').createFolder(subject).getId();
      }
    }
    else {
      revSheetSubjectFolderId = adminFolder.createFolder('Rev sheets').createFolder(subject).getId();
    }

    folderIdRange.setValue(revSheetSubjectFolderId);
  }


  revSheet.showSheet();
  revSheet.showRows(1,revSheet.getMaxRows());
  revBackend.getRange(2, 2 + subBackendOffset, revBackend.getLastRow() - 1).clear();
  revBackend.getRange(2, 2 + subBackendOffset).setValue('=RANDARRAY(counta(A$2:A))');
  SpreadsheetApp.flush();
  revBackend.getRange(2, 2 + subBackendOffset, revBackend.getMaxRows() - 1).copyValuesToRange(revBackend.getSheetId(), 2+subBackendOffset, 2+subBackendOffset, 2, 2);

  var idCol = revSheet.getRange('B1:B');
  var values = idCol.getValues(); // get all data in one call
  var heights = revSheet.getRange('E1:E');
  var heightVals = heights.getValues();
  //var imgContainerWidth = revSheet.getColumnWidth(4);
  var row = 6;

  try {
    while ( values[row-1] && values[row-1][0] != "" ) {
      var questionId = values[row-1][0];  
      var rowHeight = heightVals[row-1][0]; // rowHeights hard-coded in Rev sheet backend
      revSheet.setRowHeight(row, rowHeight);
      Logger.log(questionId + ' rowHeight: ' + rowHeight);
      row++;
    }
  }
  catch(err) {
    if (err.message.includes('Invalid argument')){
      SpreadsheetApp.getUi().alert('Error: Image not found');
    }
    else {
      SpreadsheetApp.getUi().alert(err);
    }
    return;
  }
  
  var firstEmptyRow = getFirstEmptyRow(revData, 2 + subIndex * 3);
  if (firstEmptyRow === 5) {
    var newRevSheetNumber = 1;
  }
  else {
    var revSheetLastQuestion = revData.getRange(firstEmptyRow - 1, 2 + subIndex * 3).getValue().toString();
    Logger.log('revSheetLastQuestion' + revSheetLastQuestion);
    var newRevSheetNumber = parseInt(revSheetLastQuestion.substring(revSheetLastQuestion.lastIndexOf(' ') + 1, revSheetLastQuestion.indexOf('.'))) + 1;
  }
  revSheet.getRange('E1').setValue(newRevSheetNumber);

  // hide unneeded rows, column A+G
  revSheet.hideRows(row, revSheet.getMaxRows() - row + 1);
  revSheet.hideColumns(3);
  revSheet.hideColumns(6);
  revSheet.showColumns(5);

  if (!studentName) {
    var pdfName = sub + ' Rev sheet #' + newRevSheetNumber;
  }
  else {
    var pdfName = sub + ' Rev sheet #' + newRevSheetNumber + ' for ' + studentName;
  }

  //* Create worksheets
  SpreadsheetApp.flush();
  savePdf(ss, revSheet, pdfName, revSheetSubjectFolderId);
  Logger.log(sub + ' Rev sheet #' + newRevSheetNumber + ' saved');
  //*/

  revSheet.showColumns(3);
  revSheet.showColumns(6);
  revSheet.hideColumns(5);

  //* Create answer keys
  SpreadsheetApp.flush();
  savePdf(ss, revSheet, pdfName + '~Key', revSheetSubjectFolderId);
  Logger.log(sub + ' Rev key #' + newRevSheetNumber + ' saved')
  //*/

  var dataToCopy = revSheet.getRange(6,1,row-5,2).getValues();
  revData.getRange(firstEmptyRow, 2 + subIndex * 3, row-5, 2).setValues(dataToCopy);

  revSheet.showRows(1,revSheet.getMaxRows());
  revSheet.hideSheet();

  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + revSheetSubjectFolderId + '" target="_blank" onclick="google.script.host.close()">' + sub + ' Rev sheet folder</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Rev sheet complete');
}


function getRevSubjectFolderId(revSheetFolder) {
  let revSheetSubjectFolderId;

  while (revSheetFolder.hasNext()) {
    let subfolder = revSheetFolder.next();
    let subfolderName = subfolder.getName();
    if (subfolderName.toLowerCase().includes(subject)) {
      revSheetSubjectFolderId = subfolder.getId();
      break;
    }
  }
  if (!revSheetSubjectFolderId) {
    revSheetSubjectFolderId = revSheetFolder.createFolder(subject).getId();
  }

  return revSheetSubjectFolderId;
}


function savePdf(spreadsheet, sheet, pdfName, pdfFolderId) {
  var sheetId = sheet.getSheetId();
  var url_base = spreadsheet.getUrl().replace(/edit$/,'');

  var url_ext = 'export?exportFormat=pdf&format=pdf'
  + '&gid=' + sheetId
  // following parameters are optional...
  + '&size=A4'      // paper size: legal / letter / A4
  + '&portrait=true'    // orientation, false for landscape
  + '&fitw=true'        // fit to width, false for actual size
  + '&top_margin=0.25'
  + '&bottom_margin=0'
  + '&left_margin=0.375'
  + '&right_margin=0.375'
  + '&sheetnames=false' // 
  + '&printtitle=false'
  + '&pagenumbers=false'  //hide optional headers and footers
  + '&gridlines=false'  // hide gridlines
  + '&fzr=true';       // false = do not repeat row headers (frozen rows) on each page
  var url_options = {headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),},muteHttpExceptions: true};
  var response = (function backoff(i) {
    Utilities.sleep(Math.pow(2, i) * 1000);
    let data = UrlFetchApp.fetch(url_base + url_ext, url_options);
    if (data.getResponseCode() !== 200) {
      return backoff(++i);
    }
    else {
      return data;
    }
  })(1);
  var blob = response.getBlob().getAs('application/pdf').setName(pdfName + '.pdf');
  var folder = DriveApp.getFolderById(pdfFolderId);
  folder.createFile(blob);
}


// Adapted from https://stackoverflow.com/a/9102463/1677912
function getFirstEmptyRow(sheet, colIndex) {
  var column = sheet.getRange(5, colIndex, sheet.getLastRow() - 4);
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  return (ct+5);  // +5 since starting from row 5 with 0-indexing
}


function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts')
    .addItem('New SAT student', 'NewSatFolder')
    .addItem('New ACT student', 'NewActFolder')
    .addItem('New Test prep student', 'NewTestPrepFolder')
    .addToUi();
}