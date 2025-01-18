function NewSatFolder(sourceFolderId, parentFolderId) {
  let ss, file, sourceFolder, studentName;
  if (sourceFolderId === undefined || parentFolderId === undefined) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    file = DriveApp.getFileById(ss.getId());
    sourceFolder = file.getParents().next();
    sourceFolderId = sourceFolder.getId();
    parentFolderId = sourceFolder.getParents().next().getId();
  }

  let ui = SpreadsheetApp.getUi();
  let prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);

  if (prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  } else {
    studentName = prompt.getResponseText();
  }

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();

  copyFolder(sourceFolderId, newFolderId, studentName, 'sat');
  linkSheets(newFolderId, studentName, 'sat');

  var htmlOutput = HtmlService.createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + studentName + "'s folder</a>")
    .setWidth(250)
    .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'SAT folder created successfully');
}

function NewActFolder(sourceFolderId, parentFolderId) {
  let ss, file, sourceFolder, studentName;
  if (sourceFolderId === undefined || parentFolderId === undefined) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    file = DriveApp.getFileById(ss.getId());
    sourceFolder = file.getParents().next();
    sourceFolderId = sourceFolder.getId();
    parentFolderId = sourceFolder.getParents().next().getId();
  }

  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);

  if (prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  } else {
    studentName = prompt.getResponseText();
  }

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();

  copyFolder(sourceFolderId, newFolderId, studentName, 'act');
  linkSheets(newFolderId, studentName, 'act');

  var htmlOutput = HtmlService.createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + studentName + "'s folder</a>")
    .setWidth(250)
    .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ACT folder created successfully');
}

function NewTestPrepFolder(sourceFolderId, parentFolderId) {
  let ss, file, sourceFolder, studentName;
  if (sourceFolderId === undefined || parentFolderId === undefined) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    file = DriveApp.getFileById(ss.getId());
    sourceFolder = file.getParents().next();
    sourceFolderId = sourceFolder.getId();
    parentFolderId = sourceFolder.getParents().next().getId();
  }

  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);

  if (prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  } else {
    studentName = prompt.getResponseText();
  }

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();

  copyFolder(sourceFolderId, newFolderId, studentName, 'all');
  linkSheets(newFolderId, studentName, 'all');

  var htmlOutput = HtmlService.createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + studentName + "'s folder</a>")
    .setWidth(250)
    .setHeight(50);
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
  } else if (folderType.toLowerCase() === 'act') {
    var testType = 'ACT';
  } else {
    var testType = 'Test';
  }

  while (files.hasNext()) {
    var file = files.next();
    let prefixFiles = ['Tutoring notes', 'ACT review sheet', 'SAT review sheet'];
    var fileName = file.getName();
    Logger.log(fileName);

    if (prefixFiles.includes(fileName)) {
      fileName = studentName + ' ' + fileName;
    } else if (fileName.toLowerCase().includes('template')) {
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
    } else if (testType === 'ACT' && fileName.toLowerCase().includes('sat') && fileName.toLowerCase().includes('answer analysis')) {
      newFile.setTrashed(true);
    }

    if (newFolderName.includes(folderType.toUpperCase()) && !newFolderName.includes(studentName)) {
      newFile.moveTo(newFolder.getParents().next());
      Logger.log('new location for ' + newFileName + ': ' + newFile.getParents().next().getId());
      if (isEmptyFolder(newFolder.getId())) {
        newFolder.setTrashed(true);
        Logger.log(newFolderName + ' trashed');
      }
    }
  }

  while (sourceSubFolders.hasNext()) {
    var sourceSubFolder = sourceSubFolders.next();
    var folderName = sourceSubFolder.getName();

    if (folderName === 'Student') {
      var targetFolder = newFolder.createFolder(studentName + ' ' + testType + ' prep');
    } else if (newFolderName.includes(folderType.toUpperCase()) && newFolderName !== studentName + ' ' + testType + ' prep') {
      var targetFolder = newFolder.getParents().next().createFolder(folderName);
      Logger.log(sourceSubFolder.getName() + ' moved');
    } else {
      var targetFolder = newFolder.createFolder(folderName);
    }

    targetFolderName = targetFolder.getName();

    if (targetFolderName.includes('ACT') && folderType.toLowerCase() === 'sat') {
      targetFolder.setTrashed(true);
      Logger.log(targetFolderName + ' trashed');
    } else if (targetFolderName.includes('SAT') && folderType.toLowerCase() === 'act') {
      targetFolder.setTrashed(true);
      Logger.log(targetFolderName + ' trashed');
    } else {
      copyFolder(sourceSubFolder.getId(), targetFolder.getId(), studentName, folderType);
    }
  }
}

var satSheetIds = {
  admin: null,
  student: null,
  studentData: null,
  adminData: null,
};

var satSheetDataUrls = {
  admin: null,
  student: null,
};

var actSheetIds = {
  admin: null,
  student: null,
  studentData: null,
  adminData: null,
};

var actSheetDataUrls = {
  admin: null,
  student: null,
};

function linkSheets(folderId, studentName='', prepType='all') {
  let folder = DriveApp.getFolderById(folderId);
  let files = folder.getFiles();
  let subFolders = DriveApp.getFolderById(folderId).getFolders();

  while (files.hasNext()) {
    let file = files.next();
    let fileName = file.getName();
    let fileId = file.getId();
    if (fileName.includes('SAT') && prepType !== 'act') {
      if (fileName.toLowerCase().includes('student answer sheet')) {
        satSheetIds.student = fileId;
        DriveApp.getFileById(satSheetIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      }
      else if (fileName.toLowerCase().includes('answer analysis')) {
        satSheetIds.admin = fileId;
        let ss = SpreadsheetApp.openById(fileId);

        for (i in ss.getSheets()) {
          let s = ss.getSheets()[i];
          let sName = s.getName();
          let answerSheets = getTestCodes(ss);
          answerSheets.push('Reading & Writing', 'Math', 'SLT Uniques');
          Logger.log(answerSheets);

          if (sName.toLowerCase().includes('analysis') || sName.toLowerCase().includes('opportunity')) {
            s.getRange('D4').setValue('for ' + studentName);
          }
          else if (answerSheets.includes(sName)) {
            Logger.log(sName);
            let protections = s.getProtections(SpreadsheetApp.ProtectionType.SHEET);
            for (var p = 0; p < protections.length; p++) {
              protections[p].remove();
            }
          }
        }
        let revBackend = ss.getSheetByName('Rev sheet backend');
        revBackend.getRange('K2').setValue(studentName);
        revBackend.protect().setWarningOnly(true);
      }
    }

    if (fileName.includes('ACT') && prepType !== 'sat') {
      if (fileName.toLowerCase().includes('student answer sheet')) {
        actSheetIds.student = file.getId();
        DriveApp.getFileById(actSheetIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      }
      else if (fileName.toLowerCase().includes('answer analysis')) {
        actSheetIds.admin = file.getId();
      }
    }
  }

  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    linkSheets(subFolder.getId(), studentName);
  }

  if (satSheetIds.student && satSheetIds.admin) {
    let satAdminSheet = SpreadsheetApp.openById(satSheetIds.admin);
    let satStudentSheet = SpreadsheetApp.openById(satSheetIds.student);
    satAdminSheet.getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);

    let revDataId = satAdminSheet.getSheetByName('Rev sheet backend').getRange('U3').getValue();
    let revDataSheet = SpreadsheetApp.openById(revDataId);

    let studentRevDataSheet = revDataSheet.getSheetByName(studentName);
    Logger.log('studentRevDataSheet: ' + studentRevDataSheet);
    if (!studentRevDataSheet) {
      try {
        studentRevDataSheet = revDataSheet.getSheetByName('Template').copyTo(revDataSheet).setName(studentName);
      }
      catch (err) {
        let ui = SpreadsheetApp.getUi();
        let continueScript = ui.prompt('Rev data sheet with same student name already exists. All students must have unique names for rev sheets to work properly. Are you re-creating this folder for an existing student?', ui.ButtonSet.YES_NO);

        if (continueScript === ui.Button.NO) {
          let htmlOutput = HtmlService.createHtmlOutput('<p>Please use a unique name for the new student or delete/rename the "'+ studentName + '" sheet from your <a href="https://docs.google.com/spreadsheets/d/' + revDataId + '">Rev sheet data</a></p>')
            .setWidth(400)
          SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Duplicate student name');
          return;
        }
      }
    }

    let studentQBSheet = satStudentSheet.getSheetByName('Question bank data');
    studentQBSheet.getRange('U2').setValue(studentName);
    studentQBSheet.getRange('U4').setValue(satSheetIds.admin);

    SpreadsheetApp.openById(satSheetIds.admin).getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);
  }

  Logger.log('actSheetIds.student: ' + actSheetIds.student);
  Logger.log('actSheetIds.admin: ' + actSheetIds.admin);

  if (actSheetIds.student && actSheetIds.admin) {
    SpreadsheetApp.openById(actSheetIds.admin).getSheetByName('Student responses').getRange('B1').setValue(actSheetIds.student);
  }
}

function isEmptyFolder(folderId) {
  const folders = DriveApp.getFolderById(folderId).getFolders();
  const files = DriveApp.getFolderById(folderId).getFiles();

  if (folders.hasNext() || files.hasNext()) {
    return false;
  } else {
    return true;
  }
}

function createRwRevSheet() {
  createRevSheet('RW', 0);
}

function createMathRevSheet() {
  createRevSheet('Math', 1);
}

function createRevSheet(sub, subIndex) {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let revSheet = ss.getSheetByName(sub + ' Rev sheet');
    let revResponseSheet = ss.getSheetByName('Rev sheets');
    let subBackendOffset = subIndex * 4;
    let revBackend = ss.getSheetByName('Rev sheet backend');
    let revSubjectFolderIdCell = revBackend.getRange(2, 3 + subBackendOffset);
    let revSubjectSortStart = revBackend.getRange(2, 1 + subBackendOffset).getValue();
    let revSubjectFolderId = revSubjectFolderIdCell.getValue();
    let revSheetFolderIdCell = revBackend.getRange('U2');
    let revSheetFolderId = revSheetFolderIdCell.getValue();
    // let folderKeyIdCell = revBackend.getRange(3, 3 + subBackendOffset);
    // let revKeySubjectFolderId = folderKeyIdCell;
    let studentName = revBackend.getRange('K2').getValue();
    let revDataSsId = revBackend.getRange('U3').getValue()
    let revDataSs = SpreadsheetApp.openById(revDataSsId);
    let revDataSheet = revDataSs.getSheetByName(studentName);
    let revSheetFolder, satFolder, studentFolder;

    if (!revSubjectSortStart) {
      let ui = SpreadsheetApp.getUi();
      ui.alert('Error: No missed questions available for ' + revResponseSheet.getRange(1, 3 + subIndex * 5).getValue());
      return;
    }

    let maxQuestionRange = revBackend.getRange('L2');
    let ui = SpreadsheetApp.getUi();
    let prompt = ui.prompt('Max # of questions - leave blank to use prior value of ' + maxQuestionRange.getValue(), ui.ButtonSet.OK_CANCEL);
    if (prompt.getSelectedButton() == ui.Button.CANCEL) {
      return;
    } else if (prompt.getResponseText() !== '') {
      maxQuestionRange.setValue(prompt.getResponseText());
    }

    try {
      revSheetFolder = DriveApp.getFolderById(revSheetFolderId);
    } 
    catch {
      Logger.log('Rev folder ID ' + revSheetFolderId + ' not found');
      revSheetFolderId = '';
      revSheetFolderIdCell.setValue(revSheetFolderId);
    }

    try {
      DriveApp.getFolderById(revSubjectFolderId);
    } 
    catch {
      Logger.log('Rev subject folder ID ' + revSubjectFolderId + ' not found');
      revSubjectFolderId = '';
      revSubjectFolderIdCell.setValue(revSubjectFolderId);
    }

    // try {
    //   DriveApp.getFolderById(revKeySubjectFolderId);
    // } 
    // catch {
    //   Logger.log('Key folder ID ' + revKeySubjectFolderId + ' not found');
    //   revKeySubjectFolderId = '';
    //   revSubjectFolderIdCell.setValue(revKeySubjectFolderId);
    // }

    try {
      if (!revSubjectFolderId) {
        if (sub === 'RW') {
          var subject = 'Reading & Writing';
        } else {
          var subject = 'Math';
        }

        if (revSheetFolderId) {
          revSubjectFolderId = revSheetFolder.createFolder(subject).getId();
        }
        else {
          let adminFolder = DriveApp.getFileById(ss.getId()).getParents().next();
          let adminSubfolders = adminFolder.getFolders();

          if (adminSubfolders.hasNext()) {
            while (adminSubfolders.hasNext()) {
              let adminSubfolder = adminSubfolders.next();
              let adminSubfolderName = adminSubfolder.getName();

              if (adminSubfolderName.includes('SAT')) {
                satFolder = adminSubfolder;
                revSheetFolder = satFolder.createFolder('Rev sheets');
                revSheetFolderId = revSheetFolder.getId();
                revSubjectFolderId = revSheetFolder.createFolder(subject).getId();
                break;
              }
              else if (adminSubfolderName.toLowerCase().includes(studentName.toLowerCase())) {
                studentFolder = adminSubfolder;
                let studentSubfolders = studentFolder.getFolders();
                while (studentSubfolders.hasNext()) {
                  let studentSubfolder = studentSubfolders.next();
                  let studentSubfolderName = studentSubfolder.getName();

                  if (studentSubfolderName.includes('SAT')) {
                    satFolder = studentSubfolder;
                    revSheetFolder = satFolder.createFolder('Rev sheets');
                    revSheetFolderId = revSheetFolder.getId();
                    revSubjectFolderId = revSheetFolder.createFolder(subject).getId();
                  }
                }
              }
            }
          }

          if (!revSheetFolderId) {
            revSheetFolder = adminFolder.createFolder('Rev sheets');
            revSheetFolderId = revSheetFolder.getId();
            revSubjectFolderId = revSheetFolder.createFolder(subject).getId();
          }

          revSheetFolderIdCell.setValue(revSheetFolderId);
        }

        revSubjectFolderIdCell.setValue(revSubjectFolderId);
      }
    }
    catch(err) {
      Logger.log(err.stack);
      let htmlOutput = HtmlService.createHtmlOutput('<p>Rev sheet error: ' + err.stack + '</p><p>Please copy this error and send to danny@openpathtutoring.com.</p>')
      .setWidth(400)
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error');
      return;
    }

    // if (!revKeySubjectFolderId) {
    //   if (revSheetFolder === adminSubFolder) {

    //   }
    //   revKeySubjectFolderId = adminFolder.createFolder('Rev sheet answer keys')
    // }

    revSheet.showSheet();
    revSheet.showRows(1, revSheet.getMaxRows());
    revBackend.getRange(2, 2 + subBackendOffset, revBackend.getLastRow() - 1).clear();
    revBackend.getRange(2, 2 + subBackendOffset).setValue('=RANDARRAY(counta(A$2:A))');
    SpreadsheetApp.flush();
    revBackend.getRange(2, 2 + subBackendOffset, revBackend.getMaxRows() - 1).copyValuesToRange(revBackend.getSheetId(), 2 + subBackendOffset, 2 + subBackendOffset, 2, 2);

    var idCol = revSheet.getRange('B1:B');
    var values = idCol.getValues(); // get all data in one call
    var heights = revSheet.getRange('E1:E');
    var heightVals = heights.getValues();
    //var imgContainerWidth = revSheet.getColumnWidth(4);
    var row = 6;

    try {
      while (values[row - 1] && values[row - 1][0] != '') {
        var questionId = values[row - 1][0];
        var rowHeight = heightVals[row - 1][0]; // rowHeights hard-coded in Rev sheet backend
        revSheet.setRowHeight(row, rowHeight);
        Logger.log(questionId + ' rowHeight: ' + rowHeight);
        row++;
      }
    } catch (err) {
      if (err.message.includes('Invalid argument')) {
        SpreadsheetApp.getUi().alert('Error: Image not found');
      } else {
        SpreadsheetApp.getUi().alert(err);
      }
      return;
    }

    var firstEmptyRow = getFirstEmptyRow(revDataSheet, 2 + subIndex * 3);
    if (firstEmptyRow === 5) {
      var newRevSheetNumber = 1;
    } else {
      var revSheetLastQuestion = revDataSheet
        .getRange(firstEmptyRow - 1, 2 + subIndex * 3)
        .getValue()
        .toString();
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
    } else {
      var pdfName = sub + ' Rev sheet #' + newRevSheetNumber + ' for ' + studentName;
    }

    //* Create worksheets
    SpreadsheetApp.flush();
    savePdf(ss, revSheet, pdfName, revSubjectFolderId);
    Logger.log(sub + ' Rev sheet #' + newRevSheetNumber + ' saved');
    //*/

    /* Create answer keys
    revSheet.showColumns(3);
    revSheet.showColumns(6);
    revSheet.hideColumns(5);

    SpreadsheetApp.flush();
    savePdf(ss, revSheet, pdfName + '~Key', revSubjectFolderId);
    Logger.log(sub + ' Rev key #' + newRevSheetNumber + ' saved');
    //*/

    var dataToCopy = revSheet.getRange(6, 1, row - 5, 2).getValues();
    revDataSheet.getRange(firstEmptyRow, 2 + subIndex * 3, row - 5, 2).setValues(dataToCopy);

    revSheet.showRows(1, revSheet.getMaxRows());
    revSheet.hideSheet();

    let htmlOutput = HtmlService.createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + revSubjectFolderId + '" target="_blank" onclick="google.script.host.close()">' + sub + ' Rev sheet folder</a>')
      .setWidth(250) //optional
      .setHeight(50); //optional
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Rev sheet complete');
  } catch (err) {
    let htmlOutput = HtmlService.createHtmlOutput(err.stack).setWidth(400); //optional
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error');
    Logger.log(err.stack);
  }
}

function getRevSubjectFolderId(revSheetFolder) {
  let revSubjectFolderId;

  while (revSheetFolder.hasNext()) {
    let subfolder = revSheetFolder.next();
    let subfolderName = subfolder.getName();
    if (subfolderName.toLowerCase().includes(subject.toLowerCase())) {
      revSubjectFolderId = subfolder.getId();
      break;
    }
  }
  if (!revSubjectFolderId) {
    revSubjectFolderId = revSheetFolder.createFolder(subject).getId();
  }

  return revSubjectFolderId;
}

function savePdf(spreadsheet, sheet, pdfName, pdfFolderId) {
  var sheetId = sheet.getSheetId();
  var url_base = spreadsheet.getUrl().replace(/edit$/, '');

  var url_ext =
    'export?exportFormat=pdf&format=pdf' +
    '&gid=' +
    sheetId +
    // following parameters are optional...
    '&size=A4' + // paper size: legal / letter / A4
    '&portrait=true' + // orientation, false for landscape
    '&fitw=true' + // fit to width, false for actual size
    '&top_margin=0.25' +
    '&bottom_margin=0' +
    '&left_margin=0.375' +
    '&right_margin=0.375' +
    '&sheetnames=false' + //
    '&printtitle=false' +
    '&pagenumbers=false' + //hide optional headers and footers
    '&gridlines=false' + // hide gridlines
    '&fzr=true'; // false = do not repeat row headers (frozen rows) on each page
  var url_options = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  var response = (function backoff(i) {
    Utilities.sleep(Math.pow(2, i) * 1000);
    let data = UrlFetchApp.fetch(url_base + url_ext, url_options);
    if (data.getResponseCode() !== 200) {
      return backoff(++i);
    } else {
      return data;
    }
  })(1);
  var blob = response
    .getBlob()
    .getAs('application/pdf')
    .setName(pdfName + '.pdf');
  var folder = DriveApp.getFolderById(pdfFolderId);
  folder.createFile(blob);
}

function transferOldStudentData() {
  let ui = SpreadsheetApp.getUi();
  let prompt = ui.prompt('Old admin analysis spreadsheet URL or ID:', ui.ButtonSet.OK_CANCEL);
  let oldAdminDataUrl = prompt.getResponseText();
  if (prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  let oldSsId;
  if (oldAdminDataUrl.includes('/d/')) {
    oldSsId = oldAdminDataUrl.split('/d/')[1].split('/')[0];
  }
  else {
    oldSsId = oldAdminDataUrl;
  }

  transferStudentData(oldSsId);
}

function transferStudentData(oldSsId) {
  let newAdminSs = SpreadsheetApp.getActiveSpreadsheet();

  newStudentSsId = newAdminSs.getSheetByName('Student responses').getRange('B1').getValue();
  let newStudentSs = SpreadsheetApp.openById(newStudentSsId);
  
  let ui = SpreadsheetApp.getUi();
  let confirm = ui.alert('Warning: This script will overwrite data in the new student spreadsheet. Continue?', ui.ButtonSet.YES_NO);
  if (confirm === ui.Button.NO) {
    return;
  }

  let oldSs, newStudentData, initialImportFunction;
  try {
    oldSs = SpreadsheetApp.openById(oldSsId);
    newStudentData = newAdminSs.getSheetByName('Student responses');

    // temporarily set old admin data imports
    initialImportFunction = newStudentData.getRange('A3').getFormula();
    newStudentData.getRange('A3').setValue('=importrange("' + oldSsId + '", "Question bank data!$A$1:$G10000")');
    newStudentData.getRange('H3').setValue('=importrange("' + oldSsId + '", "Question bank data!$I$1:$K10000")');
    DriveApp.getFileById(oldSsId).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    DriveApp.getFileById(newStudentSsId).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);

    let answerSheets = getTestCodes(oldSs);
    let testScores = [];

    for (let s in answerSheets) {
      let sheetName = answerSheets[s];
      let oldSheet = oldSs.getSheetByName(sheetName);
      if (oldSheet) {
        let subScore = oldSheet.getRange('G1:I1').getValues();
        testScores.push({
          'test': sheetName,
          'scores': subScore
        })
      }
    }

    answerSheets.push('Reading & Writing', 'Math', 'SLT Uniques');

    for (let s in answerSheets) {
      let sheetName = answerSheets[s];
      let newSheet = newAdminSs.getSheetByName(sheetName);
      let newStudentSheet = newStudentSs.getSheetByName(sheetName);

      if (newSheet) {
        Logger.log('Transferring answers for ' + newSheet.getName());
        let newAnswersLevel1 = newSheet.getRange('C5:C');
        let newAnswersLevel2 = newSheet.getRange('G5:G');
        let newAnswersLevel3 = newSheet.getRange('K5:K');
        let newStudentLevel1 = newStudentSheet.getRange('C5:C');
        let newStudentLevel2 = newStudentSheet.getRange('G5:G');
        let newStudentLevel3 = newStudentSheet.getRange('K5:K');
        let newRanges = [newAnswersLevel1, newAnswersLevel2, newAnswersLevel3];
        let newStudentRanges = [newStudentLevel1, newStudentLevel2, newStudentLevel3];

        for (let i = 0; i < newRanges.length; i++) {
          // let newSheetFormulas = newRanges[i].getFormulas();
          let newSheetValues = newRanges[i].getValues();

          for (let row = 0; row < newSheetValues.length; row++) {
            if (newSheetValues[row][0] === 'not found') {
              newSheetValues[row][0] = '';
            }
          }

          newStudentRanges[i].setValues(newSheetValues);

          let testScore = testScores.find(score => score.test === sheetName);
          if (testScore) {
            newSheet.getRange('G1:I1').setValues(testScore.scores);
          }
        }
      }
    }

    // set data to new student SS
    // for (let s in answerSheets) {
    //   let sheetName = answerSheets[s];
    //   let newStudentSheet = newAdminSs.getSheetByName(sheetName);

    //   if (newStudentSheet) {

    //     for (let i = 0; i < newStudentRanges.length; i++) {
          
    //     }
    //   }
    // }
    
    // build timestamp column
    let newQbSheet = newAdminSs.getSheetByName('Question bank data');
    let timestampLookup = '=xlookup(A2, \'Student responses\'!$A$4:$A$10000, \'Student responses\'!$J$4:$J$10000,"")';
    let timestampStartCell = newQbSheet.getRange('K2');
    timestampStartCell.setValue(timestampLookup);
    let timestampRange = newQbSheet.getRange('K2:K10000');
    timestampStartCell.autoFill(timestampRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    let timestampValues = timestampRange.getValues();

    for (let row = 0; row < timestampValues.length; row ++) {
      let ssRow = row + 2;
      if (timestampValues[row][0] === '') {
        Logger.log('blank row: ' + ssRow);
        timestampValues[row][0] = '=if(or(G' + ssRow + '="",I' + ssRow + '=""),"",if(K' + ssRow + ',K' + ssRow + ',if(I' + ssRow + '="","",now())))'
      }
    }
    Logger.log(timestampValues);
    timestampRange.setValues(timestampValues);
    timestampRange.setNumberFormat('MM/dd/yyy h:mm:ss');

    let htmlOutput = HtmlService.createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newStudentSsId + '" target="_blank" onclick="google.script.host.close()">Student answer sheet</a>')
    .setWidth(250)
    .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Data transfer complete');
  }
  catch (err) {
    let htmlOutput = HtmlService.createHtmlOutput('<p>Error while processing data: ' + err.stack + '</p><p>Please copy this error and send to danny@openpathtutoring.com.</p>')
      .setWidth(400)
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error');
  }

  // revert student ID and SS permissions
  newStudentData.getRange('A3').setValue(initialImportFunction);
  newStudentData.getRange('H3').setValue('');
  DriveApp.getFileById(oldSsId).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  DriveApp.getFileById(newStudentSsId).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
}

function getLastFilledRow(sheet, col) {
  const lastRow = sheet.getLastRow();
  const allVals = sheet.getRange(1, col, lastRow).getValues();
  const lastFilledRow = lastRow - allVals.reverse().findIndex((c) => c[0] != '');

  return lastFilledRow;
}

function getTestCodes() {
  const practiceTestDataSheet = SpreadsheetApp.openById('1KidSURXg5y-dQn_gm1HgzUDzaICfLVYameXpIPacyB0').getSheetByName('Practice test data');
  const lastFilledRow = getLastFilledRow(practiceTestDataSheet, 1);
  const testCodeCol = practiceTestDataSheet
    .getRange(2, 1, lastFilledRow - 1)
    .getValues()
    .map((row) => row[0]);
  const testCodes = testCodeCol.filter((x, i, a) => a.indexOf(x) == i);

  return testCodes;
}

// Adapted from https://stackoverflow.com/a/9102463/1677912
function getFirstEmptyRow(sheet, colIndex) {
  var column = sheet.getRange(5, colIndex, sheet.getRange('A1:A').getLastRow() - 4);
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while (values[ct] && values[ct][0] != '') {
    ct++;
  }
  return ct + 5; // +5 since starting from row 5 with 0-indexing
}
