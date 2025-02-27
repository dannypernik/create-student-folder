function getFolderIds(sourceFolderId, parentFolderId) {
  if (sourceFolderId === undefined || parentFolderId === undefined) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const file = DriveApp.getFileById(ss.getId());
    const sourceFolder = file.getParents().next();
    sourceFolderId = sourceFolder.getId();
    parentFolderId = sourceFolder.getParents().next().getId();
  }
  return { sourceFolderId, parentFolderId };
}

function NewSatFolder(sourceFolderId, parentFolderId) {
  const ids = getFolderIds(sourceFolderId, parentFolderId);
  sourceFolderId = ids.sourceFolderId;
  parentFolderId = ids.parentFolderId;

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
  const ids = getFolderIds(sourceFolderId, parentFolderId);
  sourceFolderId = ids.sourceFolderId;
  parentFolderId = ids.parentFolderId;

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
  const ids = getFolderIds(sourceFolderId, parentFolderId);
  sourceFolderId = ids.sourceFolderId;
  parentFolderId = ids.parentFolderId;

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

  let fileOperations = [];

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
      fileOperations.push({ file: newFile, action: 'move' });
      Logger.log('new location for ' + newFileName + ': ' + newFile.getParents().next().getId());
      if (isEmptyFolder(newFolder.getId())) {
        newFolder.setTrashed(true);
        Logger.log(newFolderName + ' trashed');
      }
    }
  }

  // Perform file operations in batch
  fileOperations.forEach(op => {
    if (op.action === 'move') {
      op.file.moveTo(newFolder.getParents().next());
    }
  });

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
  let subFolders = folder.getFolders();
  const SERVICE_ACCOUNT_EMAIL = 'score-reports@sat-score-reports.iam.gserviceaccount.com';

  let satFiles = [];
  let actFiles = [];

  while (files.hasNext()) {
    let file = files.next();
    let fileName = file.getName();
    let fileId = file.getId();

    if (fileName.includes('SAT') && prepType !== 'act') {
      satFiles.push({ fileName, fileId });
    } else if (fileName.includes('ACT') && prepType !== 'sat') {
      actFiles.push({ fileName, fileId });
    }
  }

  satFiles.forEach(({ fileName, fileId }) => {
    if (fileName.toLowerCase().includes('student answer sheet')) {
      satSheetIds.student = fileId;
      let satStudentSheet = DriveApp.getFileById(satSheetIds.student);
      satStudentSheet.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      satStudentSheet.addEditor(SERVICE_ACCOUNT_EMAIL);
    } else if (fileName.toLowerCase().includes('answer analysis')) {
      satSheetIds.admin = fileId;
      let ss = SpreadsheetApp.openById(fileId);
      DriveApp.getFileById(satSheetIds.admin).addEditor(SERVICE_ACCOUNT_EMAIL);

      ss.getSheets().forEach(s => {
        let sName = s.getName();
        let answerSheets = getTestCodes(ss);
        answerSheets.push('Reading & Writing', 'Math', 'SLT Uniques');

        if (answerSheets.includes(sName)) {
          s.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
        }
      });

      let revBackend = ss.getSheetByName('Rev sheet backend');
      revBackend.getRange('K2').setValue(studentName);
    }
  });

  actFiles.forEach(({ fileName, fileId }) => {
    if (fileName.toLowerCase().includes('student answer sheet')) {
      actSheetIds.student = fileId;
      let actStudentSheet = DriveApp.getFileById(actSheetIds.student);
      actStudentSheet.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    } else if (fileName.toLowerCase().includes('answer analysis')) {
      actSheetIds.admin = fileId;
      const ss = SpreadsheetApp.openById(fileId);
      ss.getSheetByName('Student responses').getRange('G1').setValue(studentName);
    }
  });

  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    linkSheets(subFolder.getId(), studentName, prepType); // Added prepType to recursive call
    if (prepType === 'all' && satSheetIds.student && satSheetIds.admin && actSheetIds.student && actSheetIds.admin) {
      break;
    }
    else if (prepType === 'sat' && satSheetIds.student && satSheetIds.admin) {
      break;
    }
    else if (prepType === 'act' && actSheetIds.student && actSheetIds.admin) {
      break;
    }
  }

  if (satSheetIds.student && satSheetIds.admin) {
    let satAdminSheet = SpreadsheetApp.openById(satSheetIds.admin);
    let satStudentSheet = SpreadsheetApp.openById(satSheetIds.student);
    satAdminSheet.getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);

    let revDataId = satAdminSheet.getSheetByName('Rev sheet backend').getRange('U3').getValue();
    let revDataSheet = SpreadsheetApp.openById(revDataId);

    let studentRevDataSheet = revDataSheet.getSheetByName(studentName);
    if (!studentRevDataSheet) {
      try {
        studentRevDataSheet = revDataSheet.getSheetByName('Template').copyTo(revDataSheet).setName(studentName);
      } catch (err) {
        let ui = SpreadsheetApp.getUi();
        let continueScript = ui.prompt('Rev data sheet with same student name already exists. All students must have unique names for rev sheets to work properly. Are you re-creating this folder for an existing student?', ui.ButtonSet.YES_NO);

        if (continueScript === ui.Button.NO) {
          let htmlOutput = HtmlService.createHtmlOutput('<p>Please use a unique name for the new student or delete/rename the "'+ studentName + '" sheet from your <a href="https://docs.google.com/spreadsheets/d/' + revDataId + '">Rev sheet data</a></p>')
            .setWidth(400);
          SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Duplicate student name');
          return;
        }
      }
    }

    let studentQBSheet = satStudentSheet.getSheetByName('Question bank data');
    studentQBSheet.getRange('U2').setValue(studentName);
    studentQBSheet.getRange('U4').setValue(satSheetIds.admin);

    satAdminSheet.getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);
  }

  if (actSheetIds.student && actSheetIds.admin) {
    let actAdminSheet = SpreadsheetApp.openById(actSheetIds.admin);
    actAdminSheet.getSheetByName('Student responses').getRange('B1').setValue(actSheetIds.student);
  }
}

function isEmptyFolder(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  return !folder.getFiles().hasNext() && !folder.getFolders().hasNext();
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
    let revBackend = ss.getSheetByName('Rev sheet backend');
    let maxQuestionRange = revBackend.getRange('L2');
    let ui = SpreadsheetApp.getUi();

    let prompt = ui.prompt('Max # of questions - leave blank to use prior value of ' + maxQuestionRange.getValue(), ui.ButtonSet.OK_CANCEL);
    if (prompt.getSelectedButton() == ui.Button.CANCEL) {
      return;
    }
    else if (prompt.getResponseText() !== '') {
      maxQuestionRange.setValue(prompt.getResponseText());
    }

    let subBackendOffset = subIndex * 4;
    let revSubjectSortStart = revBackend.getRange(2, 1 + subBackendOffset).getValue();
    let revResponseSheet = ss.getSheetByName('Rev sheets');

    if (!revSubjectSortStart) {
      ui.alert('Error: No missed questions available for ' + revResponseSheet.getRange(1, 3 + subIndex * 5).getValue());
      return;
    }

    let adminFolder = DriveApp.getFileById(ss.getId()).getParents().next();
    let revSheet = ss.getSheetByName(sub + ' Rev sheet');
    let revSubjectFolderIdCell = revBackend.getRange(2, 3 + subBackendOffset);
    let revSubjectFolderId = revSubjectFolderIdCell.getValue();
    let revSheetFolderIdCell = revBackend.getRange('U2');
    let revSheetFolderId = revSheetFolderIdCell.getValue();
    let revKeyFolderCell = revBackend.getRange('U4');
    let revKeyFolderId = revKeyFolderCell.getValue();
    let revKeySubjectFolderCell = revBackend.getRange(3, 3 + subBackendOffset);
    let revKeySubjectFolderId = revKeySubjectFolderCell.getValue();
    let studentName = revBackend.getRange('K2').getValue();
    let revDataSsId = revBackend.getRange('U3').getValue()
    let revDataSs = SpreadsheetApp.openById(revDataSsId);
    let revDataSheet = revDataSs.getSheetByName(studentName);
    let studentSsId = ss.getSheetByName('Student responses').getRange('B1').getValue();
    let studentSs = SpreadsheetApp.openById(studentSsId);
    let revSheetFolder, revKeyFolder, satFolder, studentFolder, subject;

    if (sub === 'RW') {
      subject = 'Reading & Writing';
    } else {
      subject = 'Math';
    }

    Logger.log('variables set');

    if (!revDataSheet) {
      revDataSheet = revDataSs.getSheetByName('Template').copyTo(revDataSs).setName(studentName);
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

    Logger.log('ids set');

    try {
      if (!revSubjectFolderId) {
        if (revSheetFolderId) {
          revSubjectFolderId = revSheetFolder.createFolder(subject).getId();
        }
        else {
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
        studentSs.getSheetByName('Question bank data').getRange(5 + subIndex, 21).setValue(revSubjectFolderId);
      }
    }
    catch(err) {
      Logger.log(err.stack);
      let htmlOutput = HtmlService.createHtmlOutput('<p>Rev sheet error: ' + err.stack + '</p><p>Please copy this error and send to danny@openpathtutoring.com.</p>')
      .setWidth(400)
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error');
      return;
    }

    Logger.log('Rev folder logic complete');

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

    Logger.log('starting rowHeights');

    try {
      while (values[row - 1] && values[row - 1][0] != '') {
        var questionId = values[row - 1][0];
        var rowHeight = heightVals[row - 1][0] + subIndex * 100; // rowHeights hard-coded in Rev sheet backend + extra for Math
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

    let revDataLastQuestionCell, revDataLastQuestion, newRevSheetNumber;
    let revDataSubjectColumn = 2 + subIndex * 3;
    let revResponseSubjectColumn = 4 + subIndex * 5;
    let lastFilledQuestionRow = getLastFilledRow(revDataSheet, revDataSubjectColumn);
    let lastFilledResponseRow = getLastFilledRow(revResponseSheet, revResponseSubjectColumn);

    if (lastFilledQuestionRow === 4) {
      newRevSheetNumber = 1;
    }
    else {
      revDataLastQuestionCell = revDataSheet.getRange(lastFilledQuestionRow, revDataSubjectColumn)
      revDataLastQuestion =  revDataLastQuestionCell.getValue().toString();
      Logger.log('revDataLastQuestion' + revDataLastQuestion);
      newRevSheetNumber = parseInt(revDataLastQuestion.substring(revDataLastQuestion.lastIndexOf(' ') + 1, revDataLastQuestion.indexOf('.'))) + 1;
      revResponseSheet.getRange(4, revResponseSubjectColumn, lastFilledResponseRow - 3).copyTo(revResponseSheet.getRange(4, revResponseSubjectColumn), {contentsOnly: true});
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

    try {
      revKeyFolder = DriveApp.getFolderById(revKeyFolderId);
    }
    catch {
      Logger.log('Key folder ID ' + revKeyFolderId + ' not found');
      revKeyFolder = adminFolder.createFolder('Rev keys');
      revKeyFolderCell.setValue(revKeyFolder.getId());
      revKeySubjectFolderId = revKeyFolder.createFolder(subject).getId();
      revKeySubjectFolderCell.setValue(revKeySubjectFolderId);
    }

    try {
      DriveApp.getFolderById(revKeySubjectFolderId);
    }
    catch {
      Logger.log('Key subject folder ID ' + revKeySubjectFolderId + ' not found');
      revKeySubjectFolderId = revKeyFolder.createFolder(subject).getId();
      revKeySubjectFolderCell.setValue(revKeySubjectFolderId);
    }

    //* Create answer keys
    revSheet.showColumns(3);
    revSheet.showColumns(6);
    revSheet.hideColumns(5);

    SpreadsheetApp.flush();
    savePdf(ss, revSheet, pdfName + '~Key', revKeySubjectFolderId);
    Logger.log(sub + ' Rev key #' + newRevSheetNumber + ' saved');
    //*/

    var dataToCopy = revSheet.getRange(6, 1, row - 5, 2).getValues();
    revDataSheet.getRange(lastFilledQuestionRow + 1, revDataSubjectColumn, row - 5, 2).setValues(dataToCopy);

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
    '&gid=' + sheetId +
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
  const startTime = new Date().getTime(); // Record the start time
  let ui = SpreadsheetApp.getUi();
  let prompt = ui.prompt(
    'Old admin analysis spreadsheet URL or ID - leave blank \r\n' +
    'to update student sheet with this admin spreadsheet\'s data:',
    ui.ButtonSet.OK_CANCEL);
  let oldAdminDataUrl = prompt.getResponseText();
  if (prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }

  let htmlOutput = HtmlService
      .createHtmlOutput('<p>If you manually cancel, you will need to restore the previous version of the spreadsheet by clicking File > Version history > See version history</p><button onclick="google.script.host.close()">OK</button>')
      .setWidth(400)
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Do not cancel this script');

  let oldAdminSsId;
  if (oldAdminDataUrl === '') {
    oldAdminSsId = SpreadsheetApp.getActiveSpreadsheet().getId();
  }
  else if (oldAdminDataUrl.includes('/d/')) {
    oldAdminSsId = oldAdminDataUrl.split('/d/')[1].split('/')[0];
  }
  else {
    oldAdminSsId = oldAdminDataUrl;
  }

  transferStudentData(oldAdminSsId, startTime);
}

function transferStudentData(oldAdminSsId, startTime) {
  const newAdminSs = SpreadsheetApp.getActiveSpreadsheet();
  const newStudentSsId = newAdminSs.getSheetByName('Student responses').getRange('B1').getValue();
  const newStudentSs = SpreadsheetApp.openById(newStudentSsId);
  const maxDuration = 5.5 * 60 * 1000; // 5 minutes and 30 seconds in milliseconds

  // temporarily relax permissions
  DriveApp.getFileById(oldAdminSsId).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  DriveApp.getFileById(newStudentSsId).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);

  let oldAdminSs, newStudentData, initialImportFunction;
  try {
    oldAdminSs = SpreadsheetApp.openById(oldAdminSsId);
    newStudentData = newAdminSs.getSheetByName('Student responses');
    initialImportFunction = newStudentData.getRange('A3').getFormula();

    if (oldAdminSsId !== newAdminSs.getId()) {
      // temporarily set old admin data imports
      newStudentData.getRange('A3').setValue('=importrange("' + oldAdminSsId + '", "Question bank data!$A$1:$G10000")');
      newStudentData.getRange('H3').setValue('=importrange("' + oldAdminSsId + '", "Question bank data!$I$1:$K10000")');

      // Copy rev data if necessary
      // issue: currently, unfilled rev answers get autofilled by old data before pasting values.
      // issue: new hard-coded values may not match old question codes
      // if (oldAdminSs.getSheetByName('Rev sheets')) {
      //   let oldRevBackend = oldAdminSs.getSheetByName('Rev sheet backend');
      //   let oldRevDataId = oldRevBackend.getRange('U3').getValue();
      //   let oldStudentName = oldRevBackend.getRange('K2').getValue();
      //   let oldRevDataStudentSheet = SpreadsheetApp.openById(oldRevDataId).getSheetByName(oldStudentName);
      //   let oldStudentRevData = oldRevDataStudentSheet.getRange(1,1,oldRevDataStudentSheet.getLastRow(), oldRevDataStudentSheet.getLastColumn()).getValues();
      //   let newRevBackend = newAdminSs.getSheetByName('Rev sheet backend');
      //   let newRevDataId = newRevBackend.getRange('U3').getValue();
      //   let newStudentName = newRevBackend.getRange('K2').getValue();
      //   let newRevDataStudentSheet = SpreadsheetApp.openById(newRevDataId).getSheetByName(newStudentName);
      //   newRevDataStudentSheet.getRange(1,1,oldRevDataStudentSheet.getLastRow(), oldRevDataStudentSheet.getLastColumn()).setValues(oldStudentRevData);
      // }
    }

    let answerSheets = getTestCodes(oldAdminSs);
    let testScores = [];

    for (let s in answerSheets) {
      let sheetName = answerSheets[s];
      let oldSheet = oldAdminSs.getSheetByName(sheetName);
      if (oldSheet) {
        let subScore = oldSheet.getRange('G1:I1').getValues();
        testScores.push({
          'test': sheetName,
          'scores': subScore
        })
      }
    }

    answerSheets.push('Reading & Writing', 'Math', 'SLT Uniques');

    let allNewAdminSheetValues = [];
    let allNewStudentSheetValues = [];

    for (let s in answerSheets) {

      let sheetName = answerSheets[s];
      let newAdminSheet = newAdminSs.getSheetByName(sheetName);
      let newStudentSheet = newStudentSs.getSheetByName(sheetName);

      if (newAdminSheet) {
        Logger.log('Transferring answers for ' + newAdminSheet.getName());
        let newAnswersLevel1, newAnswersLevel2, newAnswersLevel3;
        let newStudentLevel1, newStudentLevel2, newStudentLevel3;

        if (sheetName === 'Rev sheets') {
          newAnswersLevel1 = newAdminSheet.getRange(5, 4, getLastFilledRow(newAdminSheet, 3) - 4);
          newAnswersLevel2 = newAdminSheet.getRange(5, 9, getLastFilledRow(newAdminSheet, 8) - 4);
          newStudentLevel1 = newStudentSheet.getRange(5, 4, getLastFilledRow(newAdminSheet, 3) - 4);
          newStudentLevel2 = newStudentSheet.getRange(5, 9, getLastFilledRow(newAdminSheet, 8) - 4);
        }
        else {
          newAnswersLevel1 = newAdminSheet.getRange(5, 3, getLastFilledRow(newAdminSheet, 2) - 4);
          newAnswersLevel2 = newAdminSheet.getRange(5, 7, getLastFilledRow(newAdminSheet, 6) - 4);
          newStudentLevel1 = newStudentSheet.getRange(5, 3, getLastFilledRow(newAdminSheet, 2) - 4);
          newStudentLevel2 = newStudentSheet.getRange(5, 7, getLastFilledRow(newAdminSheet, 6) - 4);
          if (sheetName !== 'SLT Uniques') {
            newAnswersLevel3 = newAdminSheet.getRange(5, 11, getLastFilledRow(newAdminSheet, 10) - 4);
            newStudentLevel3 = newStudentSheet.getRange(5, 11, getLastFilledRow(newAdminSheet, 10) - 4);
          }
        }

        let newAdminRanges = [newAnswersLevel1, newAnswersLevel2, newAnswersLevel3];
        let newStudentRanges = [newStudentLevel1, newStudentLevel2, newStudentLevel3];

        for (let i = 0; i < newAdminRanges.length; i++) {
          if (newAdminRanges[i] && newStudentRanges[i]) {
            const currentTime = new Date().getTime();
            if (currentTime - startTime > maxDuration) {
              Logger.log("Exiting loop after 5 minutes and 30 seconds.");
              throw new Error("Process exceeded maximum duration of 5 minutes and 30 seconds. Please revert to previous version of this spreadsheet by clicking File > Version history > See versions.");
            }
            let newAdminSheetValues = newAdminRanges[i].getValues();
            let newAdminSheetFormulas = newAdminRanges[i].getFormulas();
            let newStudentSheetValues = [];

            for (let row = 0; row < newAdminSheetValues.length; row++) {
              // set blank cells blank for student sheet
              if (newAdminSheetValues[row][0] === 'not found') {
                newAdminSheetValues[row][0] = '';
              }
              // save nonblank cells as values for admin sheet
              else if (newAdminSheetValues[row][0] !== '') {
                newAdminSheetFormulas[row][0] = newAdminSheetValues[row][0];
              }

              newStudentSheetValues.push(newAdminSheetValues[row]);
            }
            // Ensure the number of rows in the source and destination ranges match
            if (newStudentSheetValues.length === newStudentRanges[i].getNumRows()) {
              allNewAdminSheetValues.push({ range: newAdminRanges[i], values: newAdminSheetFormulas });
              allNewStudentSheetValues.push({ range: newStudentRanges[i], values: newStudentSheetValues });
            } else {
              throw Error(`Mismatch in row count for ${sheetName} at level ${i + 1}`);
            }
          }
        }
        let testScore = testScores.find(score => score.test === sheetName);
        if (testScore) {
          newAdminSheet.getRange('G1:I1').setValues(testScore.scores);
        }
      }
    }

    allNewAdminSheetValues.forEach(item => item.range.setValues(item.values));
    allNewStudentSheetValues.forEach(item => item.range.setValues(item.values));

    // build timestamp column
    let newQbSheet = newAdminSs.getSheetByName('Question bank data');
    let timestampLookup = '=xlookup(A2, \'Student responses\'!$A$4:$A$10000, \'Student responses\'!$J$4:$J$10000,"")';
    let timestampStartCell = newQbSheet.getRange('K2');
    timestampStartCell.setValue(timestampLookup);
    let timestampRange = newQbSheet.getRange('K2:K10000');
    timestampStartCell.autoFill(timestampRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    let timestampValues = timestampRange.getValues();

    for (let row = 0; row < timestampValues.length; row ++) {
      const currentTime = new Date().getTime();
      Logger.log("Current time at row " + row + ": "  + currentTime);
      let ssRow = row + 2;
      if (timestampValues[row][0] === '') {
        timestampValues[row][0] = '=if(or(G' + ssRow + '="",I' + ssRow + '=""),"",if(K' + ssRow + ',K' + ssRow + ',if(I' + ssRow + '="","",now())))'
      }
    }
    timestampRange.setValues(timestampValues);
    timestampRange.setNumberFormat('MM/dd/yyy h:mm:ss');

    // build reviewed column
    newStudentData.getRange('L3').setValue('=importrange("' + oldAdminSsId + '", "Practice test data!E1:M")');
    const newPracticeDataSheet = newAdminSs.getSheetByName('Practice test data');
    const reviewedLookup = '=xlookup(E2, \'Student responses\'!$L$4:$L$10000, \'Student responses\'!$T$4:$T$10000,"")';
    const reviewedStartCell = newPracticeDataSheet.getRange('M2');
    reviewedStartCell.setValue(reviewedLookup);
    let reviewedRange = newPracticeDataSheet.getRange('M2:M10000');
    reviewedStartCell.autoFill(reviewedRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    reviewedRange.setValues(reviewedRange.getValues());
  }
  catch (err) {
    let htmlOutput = HtmlService.createHtmlOutput('<p>Error while processing data: ' + err.stack + '</p><p>Please copy this error and send to danny@openpathtutoring.com.</p>')
      .setWidth(400)
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error');
    Logger.log(err.stack);
  }

  // revert student ID and SS permissions
  newStudentData.getRange('A3').setValue(initialImportFunction);
  newStudentData.getRange('H3').setValue('');
  newStudentData.getRange('T3').setValue('');
  DriveApp.getFileById(oldAdminSsId).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  DriveApp.getFileById(newStudentSsId).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);

  let htmlOutput = HtmlService.createHtmlOutput('<a href="https://docs.google.com/spreadsheets/d/' + newStudentSsId + '" target="_blank" onclick="google.script.host.close()">Student answer sheet</a>')
    .setWidth(250)
    .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Data transfer complete');
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

function getLastFilledRow(sheet, col) {
  const lastRow = sheet.getLastRow();
  const allVals = sheet.getRange(1, col, lastRow).getValues();
  const lastFilledRow = lastRow - allVals.reverse().findIndex((c) => c[0] != '');

  return lastFilledRow;
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
