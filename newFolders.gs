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

function copyFolder(sourceFolderId = '1yqQx_qLsgqoNiDoKR9b63mLLeOiCoTwo', newFolderId = '1_qQNYnGPFAePo8UE5NfX72irNtZGF5kF', studentName = '_Aaron S', folderType = 'sat') {
  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const newFolder = DriveApp.getFolderById(newFolderId);
  const newFolderName = newFolder.getName();
  Logger.log(`${newFolderName} folder started`)

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

    Logger.log(testType);

    if (testType === 'SAT' && fileName.toLowerCase().includes('act') && fileName.toLowerCase().includes('answer analysis')) {
      newFile.setTrashed(true);
    } else if (testType === 'ACT' && fileName.toLowerCase().includes('sat') && fileName.toLowerCase().includes('answer analysis')) {
      newFile.setTrashed(true);
    }

    if (newFolderName.includes(folderType.toUpperCase()) && !newFolderName.includes(studentName)) {
      fileOperations.push({ file: newFile, action: 'move' });
    }
  }

  const newParentFolder = newFolder.getParents().next();

  // Perform file operations in batch
  fileOperations.forEach(op => {
    if (op.action === 'move') {
      op.file.moveTo(newParentFolder);
      Logger.log(file.getName() + ' moved to ' + newParentFolder.getId());
    } else if (op.action === 'trash') {
      op.folder.setTrashed(true);
      Logger.log(op.folder.getName() + ' trashed');
    }
  });

  if (isEmptyFolder(newFolder.getId()) && newFolderName.includes(folderType.toUpperCase()) && !newFolderName.includes(studentName)) {
    newFolder.setTrashed(true);
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
        let answerSheets = getSatTestCodes(ss);
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
        let continueScript = ui.alert('Rev data sheet with same student name already exists. All students must have unique names for rev sheets to work properly. Are you re-creating this folder for an existing student?', ui.ButtonSet.YES_NO);

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


