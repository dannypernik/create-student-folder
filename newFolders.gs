function NewSatFolder(sourceFolderId, parentFolderId) {
  const ids = getFolderIds(sourceFolderId, parentFolderId);
  sourceFolderId = ids.sourceFolderId;
  parentFolderId = ids.parentFolderId;

  let ui = SpreadsheetApp.getUi();
  let prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);

  if (prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  } //
  else {
    studentName = prompt.getResponseText();
  }

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();

  copyFolder(sourceFolderId, newFolderId, studentName, 'sat');

  const studentData = linkSheets(newFolderId, studentName, 'sat');
  studentData.folderId = newFolderId;

  var htmlOutput = HtmlService.createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + studentName + "'s folder</a>")
    .setWidth(250)
    .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'SAT folder created successfully');

  return studentData;
}

function NewActFolder(sourceFolderId, parentFolderId) {
  const ids = getFolderIds(sourceFolderId, parentFolderId);
  sourceFolderId = ids.sourceFolderId;
  parentFolderId = ids.parentFolderId;

  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);

  if (prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  } //
  else {
    studentName = prompt.getResponseText();
  }

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();

  copyFolder(sourceFolderId, newFolderId, studentName, 'act');

  const studentData = linkSheets(newFolderId, studentName, 'act');
  studentData.folderId = newFolderId;

  var htmlOutput = HtmlService.createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + studentName + "'s folder</a>")
    .setWidth(250)
    .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ACT folder created successfully');

  return studentData;
}

function NewTestPrepFolder(sourceFolderId, parentFolderId) {
  const ids = getFolderIds(sourceFolderId, parentFolderId);
  sourceFolderId = ids.sourceFolderId;
  parentFolderId = ids.parentFolderId;

  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);

  if (prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  } //
  else {
    studentName = prompt.getResponseText();
  }

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();

  copyFolder(sourceFolderId, newFolderId, studentName, 'all');

  const studentData = linkSheets(newFolderId, studentName, 'all');
  studentData.folderId = newFolderId;

  var htmlOutput = HtmlService.createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + studentName + "'s folder</a>")
    .setWidth(250)
    .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Test prep folder created successfully');

  return studentData;
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
  const sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const newFolder = DriveApp.getFolderById(newFolderId);
  const newFolderName = newFolder.getName();
  Logger.log(`${newFolderName} folder started`)

  var files = sourceFolder.getFiles();
  let testType

  if (folderType.toLowerCase() === 'sat') {
    testType = 'SAT';
  } //
  else if (folderType.toLowerCase() === 'act') {
    testType = 'ACT';
  } //
  else {
    testType = 'Test';
  }

  let fileOperations = [];

  while (files.hasNext()) {
    const file = files.next();
    const prefixFiles = ['Tutoring notes', 'ACT review sheet', 'SAT review sheet'];
    let filename = file.getName();

    if (prefixFiles.includes(filename)) {
      filename = studentName + ' ' + filename;
    }
    else if (filename.toLowerCase().includes('template')) {
      const rootName = filename.slice(0, filename.indexOf('-') + 2);
      filename = rootName + studentName;
    }

    const newFile = file.makeCopy(filename, newFolder);
    const newFilename = newFile.getName().toLowerCase();
    const newFileId = newFile.getId();

    if (newFilename.includes('tutoring notes')) {
      const ss = SpreadsheetApp.openById(newFileId);
      const sheet = ss.getSheetByName('Session notes');
      shId = sheet.getSheetId();
      sheet.getRange('G3').setValue('=hyperlink("https://docs.google.com/spreadsheets/d/' + newFileId + '/edit?gid=' + shId + '#gid=' + shId + '&range=B"&match(G2,B1:B,0)-1,"Go to latest session")');
    }

    if (newFilename.includes('admin notes')) {
      DocumentApp.openById(newFile.getId()).getBody().replaceText('StudentName', studentName);
    }

    if (testType === 'SAT' && filename.toLowerCase().includes('act') && filename.toLowerCase().includes('answer analysis')) {
      newFile.setTrashed(true);
    }
    else if (testType === 'ACT' && filename.toLowerCase().includes('sat') && filename.toLowerCase().includes('answer analysis')) {
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
    }
    else if (op.action === 'trash') {
      op.folder.setTrashed(true);
    }
  });

  if (isEmptyFolder(newFolder.getId()) && newFolderName.includes(folderType.toUpperCase()) && !newFolderName.includes(studentName)) {
    newFolder.setTrashed(true);
  }

  const sourceSubFolders = sourceFolder.getFolders();
  while (sourceSubFolders.hasNext()) {
    const sourceSubFolder = sourceSubFolders.next();
    const folderName = sourceSubFolder.getName();
    let targetFolder;

    if (folderName === 'Student') {
      targetFolder = newFolder.createFolder(studentName + ' ' + testType + ' prep');
    }
    else if (newFolderName.includes(folderType.toUpperCase()) && newFolderName !== studentName + ' ' + testType + ' prep') {
      targetFolder = newFolder.getParents().next().createFolder(folderName);
    }
    else {
      targetFolder = newFolder.createFolder(folderName);
    }

    targetFolderName = targetFolder.getName();

    if (targetFolderName.includes('ACT') && folderType.toLowerCase() === 'sat') {
      targetFolder.setTrashed(true);
    }
    else if (targetFolderName.includes('SAT') && folderType.toLowerCase() === 'act') {
      targetFolder.setTrashed(true);
    }
    else {
      copyFolder(sourceSubFolder.getId(), targetFolder.getId(), studentName, folderType);
    }
  }
}

function linkSheets(folderId, studentName='', prepType='all') {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const subFolders = folder.getFolders();
  const satFiles = [];
  const actFiles = [];

  while (files.hasNext()) {
    const file = files.next();
    const filename = file.getName();
    const fileId = file.getId();

    if (filename.includes('SAT') && prepType !== 'act') {
      satFiles.push({ filename, fileId });
    }
    else if (filename.includes('ACT') && prepType !== 'sat') {
      actFiles.push({ filename, fileId });
    }
  }

  satFiles.forEach(({ filename, fileId }) => {
    driveFile = DriveApp.getFileById(fileId);
    driveFile.addEditor(SERVICE_ACCOUNT_EMAIL);
    if (filename.toLowerCase().includes('student answer sheet')) {
      satSheetIds.student = fileId;
      driveFile.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    }
    else if (filename.toLowerCase().includes('answer analysis')) {
      satSheetIds.admin = fileId;
      const ss = SpreadsheetApp.openById(fileId);

      ss.getSheets().forEach(s => {
        const sName = s.getName();
        const answerSheets = getSatTestCodes(ss);
        answerSheets.push('Reading & Writing', 'Math', 'SLT Uniques');

        if (answerSheets.includes(sName)) {
          s.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
        }
      });

      const revBackend = ss.getSheetByName('Rev sheet backend');
      revBackend.getRange('K2').setValue(studentName);
    }
  });

  actFiles.forEach(({ filename, fileId }) => {
    driveFile = DriveApp.getFileById(fileId);
    driveFile.addEditor(SERVICE_ACCOUNT_EMAIL);
    if (filename.toLowerCase().includes('student answer sheet')) {
      actSheetIds.student = fileId;
      driveFile.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    } //
    else if (filename.toLowerCase().includes('answer analysis')) {
      actSheetIds.admin = fileId;
      const ss = SpreadsheetApp.openById(fileId);
      ss.getSheetByName('Student responses').getRange('G1').setValue(studentName);
    }
  });

  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    linkSheets(subFolder.getId(), studentName, prepType);
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
    const satAdminSheet = SpreadsheetApp.openById(satSheetIds.admin);
    const satStudentSheet = SpreadsheetApp.openById(satSheetIds.student);
    satAdminSheet.getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);

    const revDataId = satAdminSheet.getSheetByName('Rev sheet backend').getRange('U3').getValue();
    const revDataSs = SpreadsheetApp.openById(revDataId);

    let studentRevDataSheet = revDataSs.getSheetByName(studentName);
    if (!studentRevDataSheet) {
      try {
        studentRevDataSheet = revDataSs.getSheetByName('Template').copyTo(revDataSs).setName(studentName);
      } catch (err) {
        const ui = SpreadsheetApp.getUi();
        const continueScript = ui.alert('Rev data sheet with same student name already exists. All students must have unique names for rev sheets to work properly. Are you re-creating this folder for an existing student?', ui.ButtonSet.YES_NO);

        if (continueScript === ui.Button.NO) {
          const htmlOutput = HtmlService.createHtmlOutput('<p>Please use a unique name for the new student or delete/rename the "'+ studentName + '" sheet from your <a href="https://docs.google.com/spreadsheets/d/' + revDataId + '">Rev sheet data</a></p>')
            .setWidth(400);
          SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Duplicate student name');
          return;
        }
      }
    }

    const studentQBSheet = satStudentSheet.getSheetByName('Question bank data');
    studentQBSheet.getRange('U2').setValue(studentName);
    studentQBSheet.getRange('U4').setValue(satSheetIds.admin);


    satAdminSheet.getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);
  }

  if (actSheetIds.student && actSheetIds.admin) {
    const actAdminSheet = SpreadsheetApp.openById(actSheetIds.admin);
    actAdminSheet.getSheetByName('Student responses').getRange('B1').setValue(actSheetIds.student);
  }

  const studentData = {
    name: studentName,
    satAdminSsId: satSheetIds.admin,
    satStudentSsId: satSheetIds.student,
    actAdminSsId: actSheetIds.admin,
    actStudentSsId: actSheetIds.student,
  }

  return studentData;
}


