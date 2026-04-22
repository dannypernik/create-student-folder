function NewSatFolder(sourceFolderId, parentFolderId) {
  const ids = getFolderIds(sourceFolderId, parentFolderId);
  sourceFolderId = ids.sourceFolderId;
  parentFolderId = ids.parentFolderId;

  let ui = SpreadsheetApp.getUi();
  let prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);
  let studentName;

  if (prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  } //
  else {
    studentName = prompt.getResponseText();
  }

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();

  const studentData = copyFolder(sourceFolderId, newFolderId, studentName, 'sat');
  studentData.folderId = newFolderId;
  studentData.name = studentName;

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

  const studentData = copyFolder(sourceFolderId, newFolderId, studentName, 'act');
  studentData.folderId = newFolderId;
  studentData.name = studentName;

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

  const studentData = copyFolder(sourceFolderId, newFolderId, studentName, 'all');
  studentData.folderId = newFolderId;
  studentData.name = studentName;

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

function copyFolder(sourceFolderId = '1yqQx_qLsgqoNiDoKR9b63mLLeOiCoTwo', newFolderId = '1_qQNYnGPFAePo8UE5NfX72irNtZGF5kF', studentName = '_Aaron S', folderType = 'sat', studentData={}) {
  let sourceFolder, newFolder, newFolderName;
  function withBackoff(fn, maxRetries = 5, baseDelay = 500) {
    let attempt = 0;
    while (true) {
      try {
        return fn();
      } catch (e) {
        if (attempt >= maxRetries) throw e;
        let delay = baseDelay * Math.pow(2, attempt) + Math.floor(Math.random() * 200);
        Logger.log(`Transient error: ${e}. Retrying in ${delay}ms (attempt ${attempt + 1}/${maxRetries})`);
        Utilities.sleep(delay);
        attempt++;
      }
    }
  }
  try {
    Logger.log(`Getting source folder: ${sourceFolderId}`);
    sourceFolder = withBackoff(() => DriveApp.getFolderById(sourceFolderId));
    Logger.log(`Getting new folder: ${newFolderId}`);
    newFolder = withBackoff(() => DriveApp.getFolderById(newFolderId));
    newFolderName = withBackoff(() => newFolder.getName());
    Logger.log(`${newFolderName} folder started`);
  } catch (e) {
    Logger.log(`Error getting folders. sourceFolderId: ${sourceFolderId}, newFolderId: ${newFolderId}, error: ${e}`);
    errorNotification(e, newFolderId || sourceFolderId);
    return studentData;
  }

  try {
    var files = withBackoff(() => sourceFolder.getFiles());
    let testType;
    if (folderType.toLowerCase() === 'sat') {
      testType = 'SAT';
    } else if (folderType.toLowerCase() === 'act') {
      testType = 'ACT';
    } else {
      testType = 'Test';
    }
    let fileOperations = [];
    while (files.hasNext()) {
      let file, filename, newFile, newFilename, newFileId;
      try {
        file = withBackoff(() => files.next());
        const prefixFiles = ['Tutoring notes', 'ACT review sheet', 'SAT review sheet'];
        filename = withBackoff(() => file.getName());
        if (prefixFiles.includes(filename)) {
          filename = studentName + ' ' + filename;
        } else if (filename.toLowerCase().includes('template')) {
          const rootName = filename.slice(0, filename.indexOf('-') + 2);
          filename = rootName + studentName;
        }
        Logger.log(`Making copy of file: ${filename} in folder: ${newFolderName}`);
        newFile = withBackoff(() => file.makeCopy(filename, newFolder));
        newFilename = withBackoff(() => newFile.getName().toLowerCase());
        newFileId = withBackoff(() => newFile.getId());
      } catch (e) {
        Logger.log(`Error copying file. filename: ${filename}, error: ${e}`);
        continue;
      }
      try {
        if (newFilename && newFilename.includes('tutoring notes')) {
          const ss = withBackoff(() => SpreadsheetApp.openById(newFileId));
          const sheet = withBackoff(() => ss.getSheetByName('Session notes'));
          shId = withBackoff(() => sheet.getSheetId());
          withBackoff(() => sheet.getRange('G3').setValue('=hyperlink("https://docs.google.com/spreadsheets/d/' + newFileId + '/edit?gid=' + shId + '#gid=' + shId + '&range=B"&match(G2,B1:B,0)-1,"Go to latest session")'));
        }
      } catch (e) {
        Logger.log(`Error updating tutoring notes hyperlink. fileId: ${newFileId}, error: ${e}`);
      }
      // ...existing code for updating studentData and linking files...
      try {
        if (filename.toLowerCase().includes('answer analysis')) {
          if (filename.includes('SAT') && testType !== 'ACT') {
            studentData.satAdminSsId = newFileId;
          } else if (filename.includes('ACT') && testType !== 'SAT') {
            studentData.actAdminSsId = newFileId;
          }
        } else if (filename.toLowerCase().includes('student answer sheet')) {
          if (filename.includes('SAT') && testType !== 'ACT') {
            studentData.satStudentSsId = newFileId;
          } else if (filename.includes('ACT') && testType !== 'SAT') {
            studentData.actStudentSsId = newFileId;
          }
        } else if (filename.toLowerCase().includes('homework')) {
          studentData.homeworkSsId = newFileId;
        }
        if (testType !== 'ACT' && !studentData.isSatLinked && studentData.satAdminSsId && studentData.satStudentSsId) {
          linkSatFiles(studentData.satAdminSsId, studentData.satStudentSsId, studentName);
          studentData.isSatLinked = true;
          Logger.log('SAT files linked');
        }
        if (testType !== 'SAT' && !studentData.isActLinked && studentData.actAdminSsId && studentData.actStudentSsId) {
          linkActFiles(studentData.actAdminSsId, studentData.actStudentSsId, studentName);
          studentData.isActLinked = true;
          Logger.log('ACT files linked');
        }
        if (testType === 'SAT' && filename.includes('ACT') && filename.toLowerCase().includes('answer analysis')) {
          newFile.setTrashed(true);
        } else if (testType === 'ACT' && filename.includes('SAT') && filename.toLowerCase().includes('answer analysis')) {
          newFile.setTrashed(true);
        }
        if (newFolderName.includes(folderType.toUpperCase()) && !newFolderName.includes(studentName)) {
          fileOperations.push({ file: newFile, action: 'move' });
        }
      } catch (e) {
        Logger.log(`Error processing file metadata. filename: ${filename}, error: ${e}`);
      }
    }
    let newParentFolder;
    try {
      newParentFolder = withBackoff(() => newFolder.getParents().next());
    } catch (e) {
      Logger.log(`Error getting parent folder for: ${newFolderName}, error: ${e}`);
      newParentFolder = null;
    }
    fileOperations.forEach(op => {
      try {
        if (op.action === 'move' && newParentFolder) {
          withBackoff(() => op.file.moveTo(newParentFolder));
        } else if (op.action === 'trash') {
          withBackoff(() => op.folder.setTrashed(true));
        }
      } catch (e) {
        Logger.log(`Error in file operation (${op.action}). error: ${e}`);
      }
    });
    if (withBackoff(() => isEmptyFolder(newFolder.getId())) && newFolderName.includes(folderType.toUpperCase()) && !newFolderName.includes(studentName)) {
      try {
        withBackoff(() => newFolder.setTrashed(true));
      } catch (e) {
        Logger.log(`Error trashing empty folder: ${newFolderName}, error: ${e}`);
      }
    }
    const sourceSubFolders = withBackoff(() => sourceFolder.getFolders());
    while (sourceSubFolders.hasNext()) {
      let sourceSubFolder, folderName, targetFolder, targetFolderName;
      try {
        sourceSubFolder = withBackoff(() => sourceSubFolders.next());
        folderName = withBackoff(() => sourceSubFolder.getName());
        if (folderName === 'Student') {
          targetFolder = withBackoff(() => newFolder.createFolder(studentName + ' ' + testType + ' prep'));
          studentData.studentFolderId = withBackoff(() => targetFolder.getId());
        } else if (newFolderName.includes(folderType.toUpperCase()) && newFolderName !== studentName + ' ' + testType + ' prep') {
          targetFolder = withBackoff(() => newFolder.getParents().next().createFolder(folderName));
        } else {
          targetFolder = withBackoff(() => newFolder.createFolder(folderName));
        }
        targetFolderName = withBackoff(() => targetFolder.getName());
        if (targetFolderName.includes('ACT') && folderType.toLowerCase() === 'sat') {
          withBackoff(() => targetFolder.setTrashed(true));
        } else if (targetFolderName.includes('SAT') && folderType.toLowerCase() === 'act') {
          withBackoff(() => targetFolder.setTrashed(true));
        } else {
          copyFolder(sourceSubFolder.getId(), targetFolder.getId(), studentName, folderType, studentData);
        }
      } catch (e) {
        Logger.log(`Error copying subfolder. folderName: ${folderName}, error: ${e}`);
      }
    }
  } catch (err) {
    Logger.log(`General error in copyFolder: ${err}`);
    errorNotification(err, newFolder.getId());
  }
  return studentData;
}

function linkSatFiles(satAdminSsId, satStudentSsId, studentName='') {
  const satAdminFile = DriveApp.getFileById(satAdminSsId);
  const satStudentFile = DriveApp.getFileById(satStudentSsId);
  satAdminFile.addEditor(SERVICE_ACCOUNT_EMAIL);
  satStudentFile.addEditor(SERVICE_ACCOUNT_EMAIL);

  try {
    satStudentFile.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  }
  catch(e) {
    Logger.log('Unable to set sharing of SAT student file.');
  }

  const satAdminSs = SpreadsheetApp.openById(satAdminSsId);
  const revBackend = satAdminSs.getSheetByName('Rev sheet backend');
  revBackend.getRange('K2').setValue(studentName);
  satAdminSs.getSheetByName('Student responses').getRange('B1').setValue(satStudentSsId);
  satAdminSs.getSheets().forEach(s => {
    const sName = s.getName();
    const answerSheets = getSatTestCodes(satAdminSs);
    answerSheets.push('Reading & Writing', 'Math', 'SLT Uniques');

    if (answerSheets.includes(sName)) {
      s.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    }
  });

  const revDataId = satAdminSs.getSheetByName('Rev sheet backend').getRange('U3').getValue();
  const revDataSs = SpreadsheetApp.openById(revDataId);

  let studentRevDataSheet = revDataSs.getSheetByName(studentName);
  if (!studentRevDataSheet) {
    try {
      studentRevDataSheet = revDataSs.getSheetByName('Template').copyTo(revDataSs).setName(studentName);
    } //
    catch (err) {
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

  const satStudentSs = SpreadsheetApp.openById(satStudentSsId);
  const studentQBSheet = satStudentSs.getSheetByName('Question bank data');
  studentQBSheet.getRange('U2').setValue(studentName);
  studentQBSheet.getRange('U4').setValue(satAdminSsId);

  const satSsIds = {
    satAdminSsId: satAdminSsId,
    satStudentSsId: satStudentSsId,
  }

  return satSsIds;
}

function linkActFiles(actAdminSsId, actStudentSsId, studentName='') {
  actAdminFile = DriveApp.getFileById(actAdminSsId);
  actAdminFile.addEditor(SERVICE_ACCOUNT_EMAIL);
  actStudentFile = DriveApp.getFileById(actStudentSsId);
  actStudentFile.addEditor(SERVICE_ACCOUNT_EMAIL);
  try {
    actStudentFile.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  }
  catch(e) {
    Logger.log('Unable to set sharing of ACT student file.');
  }

  const actAdminSs = SpreadsheetApp.openById(actAdminSsId);
  actAdminSs.getSheetByName('Student responses').getRange('G1').setValue(studentName);
  actAdminSs.getSheetByName('Student responses').getRange('B1').setValue(actStudentSsId);

  const actSsIds = {
    actAdminSsId: actAdminSsId,
    actStudentSsId: actStudentSsId,
  }

  return actSsIds;
}