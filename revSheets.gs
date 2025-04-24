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
    let revDataSsId = revBackend.getRange('U3').getValue();
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
        var rowHeight = heightVals[row - 1][0]; // rowHeights including whitespace hard-coded in Rev sheet backend
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
