function getStudentData(adminFolderId, testType=null) {
  const adminFolder = DriveApp.getFolderById(adminFolderId);
  const adminSubfolders = adminFolder.getFolders();
  const studentName = adminFolder.getName();
  let satAdminSsId, satStudentSsId, actAdminSsId, actStudentSsId;

  // get testType from student folder name
  if (!testType) {
    while (adminSubfolders.hasNext()) {
      const adminSubfolder = adminSubfolders.next();
      const adminSubfolderName = adminSubfolder.getName();

      if (adminSubfolderName.includes(studentName)) {
        if (adminSubfolderName.includes('SAT')) {
          testType = 'sat'
        } //
        else if (adminSubfolderName.includes('ACT')) {
          testType = 'act'
        }
        break;
      }
    }
  }
  
  const adminFiles = adminFolder.getFiles();
  while (adminFiles.hasNext()) {
    const adminFile = adminFiles.next();
    const adminFileId = adminFile.getId();

    if (testType != 'act' && adminFile.getName().toLowerCase().includes('sat admin answer analysis')) {
      satAdminSsId = adminFileId;
      const satAdminSs = SpreadsheetApp.openById(satAdminSsId);
      satStudentSsId = satAdminSs.getSheetByName('Student responses').getRange('B1').getValue();

      if (testType === 'sat') {
        break;
      }
    }

    if (testType != 'sat' && adminFile.getName().toLowerCase().includes('act admin answer analysis')) {
      actAdminSsId = adminFileId;
      actStudentSsId = actAdminSs.getSheetByName('Student responses').getRange('B1').getValue();

      if (testType === 'act') {
        break;
      }
    }

    if ((satStudentSsId && actStudentSsId) || (testType === 'sat' && satStudentSsId) || (testType === 'act' && actStudentSsId)) {
      break;
    }
  }

  if (testType !== 'act') {
    if (!satAdminSsId) {
      satAdminSsId = findFirstIdBySubstring(adminFolderId, 'sat admin answer', 'file');
    }
    if (!satStudentSsId) {
      satStudentSsId = findFirstIdBySubstring(adminFolderId, 'sat student answer', 'file');
    }
  }
  
  if (testType !== 'sat') {
    if (!actAdminSsId) {
      actAdminSsId = findFirstIdBySubstring(adminFolderId, 'act admin answer', 'file');
    }
    if (!actStudentSsId) {
      actStudentSsId = findFirstIdBySubstring(adminFolderId, 'act student answer', 'file');
    }
  }

  const studentData = {
    name: studentName,
    adminFolderId: adminFolderId,
    satAdminSsId: satAdminSsId,
    satStudentSsId: satStudentSsId,
    actAdminSsId: actAdminSsId,
    actStudentSsId: actStudentSsId,
  }

  return studentData;
}


function getSatTestCodes() {
  const practiceTestDataSheet = SpreadsheetApp.openById('1KidSURXg5y-dQn_gm1HgzUDzaICfLVYameXpIPacyB0').getSheetByName('Practice test data');
  const lastFilledRow = getLastFilledRow(practiceTestDataSheet, 1);
  const testCodeCol = practiceTestDataSheet
    .getRange(2, 1, lastFilledRow - 1)
    .getValues()
    .map((row) => row[0]);
  const testCodes = testCodeCol.filter((x, i, a) => a.indexOf(x) == i);

  return testCodes;
}


function getActTestData(ssId, testCode) {
  // const completedEnglishCount = allData.filter((row) => row[0] === testCode && row[1] === 'English' && row[7] !== '').length;
  // const completedMathCount = allData.filter((row) => row[0] === testCode && row[1] === 'Math' && row[7] !== '').length;
  // const completedReadingCount = allData.filter((row) => row[0] === testCode && row[1] === 'Reading' && row[7] !== '').length;
  // const completedScienceCount = allData.filter((row) => row[0] === testCode && row[1] === 'Science' && row[7] !== '').length;

  // if (completedEnglishCount > 37 && completedMathCount > 30 && completedReadingCount > 20 && completedScienceCount > 20) {
  const ss = SpreadsheetApp.openById(ssId);
  let testSheet = ss.getSheetByName(testCode);

  if (testSheet) {
    const testHeaderValues = testSheet.getRange('A1:N3').getValues();
    const eScore = parseInt(testHeaderValues[2][1]) || 0;
    const mScore = parseInt(testHeaderValues[2][5]) || 0;
    const rScore = parseInt(testHeaderValues[2][9]) || 0;
    const sScore = parseInt(testHeaderValues[2][13]) || 0;
    const totalScore = Math.round(Number(testHeaderValues[0][5])) || '';
    const dateSubmitted = formatDateYYYYMMDD(testHeaderValues[0][7]);
    const isTestNew = dateSubmitted;

    return {
      test: testCode,
      eScore: eScore,
      mScore: mScore,
      rScore: rScore,
      sScore: sScore,
      total: totalScore,
      date: dateSubmitted,
      isNew: isTestNew,
    };
  }
}
// }


function getActTestCodes() {
  const dataSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('actMasterDataSsId')).getSheetByName('ACT Answers');
  const lastFilledRow = getLastFilledRow(dataSheet, 1);
  const testCodeCol = dataSheet
    .getRange(2, 1, lastFilledRow - 1)
    .getValues()
    .map((row) => row[0]);
  const testCodes = testCodeCol.filter((x, i, a) => a.indexOf(x) == i).sort().reverse();

  Logger.log(testCodes)

  return testCodes;
}


function addSatTestSheets(adminSsId) {
  const testCodes = getSatTestCodes();

  if (!adminSsId) {
    adminSsId = SpreadsheetApp.getActiveSpreadsheet().getId();
  }

  const adminSs = SpreadsheetApp.openById(adminSsId);
  const adminTemplateSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('satAdminTemplateSsId'));
  const adminTemplateSheet = adminTemplateSs.getSheetByName('SAT4');
  const studentSsId = adminSs.getSheetByName('Student responses').getRange('B1').getValue();
  const studentSs = SpreadsheetApp.openById(studentSsId);
  const studentTemplateSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('satStudentTemplateSsId'));
  const studentTemplateSheet = studentTemplateSs.getSheetByName('SAT4');

  const spreadsheets = [
    {
      'ss': studentSs,
      'templateSheet': studentTemplateSheet
    },
    {
      'ss': adminSs,
      'templateSheet': adminTemplateSheet
    }
  ]

  for (testCode of testCodes) {
    const testNumberPosition = testCode.indexOf('SAT') + 3;
    const testType = testCode.substring(0, testNumberPosition)
    const testNumber = testCode.substring(testNumberPosition);

    for (obj of spreadsheets) {
      const testSheet = obj.ss.getSheetByName(testCode);

      if (!testSheet) {
        Logger.log(`Adding ${testCode} sheet to ${obj.ss.getName()}`);
        const templateSheet = obj.templateSheet;
        const newSheet = templateSheet.copyTo(obj.ss).setName(testCode);
        const prevTestPostition = obj.ss.getSheetByName(testType + String(testNumber - 1)).getIndex();
        obj.ss.setActiveSheet(newSheet);
        obj.ss.moveActiveSheet(prevTestPostition + 1);

        newSheet.getRange('A2').setValue(testType);
        newSheet.getRange('A3').setValue(testNumber);

        const questionCodeFormulaR1C1 = '=iferror(let(worksheetNum,if(R[0]C[1]<>"",R[0]C[1]), qNum,right(worksheetNum,len(worksheetNum)-search(".",worksheetNum)),offset(R[0]C2,-1*qNum-2,0)&" "&worksheetNum),)';

        const colARange = newSheet.getRange('A5:A57');
        const colERange = newSheet.getRange('E5:E57');
        const colIRange = newSheet.getRange('I5:I57');

        colARange.setValue(questionCodeFormulaR1C1);
        colERange.setValue(questionCodeFormulaR1C1);
        colIRange.setValue(questionCodeFormulaR1C1);

        SpreadsheetApp.flush();

        colARange.copyTo(colARange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
        colERange.copyTo(colERange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
        colIRange.copyTo(colIRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
      }
    }

    if (!adminSs.getSheetByName(testCode + ' analysis')) {
      const newAnalysisSheet = adminTemplateSs.getSheetByName('SAT4 analysis').copyTo(adminSs).setName(`${testCode} analysis`);
      const prevAnalysisPostition = obj.ss.getSheetByName(testType + String(testNumber - 1) + ' analysis').getIndex();
      adminSs.setActiveSheet(newAnalysisSheet)
      adminSs.moveActiveSheet(prevAnalysisPostition + 1);

      newAnalysisSheet.getRange('A7').setValue(testType);
      newAnalysisSheet.getRange('A8').setValue(testNumber);
      Logger.log(`Added ${testCode} analysis sheet to ${adminSs.getName()}`)
    }
  }
}


function addActTestSheets(adminSsId, adminIndexAdjustment=1) {
  let adminSs;
  if (!adminSsId) {
    adminSs = SpreadsheetApp.getActiveSpreadsheet();
  }
  else {
    adminSs = SpreadsheetApp.openById(adminSsId);
  }

  const studentSsId = adminSs.getSheetByName('Student responses').getRange('B1').getValue();
  const studentSs = SpreadsheetApp.openById(studentSsId);
  const adminTemplateSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('actAdminTemplateSsId'));
  const adminTemplateSheet = adminTemplateSs.getSheetByName('202206');
  const studentTemplateSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('actStudentTemplateSsId'));
  const studentTemplateSheet = studentTemplateSs.getSheetByName('202206');

  const spreadsheets = [
    {
      'ss': studentSs,
      'templateSheet': studentTemplateSheet,
      'indexAdjustment': 1
    },
    {
      'ss': adminSs,
      'templateSheet': adminTemplateSheet,
      'indexAdjustment': adminIndexAdjustment     // 1-indexed + # of analysis sheets
    }
  ]

  const testCodes = getActTestCodes();
  for (obj of spreadsheets) {
    for (testCode of testCodes) {
      const testSheet = obj.ss.getSheetByName(testCode);

      if (!testSheet) {
        Logger.log(`Adding ${testCode} sheet to ${obj.ss.getName()}`);
        const templateSheet = obj.templateSheet;
        const newSheet = templateSheet.copyTo(obj.ss).setName(testCode);
        newSheet.getRange('B1').setValue(testCode);

        const testCodeIndex = testCodes.indexOf(testCode);

        obj.ss.setActiveSheet(newSheet);
        if (testCodeIndex > 0) {
          const prevTest = testCodes[testCodeIndex - 1];
          const prevTestPosition = obj.ss.getSheetByName(prevTest).getIndex() || 0;
          Logger.log(`Previous sheet: ${prevTest}, index ${prevTestPosition}`);
          obj.ss.moveActiveSheet(prevTestPosition + obj.indexAdjustment);
        }
        else {
          obj.ss.moveActiveSheet(1);
        }
      }
    }
  }
}

function sortActTestSheets(ssId, testCodes, isAdminSheet=true) {
  const ss = SpreadsheetApp.openById(ssId);
  if(!testCodes) {
    testCodes = getActTestCodes();
  }

  let indexAdjustment;
  if (isAdminSheet) {
    indexAdjustment = 3;   // 1-indexed + 2 analysis sheets
  }
  else {
    indexAdjustment = 1;
  }

  for (let i = 0; i < testCodes.length; i++) {
    const testCode = testCodes[i];
    const testSheet = ss.getSheetByName(testCode);
    ss.setActiveSheet(testSheet);
    ss.moveActiveSheet(i + indexAdjustment);
  }
}


function getActPageBreakRow(sheet) {
  const grandColData = sheet
    .getRange(1, 2, 111)
    .getValues()
    .map((row) => row[0]);
  const mathColData = sheet
    .getRange(1, 3, 111)
    .getValues()
    .map((row) => row[0]);

  const grandTotalIndex = grandColData.indexOf('Grand Total');
  if (0 < grandTotalIndex && grandTotalIndex < 80) {
    sheet.hideRows(grandTotalIndex + 2, 111);
    SpreadsheetApp.flush();
    return 80;
  }

  const mathTotalIndex = mathColData.indexOf('Math Total');
  if (0 < mathTotalIndex && mathTotalIndex < 80) {
    Logger.log(`Page break for analysis sheet at row ${mathTotalIndex + 1}`);
    return mathTotalIndex + 1;
  } //
  else {
    return 80;
  }
}


function getLastFilledRow(sheet, col) {
  const lastRow = sheet.getLastRow();
  const allVals = sheet.getRange(1, col, lastRow).getValues();
  const lastFilledRow = lastRow - allVals.reverse().findIndex((c) => c[0] != '');

  return lastFilledRow;
}

function getIdFromDriveUrl(url) {
  if (!url) {
    return null;
  }
  if (url.includes('/folders/')){
    id = url.split('/folders/')[1].split(/[/?]/)[0];
  }
  else if (url.includes('/d/')) {
    id = url.split('/d/')[1].split('/')[0];
  }
  else if (!url.includes('/')) {
    id = url;
  }
  else {
    throw Error('Unexpected URL format');
  }

  return id;
}


function isEmptyFolder(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  return !folder.getFiles().hasNext() && !folder.getFolders().hasNext();
}


function formatDateYYYYMMDD(dateStr) {
  const date = new Date(dateStr);
  if (!isNaN(date.getTime())) {
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    const yyyy = date.getFullYear();
    return `${yyyy}-${mm}-${dd}`;
  }
   
   return null;
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
    } //
    else {
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

/**
 * Search for the first file/folder whose name includes a given substring.
 *
 * @param {string} folderId - Root folder ID to search in.
 * @param {string} substring - Case-insensitive substring to match.
 * @param {"file"|"folder"|"both"} searchType - What to search for.
 * @return {string|null} The ID of the first match, or null if none found.
 */
function findFirstIdBySubstring(folderId, substring, searchType='file') {
  const folder = DriveApp.getFolderById(folderId);
  const lowerSubstring = substring.toLowerCase();

  // Check files if requested
  if (searchType === 'file' || searchType === 'both') {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().toLowerCase().includes(lowerSubstring)) {
        return file.getId();
      }
    }
  }

  // Check folders if requested
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();

    if (searchType === 'folder' || searchType === 'both') {
      if (subfolder.getName().toLowerCase().includes(lowerSubstring)) {
        return subfolder.getId();
      }
    }

    // Recurse into subfolders
    const foundId = findFirstIdBySubstring(subfolder.getId(), substring, searchType);
    if (foundId) {
      return foundId;
    }
  }

  return null; // nothing found
}

function getScoreReportFolderId(adminSsId, ssType='sat') {
  const adminSs = SpreadsheetApp.openById(adminSsId);
  const adminFolder = DriveApp.getFileById(adminSsId).getParents().next();
  const adminSubfolders = adminFolder.getFolders();
  let studentName, scoreReportFolder, scoreReportFolderId, studentFolder;

  if (ssType === 'sat') {
    const revBackendSheet = adminSs.getSheetByName('Rev sheet backend');
    if (revBackendSheet) {
      studentName = revBackendSheet.getRange('K2').getValue();
      scoreReportFolderId = revBackendSheet.getRange('U9').getValue();
    }
  } //
  else if (ssType === 'act') {
    const dataSheet = adminSs.getSheetByName('Data');
    scoreReportFolderId = dataSheet.getRange('W1').getValue();
  }

  if (!studentName) {
    studentName = adminFolder.getName();
  }

  if (scoreReportFolderId) {
    scoreReportFolder = DriveApp.getFolderById(scoreReportFolderId);
  }

  if (scoreReportFolder) {
    return scoreReportFolderId;
  } //
  else {
    while (adminSubfolders.hasNext()) {
      const adminSubfolder = adminSubfolders.next();

      if (adminSubfolder.getName().includes(studentName)) {
        studentFolder = adminSubfolder;
        break;
      }
    }

    if (studentFolder) {
      const studentSubfolders = studentFolder.getFolders();

      while (studentSubfolders.hasNext()) {
        const studentSubfolder = studentSubfolders.next();

        if (studentSubfolder.getName().toLowerCase().includes('score report')) {
          scoreReportFolderId = studentSubfolder.getId();
          break;
        } //
        else {
          scoreReportFolderId = studentFolder.createFolder('Score reports').getId();
          break;
        }
      }
    } //
    else {
      scoreReportFolderId = adminFolder.createFolder('Score reports').getId();
    }
  }
  
  if (ssType === 'sat') {
    revBackendSheet.getRange('T9:U9').setValues([['Score report folder ID:', scoreReportFolderId]]);
  }
  else if (ssType === 'act') {
    dataSheet.getRange('V1:W1').setValues([['Score report folder ID:', scoreReportFolderId]]);
  }
  
  Logger.log(scoreReportFolderId);
  return scoreReportFolderId;
}


function errorNotification(error, ssId) {
  
  // const htmlOutput = HtmlService.createHtmlOutput(`<p>We have been notified of the following error: ${error.message}</p><p>${error.stack}`)
  const htmlOutput = HtmlService.createHtmlOutput(`<p>Please copy-paste the following details and send to ${ADMIN_EMAIL}. Sorry about that!</p><p> ${error.message}</p><p>${error.stack}`)
    .setWidth(500) //optional
    .setHeight(300); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Error`);

  // const ss = SpreadsheetApp.openById(ssId);
  // const editorEmails = []
  // ss.getEditors().forEach(editor => editorEmails.push(editor.getEmail()));

  // const message = `
  //   <p>Error details: ${error.stack}</p>
  //   <p><a href="https://docs.google.com/spreadsheets/d/${ssId}" target="_blank">${ss.getName()}</a></p>
  //   <p>Editors: ${editorEmails}</p>
  // `

  // MailApp.sendEmail({
  //   to: ADMIN_EMAIL,
  //   subject: `Spreadsheet error: ${error.message}`,
  //   htmlBody: message
  // });

  Logger.log(error.message + '\n\n' + error.stack);
  throw new Error(error.message + '\n\n' + error.stack);
}