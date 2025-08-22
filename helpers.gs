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


function getLastFilledRow(sheet, col) {
  const lastRow = sheet.getLastRow();
  const allVals = sheet.getRange(1, col, lastRow).getValues();
  const lastFilledRow = lastRow - allVals.reverse().findIndex((c) => c[0] != '');

  return lastFilledRow;
}

function getIdFromDriveUrl(url) {
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
