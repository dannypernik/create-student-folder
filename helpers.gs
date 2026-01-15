function getAllStudentData(
  client={
    index: null,
    name: null,
    studentsFolderId: null,
    studentsData: null,
    studentsDataCell: null
  },
  checkAllKeys=false)
  {
  const index = client.index || 0;

  Logger.log(index + '. ' + client.name + ' started');

  const studentFolders = DriveApp.getFolderById(client.studentsFolderId).getFolders();
  const studentFolderIds = [];

  const studentFolderList = sortFoldersByName(studentFolders);
  for (const studentFolder of studentFolderList) {
    const studentFolderId = studentFolder.getId();
    const studentName = studentFolder.getName();
    studentFolderIds.push(studentFolderId);
    let studentObj = client.studentsData.find(obj => obj.folderId === studentFolderId);

    if (studentName.includes('Ξ')) {
      continue;
    }

    if ((!studentObj) || (checkAllKeys && !studentObj.updateComplete)) {
      if (!studentObj) {
        studentObj = { folderId: studentFolderId };
      }
      studentObj = getStudentData(studentObj);
      Logger.log(`All data checked for ${studentName}`);
    } //

    if (!studentObj.testType) {
      studentObj.testType = getStudentTestType(studentObj.folderId, studentName);
      Logger.log(`${studentName} testType = ${studentObj.testType}`);
    }

    if ((!studentObj.satAdminSsId && studentObj.testType !== 'act') || (!studentObj.actAdminSsId && studentObj.testType !== 'sat')) {
      const ssIdObj = getSsIds(studentObj.folderId, studentObj.testType);
      studentObj.satAdminSsId = ssIdObj.satAdminSsId;
      studentObj.satStudentSsId = ssIdObj.satStudentSsId;
      studentObj.actAdminSsId = ssIdObj.actAdminSsId;
      studentObj.actStudentSsId = ssIdObj.actStudentSsId;
      studentObj.homeworkSsId = ssIdObj.homeworkSsId;
      if (studentObj.homeworkSsId) {
        Logger.log(`${studentName} ssIds updated with homeworkSsId = ${studentObj.homeworkSsId}`);
      } //
      else if ((studentObj.satAdminSsId && studentObj.testType !== 'act') || (studentObj.actAdminSsId && studentObj.testType !== 'sat')) {
        Logger.log(`${studentName} ssIds updated`);
      }
    }

    if (studentObj.testType === 'sat') {
      studentObj.actAdminSsId = null;
      studentObj.actStudentSsId = null;
    } //
    else if (studentObj.testType === 'act') {
      studentObj.satAdminSsId = null;
      studentObj.satStudentSsId = null;
    }

    studentObj.name = studentName;
    studentObj.updateComplete = true
    client.studentsData = updateStudentsJSON(studentObj, client.studentsData);

    if (client.studentsDataCell) {
      client.studentsDataCell.setValue(JSON.stringify(client.studentsData));
    }
  }

  return client.studentsData;
}

function getStudentData(studentData={}) {
  const studentFolder = DriveApp.getFolderById(studentData.folderId);
  studentData.name = studentFolder.getName();

  if (!studentData.testType) {
    studentData.testType = getStudentTestType(studentData.folderId, studentData.name);
  }

  const ssIdObj = getSsIds(studentData.folderId, studentData.testType);
  studentData.satAdminSsId = ssIdObj.satAdminSsId;
  studentData.satStudentSsId = ssIdObj.satStudentSsId;
  studentData.actAdminSsId = ssIdObj.actAdminSsId;
  studentData.actStudentSsId = ssIdObj.actStudentSsId;
  studentData.homeworkSsId = ssIdObj.homeworkSsId;

  return studentData;
}

function getStudentTestType(studentFolderId, studentName) {
  let testType = 'all';
  const studentFolder = DriveApp.getFolderById(studentFolderId);
  const subfolders = studentFolder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const subfolderName = subfolder.getName();
    if (subfolderName.includes(studentName)) {
      if (subfolderName.includes('SAT')) testType = 'sat';
      else if (subfolderName.includes('ACT')) testType = 'act';
      break;
    }
  }

  return testType;
}

function getSsIds(studentFolderId, testType) {
  const studentFolder = DriveApp.getFolderById(studentFolderId);
  const files = studentFolder.getFiles();
  let satAdminSsId, satStudentSsId, actAdminSsId, actStudentSsId, homeworkSsId;
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName().toLowerCase();
    const fileId = file.getId();

    if (fileName.includes('sat admin answer')) {
      satAdminSsId = fileId;
      const satAdminSs = SpreadsheetApp.openById(satAdminSsId);
      satStudentSsId = satAdminSs.getSheetByName('Student responses').getRange('B1').getValue();
      const revBackendSheet = satAdminSs.getSheetByName('Rev sheet backend');

      if (revBackendSheet) {
        homeworkSsId = revBackendSheet.getRange('U8').getValue();
      }

      if (testType === 'sat') break;
    }

    if (fileName.includes('act admin answer')) {
      actAdminSsId = fileId;
      const adminSs = SpreadsheetApp.openById(actAdminSsId);
      actStudentSsId = adminSs.getSheetByName('Student responses').getRange('B1').getValue();
      homeworkSsId = adminSs.getSheetByName('Data').getRange('U2').getValue();

      if (testType === 'act') break;
    }

    if (homeworkSsId) {
      const homeworkSs = DriveApp.getFileById(homeworkSsId);
      const editors = homeworkSs.getEditors();
      editorEmails = editors.map(editor => editor.getEmail());

      if (!editorEmails.includes(ADMIN_EMAIL)){
        Logger.log('Added editor');
        homeworkSs.addEditor(ADMIN_EMAIL);
      }
    }

    if ((satStudentSsId && actStudentSsId) || (testType === 'sat' && satStudentSsId) || (testType === 'act' && actStudentSsId)) {
      break;
    }
  }

  // ---------- Fallback search if IDs not found ----------
  if (testType !== 'act') {
    if (!satAdminSsId) satAdminSsId = findFirstIdBySubstring(studentFolderId, 'sat admin answer', 'file');
    if (!satStudentSsId) satStudentSsId = findFirstIdBySubstring(studentFolderId, 'sat student answer', 'file');
  }
  if (testType !== 'sat') {
    if (!actAdminSsId) actAdminSsId = findFirstIdBySubstring(studentFolderId, 'act admin answer', 'file');
    if (!actStudentSsId) actStudentSsId = findFirstIdBySubstring(studentFolderId, 'act student answer', 'file');
  }

  ssIds = {
    satAdminSsId: satAdminSsId,
    satStudentSsId: satStudentSsId,
    actAdminSsId: actAdminSsId,
    actStudentSsId: actStudentSsId,
    homeworkSsId: homeworkSsId
  }

  return ssIds;
}

function updateStudentsJSON(studentData, studentsJSON) {
  let existing = studentsJSON.find(obj => obj.folderId === studentData.folderId);

  if (existing) {
    let changed = false;
    for (let key in studentData) {
      if (studentData[key] && studentData[key] !== existing[key]) {
        existing[key] = studentData[key];
        Logger.log(`${studentData.name} ${key} updated`);
        changed = true;
      }
    }
    if (!changed) {
      Logger.log(`${studentData.name} unchanged`);
    }
  } //
  else {
    studentsJSON.push(studentData);
    Logger.log(`Added ${studentData.name}`);
  }

  return studentsJSON;
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

function getActTestCodes(dataSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('actMasterDataSsId')).getSheetByName('ACT Answers')) {
  // Only column A, from row 2 down to last row
  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return [];

  const colA = dataSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // [[A1],[A2],...]
  const set = new Set();

  for (let i = 0; i < colA.length; i++) {
    const v = colA[i][0];
    if (v !== '' && v != null) set.add(String(v));
  }

  const testCodes = Array.from(set).sort().reverse();
  Logger.log(testCodes);
  return testCodes;
}


// function updateActTestSheets() {
//   const startTime = new Date().getTime(); // Record the start time
//   const maxDuration = 5.5 * 60 * 1000; // 5 minutes and 30 seconds in milliseconds
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const ui = SpreadsheetApp.getUi();
//   const button = ui.alert('Use official scoring?', ui.ButtonSet.YES_NO_CANCEL);
//   let templateSs;

//   if (button === ui.Button.YES) {
//     templateSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('actTemplateSsId'));
//     ui.alert('Template SS has not been set up to enable official scoring.');
//     return;
//   } //
//   else if (button === ui.Button.NO) {
//     templateSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('actTemplateSsId'));
//   } //
//   else {
//     return;
//   }

//   const sheets = ss.getSheets();
//   Logger.log(`Starting test codes`);
//   const testCodes = getActTestCodes();
//   Logger.log(`Retreived test codes`);

//   const legacyTemplateSheet = templateSs.getSheetByName('Admin legacy');
//   const enhancedTemplateSheet = templateSs.getSheetByName('Admin enhanced adjusted');
//   const newLegacySheet = legacyTemplateSheet.copyTo(ss);
//   const newEnhancedSheet = enhancedTemplateSheet.copyTo(ss);

//   Logger.log('Copied test sheets');

//   const templateLegacyHeader = newLegacySheet.getRange('A1:P4');
//   const templateLegacyAnswers = newLegacySheet.getRange('A5:P80');
//   const templateEnhancedHeader = newEnhancedSheet.getRange('A1:P4');
//   const templateEnhancedAnswers = newEnhancedSheet.getRange('A5:P55');
//   const styleSheet = ss.getSheetByName('201904');
//   const headerBgColor = styleSheet.getRange('A1').getBackground();
//   // const headerFontColor = styleSheet.getFontColorObject().asRgbColor().asHexString();

//   Logger.log(`Got ranges and colors`);

//   // if (headerFontColor.toLowerCase() !== '#ffffff') {
//   //   alert('Implement header font color');
//   //   errorNotification('Implement header font color', ss.getId());
//   // }

//   try {
//     for (let sh of sheets) {
//       const currentTime = new Date().getTime();
//       if (currentTime - startTime > maxDuration) {
//         Logger.log("Exiting loop after 5 minutes and 30 seconds.");
//         throw new Error("Process exceeded maximum duration of 5 minutes and 30 seconds. Cleaning up.");
//       }

//       const sheetName = sh.getName();
//       if (testCodes.includes(sheetName)) {
//         Logger.log('Get ranges');
//         const mergeRanges = sh.getRange('A1:N1').getMergedRanges();
//         const headerRange = sh.getRange('A1');
//         const bodyRange = sh.getRange('A5');
//         const compositeCell = sh.getRange('E1');
//         const infoCell = sh.getRange('G1');
//         const testCodeCell = sh.getRange('B1');
//         const enhancedCheckCells = sh.getRangeList(['C3', 'G3', 'K3'])

//         Logger.log('Start changes');
//         mergeRanges.forEach(range => range.breakApart());
//         if (sheetName > '202502') {
//           // Enhanced
//           templateEnhancedHeader.copyTo(headerRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
//           templateEnhancedAnswers.copyTo(bodyRange);
//           compositeCell.setHorizontalAlignment('right');
//           infoCell.setHorizontalAlignment('left');
//         } // Legacy
//         else {
//           templateLegacyHeader.copyTo(headerRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
//           enhancedCheckCells.setFontColor(headerBgColor);
//           templateLegacyAnswers.copyTo(bodyRange);
//           compositeCell.setHorizontalAlignment('right');
//           infoCell.setHorizontalAlignment('left');
//         }
//         // sh.getRange('A1:P4').setBackground(headerBgColor).setFontColor(headerFontColor).setBorder(true,true,true,true,true,true,headerBgColor,SpreadsheetApp.BorderStyle.SOLID);
//         testCodeCell.setValue(sheetName);

//         // setScoreColor(sh);

//         Logger.log(`Updated ${sheetName}`);
//       }
//     }
//   }
//   catch (err) {
//     Logger.log(err)
//   }
//   finally {
//     ss.deleteSheet(newLegacySheet);
//     ss.deleteSheet(newEnhancedSheet);
//     Logger.log('Removed template sheets');
//   }
// }

function updateActTestSheets() {
  const startTime = Date.now();
  const maxDuration = 5.5 * 60 * 1000; // 5m30s
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const button = ui.alert('Use official scoring?', ui.ButtonSet.YES_NO_CANCEL);
  const templateSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('actTemplateSsId'));

  let enhancedTemplateSheet;

  if (button === ui.Button.YES) {
    ui.alert('Template SS has not been set up to enable official scoring.');
    enhancedTemplateSheet = templateSs.getSheetByName('Admin enhanced official');
    return;
  } else if (button === ui.Button.NO) {
    enhancedTemplateSheet = templateSs.getSheetByName('Admin enhanced adjusted');
  } else {
    return;
  }

  // Template sheets (no need to copy them into ss anymore)
  const legacyTemplateSheet = templateSs.getSheetByName('Admin legacy');

  if (!legacyTemplateSheet || !enhancedTemplateSheet) {
    throw new Error('Missing template sheet(s): Admin legacy / Admin enhanced adjusted');
  }

  // Header ranges in the template sheets
  const LEGACY_HEADER_A1_P4 = legacyTemplateSheet.getRange('A1:P4');
  const ENH_HEADER_A1_P4 = enhancedTemplateSheet.getRange('A1:P4');

  // Capture BOTH formulas and values for headers
  // (We write formulas first, then values to preserve constants/labels in the template)
  const legacyHeaderFormulas = LEGACY_HEADER_A1_P4.getFormulasR1C1();
  const legacyHeaderValues = LEGACY_HEADER_A1_P4.getValues();

  const enhHeaderFormulas = ENH_HEADER_A1_P4.getFormulasR1C1();
  const enhHeaderValues = ENH_HEADER_A1_P4.getValues();

  const dataSheet = ss.getSheetByName('Data');
  const testFilterCell = dataSheet.getRange('M2');
  testFilterCell.setFormula('=unique(A2:A)');
  // Style reference
  const styleSheet = ss.getSheetByName('201904');
  if (!styleSheet) throw new Error('Missing style sheet: 201904');
  const headerBgColor = styleSheet.getRange('A1').getBackground();

  const testCodes = getActTestCodes();

  try {
    // Iterate only the test sheets by name, instead of scanning ss.getSheets()
    for (const sheetName of testCodes) {
      if (Date.now() - startTime > maxDuration) {
        Logger.log('Exiting loop after 5 minutes and 30 seconds.');
        throw new Error('Process exceeded maximum duration of 5 minutes and 30 seconds.');
      }

      const testDataRow = testCodes.indexOf(sheetName) + 2;
      Logger.log(testDataRow)
      const testDataCell = dataSheet.getRange(testDataRow, 14);
      if (testDataCell.getValue() === 'done') {
        Logger.log(`Skipping ${sheetName}`);
        continue;
      }

      const sh = ss.getSheetByName(sheetName);
      if (!sh) {
        Logger.log(`Sheet not found, skipping: ${sheetName}`);
        continue;
      }

      // Unmerge header row (only if there are merges)
      const mergeRanges = sh.getRange('A1:N1').getMergedRanges();
      if (mergeRanges.length) {
        mergeRanges.forEach(r => r.breakApart());
      }

      // Common header cells
      const headerRange = sh.getRange('A1'); // top-left anchor
      const compositeCell = sh.getRange('E1');
      const infoCell = sh.getRange('G1');
      const testCodeCell = sh.getRange('B1');

      const enhancedCheckCells = sh.getRangeList(['C3', 'G3', 'K3']);

      if (sheetName > '202502') {
        // Enhanced: write header (formulas + values)
        // Break merges across header once (prevents Service error)
        // sh.getRange('A1:P4').breakApart();

        // Row 1
        sh.getRange('A1:P1')
          .setFormulasR1C1(enhHeaderFormulas.slice(0, 1))
          .setValues(enhHeaderValues.slice(0, 1));

        // Rows 3–4 (skip row 2)
        sh.getRange('A3:P4')
          .setFormulasR1C1(enhHeaderFormulas.slice(2, 4))
          .setValues(enhHeaderValues.slice(2, 4));
        compositeCell.setHorizontalAlignment('right');
        infoCell.setHorizontalAlignment('left');
      } //
      else {
        // Legacy: write header (formulas + values)
        // Break merges across header once (prevents Service error)
        // sh.getRange('A1:P4').breakApart();

        // Row 1
        sh.getRange('A1:P1')
          .setFormulasR1C1(legacyHeaderFormulas.slice(0, 1))
          .setValues(legacyHeaderValues.slice(0, 1));

        // Rows 3–4 (skip row 2)
        sh.getRange('A3:P4')
          .setFormulasR1C1(legacyHeaderFormulas.slice(2, 4))
          .setValues(legacyHeaderValues.slice(2, 4));
        compositeCell.setHorizontalAlignment('right');
        infoCell.setHorizontalAlignment('left');
        enhancedCheckCells.setFontColor(headerBgColor);

        // Conditional formatting: add ONLY 1st 3 rules (body scope is A5:P80)
        // replaceLegacyRules(legacyTemplateSheet, sh);
      }

      setScoreColor(sh);
      testCodeCell.setValue(sheetName);
      testDataCell.setValue('done');

      Logger.log(`Updated ${sheetName}`);
    }
  } catch (err) {
    Logger.log(err && err.stack ? err.stack : String(err));
    ui.alert('Update is not finished. Please run update again.');
    return;
  }

  dataSheet.getRange(2, 13, testCodes.length, 2).setValue('');
}

function addScaleDownFormatting() {
  const startTime = Date.now();
  const maxDuration = 5.5 * 60 * 1000; // 5m30s
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const testCodes = getActTestCodes();
  const dataSheet = ss.getSheetByName('Data');
  const testFilterCell = dataSheet.getRange('M2');
  testFilterCell.setFormula('=unique(A2:A)');
  const templateSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('actTemplateSsId'));
  const legacyTemplateSheet = templateSs.getSheetByName('Admin legacy');

  try {
    // Iterate only the test sheets by name, instead of scanning ss.getSheets()
    for (const sheetName of testCodes) {
      if (Date.now() - startTime > maxDuration) {
        Logger.log('Exiting loop after 5 minutes and 30 seconds.');
        throw new Error('Process exceeded maximum duration of 5 minutes and 30 seconds.');
      }

      if (sheetName > '202502') {
        Logger.log(`Skipping enhanced ${sheetName}`);
        continue;
      }

      const testDataRow = testCodes.indexOf(sheetName) + 2;
      const testDataCell = dataSheet.getRange(testDataRow, 14);
      if (testDataCell.getValue() === 'done') {
        Logger.log(`${sheetName} done, skipping`);
        continue;
      }

      const sh = ss.getSheetByName(sheetName);
      if (!sh) {
        Logger.log(`${sheetName} sheet not found, skipping`);
        continue;
      }

      Logger.log(`Starting ${sheetName}`)

      replaceLegacyRules(legacyTemplateSheet, sh);
      testDataCell.setValue('done');
    }
  }
  catch (err) {
    Logger.log(err && err.stack ? err.stack : String(err));
    ui.alert('Update is not finished. Please run update again.');
    return;
  }

  dataSheet.getRange(2, 13, testCodes.length, 2).setValue('');
}

function replaceLegacyRules(legacyTemplateSheet, targetSheet) {
  const templateRules = legacyTemplateSheet.getConditionalFormatRules();

  // clear existing rules
  targetSheet.setConditionalFormatRules([]);

  const rebuilt = templateRules.map((r) => {
    const bool = r.getBooleanCondition && r.getBooleanCondition();
    if (!bool || bool.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
      throw new Error('Legacy rules be CUSTOM_FORMULA rules.');
    }

    // 1) Rebuild ranges so they belong to targetSheet
    const targetRanges = r.getRanges().map(tr => targetSheet.getRange(tr.getA1Notation()));

    const formula = bool.getCriteriaValues()[0];

    const b = SpreadsheetApp.newConditionalFormatRule()
      .setRanges(targetRanges)
      .whenFormulaSatisfied(formula);

    const bg = r.getBackground && r.getBackground();
    if (bg) b.setBackground(bg);

    const fontColor = r.getFontColor && r.getFontColor();
    if (fontColor) b.setFontColor(fontColor);

    const bold = r.isBold && r.isBold();
    if (bold !== null && bold !== undefined) b.setBold(!!bold);

    const italic = r.isItalic && r.isItalic();
    if (italic !== null && italic !== undefined) b.setItalic(!!italic);

    const underline = r.isUnderline && r.isUnderline();
    if (underline !== null && underline !== undefined) b.setUnderline(!!underline);

    const strikethrough = r.isStrikethrough && r.isStrikethrough();
    if (strikethrough !== null && strikethrough !== undefined) b.setStrikethrough(!!strikethrough);

    return b.build();
  });

  // Replace ALL conditional formatting rules on the target sheet
  targetSheet.setConditionalFormatRules(rebuilt);
}

function addSatTestSheets(adminSsId = SpreadsheetApp.getActiveSpreadsheet().getId()) {
  const testCodes = getSatTestCodes();
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

function addActTestSheets(adminSsId, adminIndexAdjustment=2) {
  // adminIndexAdjustment = number of analysis sheets preceding test sheets
  let adminSs;
  if (!adminSsId) {
    adminSs = SpreadsheetApp.getActiveSpreadsheet();
  }
  else {
    adminSs = SpreadsheetApp.openById(adminSsId);
  }

  const studentSsId = adminSs.getSheetByName('Student responses').getRange('B1').getValue();
  const studentSs = SpreadsheetApp.openById(studentSsId);
  const actTemplateSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('actTemplateSsId'));
  const adminTemplateSheet = actTemplateSs.getSheetByName('Admin legacy');
  const adminEnhancedTemplateSheet = actTemplateSs.getSheetByName('Admin enhanced adjusted');
  const studentTemplateSheet = actTemplateSs.getSheetByName('Student legacy');
  const studentEnhancedTemplateSheet = actTemplateSs.getSheetByName('Student enhanced');

  const templateSheet = adminSs.getSheetByName('201904');
  const templateHeaderCell = templateSheet.getRange('A1');
  const templateBodyCell = templateSheet.getRange('A5');
  const headerBgColor = templateHeaderCell.getBackground();
  const headerFontColor = templateHeaderCell.getFontColorObject().asRgbColor().asHexString();
  const bodyFontColor = templateBodyCell.getFontColorObject().asRgbColor().asHexString();


  const spreadsheets = [
    {
      'ss': studentSs,
      'templateSheet': studentTemplateSheet,
      'enhancedTemplateSheet': studentEnhancedTemplateSheet,
      'indexAdjustment': 1,
      'isAdmin': false
    },
    {
      'ss': adminSs,
      'templateSheet': adminTemplateSheet,
      'enhancedTemplateSheet': adminEnhancedTemplateSheet,
      'indexAdjustment': adminIndexAdjustment,     // 1-indexed + # of analysis sheets,
      'isAdmin': true
    }
  ]

  const testCodes = getActTestCodes();
  for (obj of spreadsheets) {
    for (testCode of testCodes) {
      const testSheet = obj.ss.getSheetByName(testCode);

      if (!testSheet) {
        Logger.log(`Adding ${testCode} sheet to ${obj.ss.getName()}`);

        let sheetToCopy = obj.templateSheet;
        if (testCode.includes('MC')) {
          sheetToCopy = obj.enhancedTemplateSheet;
        }

        const newSheet = sheetToCopy.copyTo(obj.ss).setName(testCode);
        newSheet.getRange('B1').setValue(testCode);

        const headerRange = newSheet.getRange('A1:P4');
        headerRange.setBackground(headerBgColor).setFontColor(headerFontColor).setBorder(true, true, true, true, true, true, headerBgColor, SpreadsheetApp.BorderStyle.SOLID);
        newSheet.getRange('A5:P80').setFontColor(bodyFontColor);

        if (obj.isAdmin) {
          setScoreColor(testSheet);
        }

        const testCodeIndex = testCodes.indexOf(testCode);

        obj.ss.setActiveSheet(newSheet);
        if (testCodeIndex > 0) {
          const prevTest = testCodes[testCodeIndex - 1];
          const prevTestPosition = obj.ss.getSheetByName(prevTest).getIndex() + 1 || 1;
          Logger.log(`Previous sheet: ${prevTest}, index ${prevTestPosition}`);
          obj.ss.moveActiveSheet(prevTestPosition);
        }
        else if (obj.isAdmin) {
          obj.ss.moveActiveSheet(obj.indexAdjustment + 1); // Move after analysis sheets
        }
        else {
          obj.ss.moveActiveSheet(1);
        }
      }
    }
  }
}

function setScoreColor(sheet) {
  const scoreColor = '#93c47d'
  sheet.getRange('F1').setBackground(scoreColor);
  sheet.getRangeList(['B3', 'F3', 'J3', 'N3']).setBorder(true, true, true, true, true, true, scoreColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
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

function getIdFromImportFormula(cell) {
  const formulaString = cell.getFormula();
  if (!formulaString) return "";

  const formula = formulaString.toString().trim();

  // 1. Extract first argument inside IMPORTRANGE(...)
  const argMatch = formula.match(/^=IMPORTRANGE\(([^,]+),/i);
  if (!argMatch) return "";

  let firstArg = argMatch[1].trim();

  // 2. Strip wrapping quotes
  if ((firstArg.startsWith('"') && firstArg.endsWith('"')) ||
      (firstArg.startsWith("'") && firstArg.endsWith("'"))) {
    firstArg = firstArg.slice(1, -1);
  }

  // 3. If it looks like a URL, extract ID
  const urlMatch = firstArg.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch) return urlMatch[1];

  // 4. If it’s a cell reference (with optional sheet prefix)
  if (/^'?[^'!]+!'?[A-Z]+\d+$/.test(firstArg) || /^[A-Z]+\d+$/.test(firstArg)) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet, cellRef;

    if (firstArg.includes("!")) {
      const parts = firstArg.split("!");
      const sheetName = parts[0].replace(/^'|'$/g, "");
      cellRef = parts[1];
      sheet = ss.getSheetByName(sheetName);
    } else {
      sheet = cell.getSheet();
      cellRef = firstArg;
      Logger.log(`${sheet.getName()} ${cellRef}`);
    }

    if (!sheet) return "";
    const value = sheet.getRange(cellRef).getValue().toString().trim();
    return value;
  }

  // 5. Otherwise assume it’s already an ID
  try {
    SpreadsheetApp.openById(firstArg); // will throw if invalid
    return firstArg;
  } catch (e) {
    return ""; // avoid recursive errorNotification loop
  }
}

function sortFoldersByName(folderIterator) {
  if (!folderIterator.hasNext()) return [];

  const folderList = [];
  while (folderIterator.hasNext()) {
    const folder = folderIterator.next();
    folderList.push({folder, name: folder.getName()});
  }

  folderList.sort((a, b) => a.name.localeCompare(b.name));

  return folderList.map(obj => obj.folder);
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
  let studentName, scoreReportFolderId, studentFolder, revBackendSheet;

  if (ssType === 'sat') {
    revBackendSheet = adminSs.getSheetByName('Rev sheet backend');
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
      }

      if (!scoreReportFolderId) {
        scoreReportFolderId = studentFolder.createFolder('Score reports').getId();
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

  return scoreReportFolderId;
}


function errorNotification(error, ssId) {
  let ss;
  try {
    ss = SpreadsheetApp.openById(ssId);
  }
  catch {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  try {
    const htmlOutput = HtmlService.createHtmlOutput(`<p>We have been notified of the following error: ${error.message}</p><p>${error.stack}`)
    // const htmlOutput = HtmlService.createHtmlOutput(`<p>Please copy-paste the following details and send to ${ADMIN_EMAIL}. Sorry about that!</p><p> ${error.message}</p><p>${error.stack}`)
      .setWidth(500) //optional
      .setHeight(300); //optional
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Error`);
  }
  catch {}

  const editorEmails = []
  ss.getEditors().forEach(editor => editorEmails.push(editor.getEmail()));

  const url = getDriveUrl(ssId);
  const message = `
    <p>Error details: ${error.stack}</p>
    <p><a href="${url}" target="_blank">${url}</a></p>
    <p>Editors: ${editorEmails}</p>
  `
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: `Spreadsheet error: ${error.message}`,
    htmlBody: message
  });

  Logger.log(error.message + '\n\n' + error.stack);
  throw new Error(error.message + '\n\n' + error.stack);
}

function getDriveUrl(id) {
  const file = DriveApp.getFileById(id);
  if (file) {
    return file.getUrl();
  } //
  else {
    const folder = DriveApp.getFolderById(id);

    if (folder) {
      return folder.getUrl();
    }
    else {
      Logger.log('ID not recognized as file or folder');
      return;
    }
  }
}