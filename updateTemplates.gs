function updateConceptData(adminSsId, studentSsId = null) {
  if (!adminSsId) {
    const adminSs = SpreadsheetApp.getActiveSpreadsheet();
    adminSsId = adminSs.getId();
    studentSsId = adminSs.getSheetByName('Student responses').getRange('B1').getValue();
  }

  try {
    const satDataSsId = '1XoANqHEGfOCdO1QBVnbA3GH-z7-_FMYwoy7Ft4ojulE';
    const qbDataSheetName = 'Question bank data updated ' + dataLatestDate;
    const ptDataSheetName = 'Practice test data updated ' + dataLatestDate;
    const qbDataSh = SpreadsheetApp.openById(satDataSsId).getSheetByName(qbDataSheetName);
    const qbDataVals = qbDataSh.getRange(1,1, getLastFilledRow(qbDataSh, 1), 15).getValues();

    for (id of [adminSsId, studentSsId]) {
      if (!id) continue;

      const ss = SpreadsheetApp.openById(id);
      const isAdminSs = (id === adminSsId);

      for (subject of subjectData) {
        const sh = ss.getSheetByName(subject.name);

        const mergeRanges = sh.getDataRange().getMergedRanges();
        mergeRanges
          .filter(range => range.getRow() >= subject.rowOffset)
          .forEach(range => range.breakApart());

        const conceptData = getConceptHeaderRows(ss, subject);
        for (concept of conceptData) {
          for (let level = 1; level < 4; level ++) {
            let count = 0;
            for (let r = 0; r < qbDataVals.length; r++) {
              if (qbDataVals[r][3].toLowerCase() === concept['name'].toLowerCase() && Number(qbDataVals[r][4]) === level && qbDataVals[r][6].slice(0,3) !== 'SAT' && qbDataVals[r][6].slice(0,4) !== 'PSAT' && qbDataVals[r][6].slice(0,3) !== 'SLT') {
                count ++;
              }
            }
            concept['level' + level] = count;
          }
        }

        Logger.log('Calculating row modifications for ' + subject.name);
        const modifications = [];
        for (concept of conceptData) {
          const rowsNeeded = Math.max(concept['level1'], concept['level2'], concept['level3']) + 4;
          const nextConcept = conceptData.find(c => c.id === concept['id'] + 1);
          let rowsToAdd, endRow;

          if (nextConcept) {
            endRow = nextConcept.row;
          }
          else {
            endRow = sh.getMaxRows() + 1;
          }
          rowsToAdd = concept['row'] + rowsNeeded - endRow;

          if (rowsToAdd > 0) {
            modifications.push({
              'position': endRow - 1,
              'rows': rowsToAdd
            });
          }
          else if (rowsToAdd < 0) {
            modifications.push({
              'position': concept['row'] + rowsNeeded + rowsToAdd - 1, // rowsToAdd negative
              'rows': rowsToAdd
            });
          }
        }
        modifyRowsAtPositions(sh, modifications);

        const shNewRange = sh.getRange(subject['rowOffset'], 1, sh.getMaxRows() - subject['rowOffset'], sh.getMaxColumns());
        shNewRange.setNumberFormat('@STRING@');
        const shNewVals = shNewRange.getValues();
        const shFormulas = shNewRange.getFormulas();
        const newConceptRows = shNewVals.map(row => row[1]);

        Logger.log(conceptData);

        for (let level = 1; level < 4; level++) {
          const levelStartCol = (level - 1) * 4;

          for (concept of conceptData) {

            // Since newConceptRows starts at subject['rowOffset'] and includes blanks,
            // concept['index'] is 0-indexed position of concept name starting at 1st concept
            concept['index'] = newConceptRows.indexOf(concept['name']);

            const levelRow = concept['index'] + 2;
            shNewVals[levelRow][levelStartCol + 1] = shFormulas[levelRow][levelStartCol + 1];

            for (qNum = 1; qNum <= concept['level' + level]; qNum++) {
              const qRow = levelRow + qNum;

              // Find the matching row in Question bank data
              const dataRow = qbDataVals.find(row => row[3].toLowerCase() === concept['name'].toLowerCase() && Number(row[4]) === level && Number(row[5]) === qNum);

              shNewVals[qRow] = []
              shNewVals[qRow][levelStartCol] = dataRow[0];
              shNewVals[qRow][levelStartCol + 1] = level + '.' + qNum;
            }
          }

          const outputValues = [];
          for (let i = 0; i < shNewVals.length; i++) {
            outputValues.push([
              shNewVals[i][levelStartCol],
              shNewVals[i][levelStartCol + 1]
            ]);
          }

          sh.getRange(subject['rowOffset'], levelStartCol + 1, outputValues.length, 2).setValues(outputValues);
          sh.getRange(subject['rowOffset'], levelStartCol + 2, outputValues.length).setHorizontalAlignment('center').setFontWeight('bold')
        }


        const answerFormulaR1C1 = '=let(worksheetNum,R[0]C[-1],if(worksheetNum="","", if(iserror(search(".",worksheetNum)),"", let(id,R[0]C[-2], xlookup(id,\'Student responses\'!R4C1:C1,\'Student responses\'!R4C8:C8,"not found")))))'
        const correctedFormulaR1C1 = '=let(worksheetNum,R[0]C[-2],if(worksheetNum="","", if(left(worksheetNum,5)="Level","Corrected", if(iserror(search(".",worksheetNum)), "", let(id,R[0]C[-3], result,xlookup(id,\'Question bank data\'!R2C1:C1,\'Question bank data\'!R2C8:C8,"not found"), if(result=R[0]C[-1],"",result))))))'

        if (isAdminSs) {
          for (let level = 1; level < 4; level++) {
            const answerCol = 4 * (level - 1) + 3;
            if (studentSsId) {
              const answerRange = sh.getRange(subject['rowOffset'] + 3, answerCol, sh.getMaxRows() - subject['rowOffset'] - 3);
              const answerVals = answerRange.getValues();

              for (let r = 0; r < answerVals.length; r ++) {
                let startRow = subject['rowOffset'] + 3 + r;
                let numRows = 0;

                // Set R1C1 formula for consecutive blank rows, leave values as is
                while (r < answerVals.length && answerVals[r][0] === '') {
                  numRows ++;
                  r ++;
                }
                if (numRows > 0) {
                  sh.getRange(startRow, answerCol, numRows).setFormulaR1C1(answerFormulaR1C1);
                }
              }
              answerRange.setHorizontalAlignment('center').setFontWeight('normal');
            }
            const correctedRange = sh.getRange(subject['rowOffset'] + 3, answerCol + 1, sh.getMaxRows() - subject['rowOffset'] - 3);
            correctedRange.setHorizontalAlignment('center').setFontWeight('normal').setFormulaR1C1(correctedFormulaR1C1);
          }

          Logger.log('formulas updated for ' + subject.name);
        }

        for (concept of conceptData) {
          const headerStartRow = concept['index'] + subject['rowOffset'];
          sh.getRange(headerStartRow, 2, 1, 11).merge().setHorizontalAlignment('left');
          sh.getRange(headerStartRow, 2, 3, 11).setFontWeight('bold');
        }

        sh.getRange('A1:A').setFontColor('#ffffff');
        sh.getRange('E1:E').setFontColor('#ffffff');
        sh.getRange('I1:I').setFontColor('#ffffff');

        modifyConceptFormatRules(sh, isAdminSs);
        // mergeRanges.forEach(range => range.merge());
      }

      if (isAdminSs) {
        const qbDataFormula = ss.getSheetByName('Question bank data').getRange('A1').getFormula();
        const qbCommaIndex = qbDataFormula.indexOf(',');
        const qbFormulaStart = qbDataFormula.toString().slice(0,qbCommaIndex);
        const ptDataFormula = ss.getSheetByName('Practice test data').getRange('A1').getFormula();
        const ptCommaIndex = ptDataFormula.indexOf(',');
        const ptFormulaStart = ptDataFormula.toString().slice(0,ptCommaIndex);
        const revBackendSheet = ss.getSheetByName('Rev sheet backend');

        if (revBackendSheet) {
          const satAdminDataSsId = revBackendSheet.getRange('U5').getValue();
          const satAdminDataSs = SpreadsheetApp.openById(satAdminDataSsId);

          if (!satAdminDataSs.getSheetByName(qbDataSheetName)) {
            const newQbDataSheet = satAdminDataSs.getSheetByName('Question bank data').copyTo(satAdminDataSs).setName(qbDataSheetName);
            newQbDataSheet.getRange('A1').setFormula('=importrange("' + satDataSsId + '", "' + qbDataSheetName + '!A1:H10000")')
          }

          if (!satAdminDataSs.getSheetByName(ptDataSheetName)) {
            const newPtDataSheet = satAdminDataSs.getSheetByName('Practice test data').copyTo(satAdminDataSs).setName(ptDataSheetName);
            newPtDataSheet.getRange('A1').setFormula('=importrange("' + satDataSsId + '", "' + ptDataSheetName + '!A1:J10000")')
          }
        }

        ss.getSheetByName('Question bank data').getRange('A1').setValue(qbFormulaStart + ', "' + qbDataSheetName + '!A1:H10000")');
        ss.getSheetByName('Practice test data').getRange('A1').setValue(ptFormulaStart + ', "' + ptDataSheetName + '!A1:J10000")');
        Logger.log('sat admin data URLs updated')
      }
      else {
        // Student sheets do not always have separate studentDataId cell
        let satStudentDataSsId = ss.getSheetByName('Question bank data').getRange('U7').getValue();
        const qbImportCell = ss.getSheetByName('Question bank data').getRange('A1');
        const ptImportCell = ss.getSheetByName('Practice test data').getRange('A1');
        const qbImportValue = qbImportCell.getFormula();
        const ptImportValue = ptImportCell.getFormula();
        const newQbImportValue = qbImportValue.replace(/bank data.*?!/, `bank data updated ${dataLatestDate}!`);
        const newPtImportValue = ptImportValue.replace(/test data.*?!/, `test data updated ${dataLatestDate}!`);
        qbImportCell.setFormula(newQbImportValue);
        ptImportCell.setFormula(newPtImportValue);
        Logger.log('sat student data URLs updated');

        if (!satStudentDataSsId){
          const openQuoteIndex = qbImportValue.indexOf('"');
          const closeQuoteIndex = qbImportValue.indexOf('"',openQuoteIndex + 1);
          satStudentDataSsId = qbImportValue.slice(openQuoteIndex + 1, closeQuoteIndex);
          Logger.log('satStudentDataSsId: ' + satStudentDataSsId);
          if (satStudentDataSsId.includes('/')) {
            satStudentDataSsId = satStudentDataSsId.split('/d/')[1].split('/')[0];
          }
        }

        const satStudentDataSs = SpreadsheetApp.openById(satStudentDataSsId);
        if (!satStudentDataSs.getSheetByName(qbDataSheetName)) {
          const newQbDataSheet = satStudentDataSs.getSheetByName('Question bank data').copyTo(satStudentDataSs).setName(qbDataSheetName);
          newQbDataSheet.getRange('A1').setFormula('=importrange("' + satDataSsId + '", "' + qbDataSheetName + '!A1:G10000")')
        }
        if (!satStudentDataSs.getSheetByName(ptDataSheetName)) {
          const newPtDataSheet = satStudentDataSs.getSheetByName('Practice test data').copyTo(satStudentDataSs).setName(ptDataSheetName);
          newPtDataSheet.getRange('A1').setFormula('=importrange("' + satDataSsId + '", "' + ptDataSheetName + '!A1:E10000")')
        }
      }
    }
  }
  catch (err) {
    errorNotification(err, adminSsId);
  }
}

function updateConceptDataAllSpreadsheets() {
  updateAllSpreadsheets(updateConceptData, 2)
}

function updateAllSpreadsheets(updateFunction, dataRow) {
  const ui = SpreadsheetApp.getUi();
  const templateSs = SpreadsheetApp.getActiveSpreadsheet();
  const revBackendSheet = templateSs.getSheetByName('Rev sheet backend');
  const studentsFolderIdCell = revBackendSheet.getRange(dataRow, 25);
  let studentsFolderId = studentsFolderIdCell.getValue();

  if (studentsFolderId) {
    Logger.log(`Continuing folder update for ${updateFunction.name}`);
  } //
  else {
    const prompt = ui.prompt('URL of Drive folder where student folders are located (leave blank to use the parent folder of the template folder, which is where student folders are saved by default)');
    const response = prompt.getResponseText();
    studentsFolderId = getIdFromDriveUrl(response);
  }

  if (!studentsFolderId) {
    const templateSsId = templateSs.getId();
    const templateParentFolder = DriveApp.getFileById(templateSsId).getParents().next().getParents().next();
    studentsFolderId = templateParentFolder.getId();
  }

  const studentsFolder = DriveApp.getFolderById(studentsFolderId);
  const allStudentDataCell = revBackendSheet.getRange(dataRow, 24);
  const allStudentDataValue = allStudentDataCell.getValue();
  const allStudentData = allStudentDataValue ? JSON.parse(allStudentDataValue) : [];
  const allCompletedAdminFolderIds = allStudentData.map(s => s.adminFolderId);

  const adminFolders = studentsFolder.getFolders();
  const adminFolderList = sortFoldersByName(adminFolders);

  let startingFolderIndexCell = revBackendSheet.getRange(dataRow, 26);
  let startingFolderIndex = startingFolderIndexCell.getValue() || 0;
  studentsFolderIdCell.setValue(studentsFolderId);
  startingFolderIndexCell.setValue(startingFolderIndex);

  for (let i = startingFolderIndex; i < adminFolderList.length; i ++) {
    const adminFolder = adminFolderList[i];
    const adminFolderId = adminFolder.getId();
    Logger.log(`Starting ${adminFolder.getName()}`);
    const adminFolderIndex = adminFolderList.findIndex(folder => folder.getId() === adminFolderId);
    const studentData = getStudentData(adminFolderId, 'sat');

    if (!allCompletedAdminFolderIds.includes(studentData.folderId)) {
      if (studentData.satAdminSsId) {
        const htmlOutput = HtmlService.createHtmlOutput(`${adminFolderIndex + 1} of ${adminFolderList.length}`)
          .setWidth(400)
          .setHeight(50);
        ui.showModalDialog(htmlOutput, `Starting update for ${studentData.name}`);
        updateFunction(studentData.satAdminSsId, studentData.satStudentSsId);
        allStudentData.push(studentData);
      }
      else {
        Logger.log(`Completed update for ${studentData.name}`);
        const htmlOutput = HtmlService.createHtmlOutput(`Continuing`)
          .setWidth(250)
          .setHeight(50);
        ui.showModalDialog(htmlOutput, `No SAT admin sheet found for ${studentData.name}`);
      }
      const allStudentDataStr = JSON.stringify(allStudentData);
      allStudentDataCell.setValue(allStudentDataStr);

      Logger.log(`Completed update for ${studentData.name}`);
      const htmlOutput = HtmlService.createHtmlOutput(`
        <a href="https://docs.google.com/spreadsheets/d/${studentData.satAdminSsId}" target="_blank">${studentData.name}'s SAT admin spreadsheet</a><br>
        <a href="https://docs.google.com/spreadsheets/d/${studentData.satStudentSsId}" target="_blank">${studentData.name}'s SAT student spreadsheet</a>`)
        .setWidth(250)
        .setHeight(80);
      ui.showModalDialog(htmlOutput, `Update complete for ${studentData.name}`);
    }
    else {
      Logger.log(`${studentData.name} already completed`);
    }

    startingFolderIndex += 1;
    startingFolderIndexCell.setValue(startingFolderIndex);
  }
  ui.alert(`Update complete for all students in ${studentsFolder.getName()}`)
  studentsFolderIdCell.clear();
  startingFolderIndexCell.clear();
  allStudentDataCell.clear();
}


function updateSatWorksheetLinks() {
  const templateAdminSs = SpreadsheetApp.getActiveSpreadsheet();
  const revBackendSheet = templateAdminSs.getSheetByName('Rev sheet backend');
  const adminDataSsId = revBackendSheet.getRange('U5').getValue();
  const adminDataSs = SpreadsheetApp.openById(adminDataSsId);
  const templateStudentSsId = templateAdminSs.getSheetByName('Student responses').getRange('B1').getValue();
  const templateStudentSs = SpreadsheetApp.openById(templateStudentSsId);
  const studentDataImportCell = templateStudentSs.getSheetByName('Question bank data').getRange('A1');
  const studentDataSsId = getIdFromImportFormula(studentDataImportCell);
  const studentDataSs = SpreadsheetApp.openById(studentDataSsId);

  if (!adminDataSs.getSheetByName('Links')) {
    const linksSheet = adminDataSs.insertSheet('Links');
    const linksDataImportValue = '=importrange("1XoANqHEGfOCdO1QBVnbA3GH-z7-_FMYwoy7Ft4ojulE","Links!I1:O50")';
    linksSheet.getRange('A1').setValue(linksDataImportValue);
  }

  if (!studentDataSs.getSheetByName('Links')) {
    const linksSheet = studentDataSs.insertSheet('Links');
    const linksDataImportValue = `=importrange("${adminDataSsId}","Links!A1:D50")`;
    linksSheet.getRange('A1').setValue(linksDataImportValue);
  }

  updateAllSpreadsheets(updateWorksheetLinks, 3);
  return;
}


function updateWorksheetLinks(adminSsId, studentSsId) {
  for (ssId of [adminSsId, studentSsId]) {
    const ss = SpreadsheetApp.openById(ssId);
    const isAdminSs = (ssId === adminSsId);
    let linksImportValue, linksImportRange;

    if (isAdminSs) {
      linksImportValue = '=importrange(U5,"Links!A1:G50")';
      const revBackendSheet = ss.getSheetByName('Rev sheet backend');

      if (!revBackendSheet) continue;

      linksImportRange = revBackendSheet.getRange('T19');
    } //
    else {
      linksImportValue = '=importrange(U7,"Links!A1:D50")';
      linksImportRange = ss.getSheetByName('Question bank data').getRange('T19');

      // Save studentDataSsId to student ss
      const studentQbSheet = ss.getSheetByName('Question bank data')
      const studentDataImportCell = studentQbSheet.getRange('A1');
      const studentDataSsId = getIdFromImportFormula(studentDataImportCell);

      studentQbSheet.getRange('T7:U7').setValues([['Student data SS ID', studentDataSsId]]);
      studentQbSheet.getRange('T9:U9').setValues([["",""]]);
    }

    linksImportRange.setValue(linksImportValue);

    for (subject of subjectData) {
      const sh = ss.getSheetByName(subject.name);
      const correctedColVal = isAdminSs ? "Corrected" : "";
      const linkTableRange = isAdminSs ? "'Rev sheet backend'!R20C20:R70C26" : "'Question bank data'!R20C20:R70C23"
      const templateRowVals = [
        [
          `=let(level,ifs(column()=2,1,column()=6,2,column()=10,3),hyperlink(vlookup(R[-2]C2,${linkTableRange},level+1,FALSE),"Level "&level))`,
          "'Answer",
          correctedColVal,
          "",
          `=let(level,ifs(column()=2,1,column()=6,2,column()=10,3),hyperlink(vlookup(R[-2]C2,${linkTableRange},level+1,FALSE),"Level "&level))`,
          "'Answer",
          correctedColVal,
          "",
          `=let(level,ifs(column()=2,1,column()=6,2,column()=10,3),hyperlink(vlookup(R[-2]C2,${linkTableRange},level+1,FALSE),"Level "&level))`,
          "'Answer",
          correctedColVal,
        ]
      ]

      const conceptData = getConceptHeaderRows(ss, subject);

      for (concept of conceptData) {
        sh.getRange(concept.row + 2, 2, 1, 11).setValues(templateRowVals);
      }
    }
  }
}


function getConceptHeaderRows(ss, subjectData) {
  const conceptData = [];
  const sh = ss.getSheetByName(subjectData['name']);
  const conceptColVals = sh.getRange(subjectData['rowOffset'], 2, sh.getMaxRows() - subjectData['rowOffset']).getValues();
  let id = 1;

  for (let x = 0; x < conceptColVals.length; x++) {
    if (cats.includes(conceptColVals[x][0])) {
      var row = x + subjectData['rowOffset'];
      conceptData.push({
        'name': conceptColVals[x][0],
        'row': row,                        // row 1-indexed
        'id': id
      });
      id ++;
    }
  }

  return conceptData;
}


function modifyRowsAtPositions(sheet, modifications) {
  // Sort modifications in descending order of positions to avoid shifting issues
  modifications.sort((a, b) => b.position - a.position);

  // Apply each modification
  modifications.forEach(mod => {
    if (mod.rows > 0) {
      // Insert rows if `rows` is positive
      sheet.insertRows(mod.position, mod.rows);
    } //
    else if (mod.rows < 0) {
      // Delete rows if `rows` is negative
      sheet.deleteRows(mod.position, Math.abs(mod.rows));
    }
  });
  Logger.log('Row modifications complete')
}


function modifyConceptFormatRules(sheet, isAdminSs) {
  const alertColor = '#cc0000';
  const darkGreen = '#6aa84f';
  const lightGreen = '#b7e1cd';
  const darkRed = '#e06666';
  const lightRed = '#f4c7c3';
  const grey = '#f3f3f3';
  const yellow = '#fff2cc';
  // Get all existing conditional formatting rules
  var rules = sheet.getConditionalFormatRules();

  // Create an array to store updated rules
  var updatedRules = [];

  // Iterate through each rule
  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    var bgColor = rule.getBooleanCondition().getBackgroundObject().asRgbColor().asHexString();

    if (bgColor === alertColor) {
      // Modify the rule
      rule = SpreadsheetApp.newConditionalFormatRule()
        .setRanges([sheet.getRange('A10:A'), sheet.getRange('E10:E'), sheet.getRange('I10:I')])
        .whenFormulaSatisfied('=and(len(A10)<>8,B10<>"",B9<>"")')
        .setBackground(alertColor)
        .setFontColor('#ffffff')
        .build();
    }
    else if (bgColor === yellow) {
      rule = SpreadsheetApp.newConditionalFormatRule()
        .setRanges([sheet.getRange('C10:C'), sheet.getRange('G10:G'), sheet.getRange('K10:K')])
        .whenFormulaSatisfied('=and(or(C10="",C10="-"),B10<>"",B9<>"")')
        .setBackground(yellow)
        .build();
    }

    if (isAdminSs) {
      if (bgColor === darkGreen) {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .setRanges([sheet.getRange('C10:C'), sheet.getRange('G10:G'), sheet.getRange('K10:K')])
          .whenFormulaSatisfied('=and(C10<>"",D10="",isformula(C10))')
          .setBackground(darkGreen)
          .setFontColor('#ffffff')
          .build();
      }
      else if (bgColor === lightGreen) {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .setRanges([sheet.getRange('C10:C'), sheet.getRange('G10:G'), sheet.getRange('K10:K')])
          .whenFormulaSatisfied('=and(C10<>"",D10="",B9<>"")')
          .setBackground(lightGreen)
          .build();
      }
      else if (bgColor === darkRed) {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .setRanges([sheet.getRange('C10:C'), sheet.getRange('G10:G'), sheet.getRange('K10:K')])
          .whenFormulaSatisfied('=and(C10<>"",D10<>"",isformula(C10),B9<>"")')
          .setBackground(darkRed)
          .setFontColor('#ffffff')
          .build();
      }
      else if (bgColor === lightRed) {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .setRanges([sheet.getRange('C10:C'), sheet.getRange('G10:G'), sheet.getRange('K10:K')])
          .whenFormulaSatisfied('=and(C10<>"",D10<>"",B9<>"")')
          .setBackground(lightRed)
          .build();
      }
      else if (bgColor === grey) {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .setRanges([sheet.getRange('C10:C'), sheet.getRange('G10:G'), sheet.getRange('K10:K')])
          .whenFormulaSatisfied('=and(not(isformula(C10)),C10="",B10<>"",B9<>"")')
          .setBackground(grey)
          .build();
      }
    }

    updatedRules.push(rule); // Add updated or existing rule
  }

  // Reapply all rules to the sheet
  sheet.setConditionalFormatRules(updatedRules);
}

function findClientFileIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssName = ss.getName();
  const clientName = ssName.slice(ssName.indexOf('Template for') + 13);
  const clientDataRange = ss.getSheetByName('Rev sheet backend').getRange('T1:U10');
  const clientData = clientDataRange.getValues();
  const satAdminDataId = clientData[4][1];
  const satStudentDataId = clientData[6][1];
  const revDataId = clientData[2][1];
  let studentsFolderId = clientData[7][1];

  // TODO: Automate entering satStudentDataId into clientData[6][1]
  try {
    DriveApp.getFolderById(studentsFolderId);
  }
  catch (e) {
    const ui = SpreadsheetApp.getUi();
    const prompt = ui.prompt('Students folder URL or ID', ui.ButtonSet.OK_CANCEL);
    response = prompt.getResponseText();

    if (prompt.getSelectedButton() == ui.Button.CANCEL) {
      return;
    } //
    else {
      studentsFolderId = getIdFromDriveUrl(response);
    }

    clientData[7][1] = studentsFolderId;
  }

  let studentsStr = clientData[8][1];
  const studentsJSON = studentsStr ? JSON.parse(studentsStr) : [];

  const client = {
    name: clientName,
    satAdminSsId: ss.getId(),
    satStudentSsId: ss.getSheetByName('Student responses').getRange('B1').getValue(),
    satAdminDataId: satAdminDataId,
    satStudentDataId: satStudentDataId,
    revDataId: revDataId,
    studentsFolderId: studentsFolderId,
    studentsDataJSON: studentsJSON
  }

  const students = getAllStudentData(client);
  studentsStr = JSON.stringify(students);
  clientData[8][1] = studentsStr;

  clientDataRange.setValues(clientData);

  return client;
}


function ssUpdate202505(students = {}, client = {}) {
  const qbResArrayVal =
    '=let(testCodes,\'Practice test data\'!$E$2:E, testResponses,\'Practice test data\'!$K$2:$K,\n' +
    '    worksheetRanges,vstack(\'Reading & Writing\'!A10:C,\'Reading & Writing\'!E10:G,\'Reading & Writing\'!I10:K,\n' +
    '                           Math!A13:C,Math!E13:G,Math!I13:K,\'SLT Uniques\'!A5:C,\'SLT Uniques\'!E5:G),\n' +
    '    z,counta(A2:A),\n' +
    '    map(offset(G1,1,0,z),offset(B1,1,0,z),offset(E1,1,0,z),offset(A1,1,0,z),\n' +
    '    lambda(    skillCode,       subject,         difficulty,      id,\n' +
    '           if(or(left(skillCode,3)="SAT",left(skillCode,4)="PSAT"),\n' +
    '           xlookup(skillCode,testCodes,testResponses,"not found"),\n' +
    '           vlookup(id,worksheetRanges,3,FALSE)))))'

  const sltFilterR1C1 = "=FILTER({'Question bank data'!R2C1:C1,'Question bank data'!R2C7:C7},left('Question bank data'!R2C7:C7,3)=\"SLT\",'Question bank data'!R2C2:C2=R[-3]C[1])"

  const studentsDataCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rev sheet backend').getRange('U9');

  // iterate through all folders in Students including template folder
  for (let student of students) {
    if (student.satAdminSsId && !student.updateComplete) {
      Logger.log('Starting student: ' + student['name']);

      for (let ssId of [student.satAdminSsId, student.satStudentSsId]) {
        const ss = SpreadsheetApp.openById(ssId);

        for (let sheetName of ['Reading & Writing', 'Math']) {
          const sheet = ss.getSheetByName(sheetName);

          sheet.getRange('A10:A').setFontColor('#ffffff');
          sheet.getRange('E10:E').setFontColor('#ffffff');
          sheet.getRange('I10:I').setFontColor('#ffffff');
        }
      }

      const adminSs = SpreadsheetApp.openById(student['satAdminSsId']);
      adminSs.getSheetByName('Question bank data').getRange('I2').setValue(qbResArrayVal);

      const revBackendSheet = adminSs.getSheetByName('Rev sheet backend');
      if (revBackendSheet) {
        revBackendSheet.getRange('U5').setValue(client['satAdminDataId']);
      }
      Logger.log('Admin values updated')
      adminSs.getSheetByName('SLT Uniques').getRange('B5').setValue('');
      adminSs.getSheetByName('SLT Uniques').getRange('F5').setValue('');
      adminSs.getSheetByName('SLT Uniques').getRange('A5').setValue(sltFilterR1C1);
      adminSs.getSheetByName('SLT Uniques').getRange('E5').setValue(sltFilterR1C1);
      Logger.log('SLT Uniques filter fixed')

      const studentSs = SpreadsheetApp.openById(student['satStudentSsId']);
      const studentRevSheet = studentSs.getSheetByName('Rev sheets');
      if (studentRevSheet) {
        studentRevSheet.getRange('B5:B').setFontWeight('bold');
        studentRevSheet.getRange('F5:F').setFontWeight('bold');
      }
      modifyTestFormatRules(student['satAdminSsId']);
      updateConceptData(student['satAdminSsId'], student['satStudentSsId']);
      student.updateComplete = true;
      studentsDataCell.setValue(JSON.stringify(students));
    }
    else if (student.updateComplete) {
      Logger.log(`${student.name} data is updated`);
    }
    else if (!student.satAdminSsId) {
      student.updateComplete = true;
      studentsDataCell.setValue(JSON.stringify(students));
      Logger.log(`No SAT data found for ${student.name}`)
    }
  }

  updateSatDataSheets(client.satAdminDataId, client.satStudentDataId);

  let htmlOutput = HtmlService
    .createHtmlOutput('<p></p><button onclick="google.script.host.close()">OK</button>')
    .setWidth(400)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'All spreadsheets updated');
}

function updateSatDataSheets(satAdminDataSsId, satStudentDataSsId) {
  const satAdminDataSs = SpreadsheetApp.openById(satAdminDataSsId);
  const satStudentDataSs = SpreadsheetApp.openById(satStudentDataSsId);
  let newAdminQbSheet = satAdminDataSs.getSheetByName('Question bank data updated ' + dataLatestDate);
  let newAdminPtSheet = satAdminDataSs.getSheetByName('Practice test data updated ' + dataLatestDate);
  let newStudentQbSheet = satStudentDataSs.getSheetByName('Question bank data updated ' + dataLatestDate);
  let newStudentPtSheet = satStudentDataSs.getSheetByName('Practice test data updated ' + dataLatestDate);

  if (!newAdminQbSheet) {
    newAdminQbSheet = satAdminDataSs.getSheetByName('Question bank data').copyTo(satAdminDataSs).setName('Question bank data updated ' + dataLatestDate);
  }
  if (!newAdminPtSheet) {
    newAdminPtSheet = satAdminDataSs.getSheetByName('Practice test data').copyTo(satAdminDataSs).setName('Practice test data updated ' + dataLatestDate);
  }
  if (!newStudentQbSheet) {
    newStudentQbSheet = satStudentDataSs.getSheetByName('Question bank data').copyTo(satStudentDataSs).setName('Question bank data updated ' + dataLatestDate);
  }
  if (!newStudentPtSheet) {
    newStudentPtSheet = satStudentDataSs.getSheetByName('Practice test data').copyTo(satStudentDataSs).setName('Practice test data updated ' + dataLatestDate);
  }

  newAdminQbSheet.getRange('A1').setValue('=importrange("1XoANqHEGfOCdO1QBVnbA3GH-z7-_FMYwoy7Ft4ojulE", "Question bank data updated ' + dataLatestDate + '!A1:H10000")');
  newAdminPtSheet.getRange('A1').setValue('=importrange("1XoANqHEGfOCdO1QBVnbA3GH-z7-_FMYwoy7Ft4ojulE", "Practice test data updated ' + dataLatestDate + '!A1:J10000")');
  Logger.log('sat admin data sheets updated')
  newStudentQbSheet.getRange('A1').setValue('=importrange("' + satAdminDataSsId + '", "Question bank data updated ' + dataLatestDate + '!A1:G10000")');
  newStudentPtSheet.getRange('A1').setValue('=importrange("' + satAdminDataSsId + '", "Practice test data updated ' + dataLatestDate + '!A1:E10000")');
  Logger.log('sat student data sheets updated')
}

function modifyTestFormatRules(satAnswerSheetId='1FW_3GIWmytdrgBdfSuIl2exy9hIAnQoG8IprF8k9uEY') {
  const ss = SpreadsheetApp.openById(satAnswerSheetId);
  const tests = getSatTestCodes();

  for (test of tests) {
    const sh = ss.getSheetByName(test);
    if (sh) {
      var rules = sh.getConditionalFormatRules();
      const alertColor = '#cc0000';
      const updatedRules = [];

      for (var i = 0; i < rules.length; i++) {
        var rule = rules[i];
        var bgColor = rule.getBooleanCondition().getBackgroundObject().asRgbColor().asHexString();

        if (bgColor !== alertColor) {
          updatedRules.push(rule);
        }
      }

      const rwRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([sh.getRange('A5:A31'), sh.getRange('E5:E31'), sh.getRange('I5:I31')])
      .whenFormulaSatisfied('=A5<>$B$2&" "&B5')
      .setBackground(alertColor)
      .setFontColor('#ffffff')
      .build();

      const mathRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([sh.getRange('A36:A57'), sh.getRange('E36:E57'), sh.getRange('I36:I57')])
      .whenFormulaSatisfied('=A36<>$B$33&" "&B36')
      .setBackground(alertColor)
      .setFontColor('#ffffff')
      .build();

      updatedRules.push(rwRule, mathRule);
      sh.setConditionalFormatRules(updatedRules);
    }
  }
  Logger.log('Test sheets formatting updated');
}