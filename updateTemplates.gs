function updateConceptData(adminSsId = '1sdnVpuX8mVkpTdrqZgwz7zph1NdFpueX6CP45JHiNP8', studentSsId = null) {
  const qbDataSh = SpreadsheetApp.openById('1XoANqHEGfOCdO1QBVnbA3GH-z7-_FMYwoy7Ft4ojulE').getSheetByName('Question bank data updated 03/2025');
  const qbDataVals = qbDataSh.getRange(1,1, getLastFilledRow(qbDataSh, 1), 15).getValues();

  const subjectData = [
    {
      'name': 'Reading & Writing',
      'rowOffset': 7,
    },
    {
      'name': 'Math',
      'rowOffset': 10
    }
  ]
  
  for (id of [adminSsId, studentSsId]) { 

    const ss = SpreadsheetApp.openById(id);
    let isAdminSs;
    
    if (id === adminSsId) {
      isAdminSs = true;
    }
    else {
      isAdminSs = false;
    }
    
    for (subject of subjectData) {
      const sh = ss.getSheetByName(subject['name']);
      const conceptColVals = sh.getRange(subject['rowOffset'], 2, sh.getMaxRows() - subject['rowOffset']).getValues();
      const conceptData = [];
      const modifications = [];
      let id = 1;

      for (let x = 0; x < conceptColVals.length; x++) {
        if (cats.includes(conceptColVals[x][0])) {
          var row = x + subject['rowOffset'];
          conceptData.push({
            'name': conceptColVals[x][0],
            'row': row,                        // row 1-indexed
            'id': id
          });
          id ++;
        }
      }

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
      
      modifyConceptFormatRules(sh, isAdminSs);
    }

    if (isAdminSs) {
      ss.getSheetByName('Question bank data').getRange('A1').setValue('=importrange(\'Rev sheet backend\'!U5, "Question bank data updated ' + dataLatestDate + '!A1:H10000")');
      ss.getSheetByName('Practice test data').getRange('A1').setValue('=importrange(\'Rev sheet backend\'!U5, "Practice test data updated ' + dataLatestDate + '!A1:J10000")');
      Logger.log('sat admin data URLs updated')
    }
    else {
      // Student sheets do not have separate studentDataId cell 
      const qbImportCell = ss.getSheetByName('Question bank data').getRange('A1');
      const ptImportCell = ss.getSheetByName('Practice test data').getRange('A1');
      const qbImportValue = qbImportCell.getFormula();
      const ptImportValue = ptImportCell.getFormula();
      const newQbImportValue = qbImportValue.replace(/data.*?!/, `data updated ${dataLatestDate}!`);
      const newPtImportValue = ptImportValue.replace(/data.*?!/, `data updated ${dataLatestDate}!`);
      qbImportCell.setFormula(newQbImportValue);
      ptImportCell.setFormula(newPtImportValue);
      Logger.log('sat student data URLs updated');
    }
  }
}


function addTestSheets(adminSsId) {
  const testCodes = getTestCodes();
  
  if (!adminSsId) {
    adminSsId = SpreadsheetApp.getActiveSpreadsheet().getId();
  }

  const adminSs = SpreadsheetApp.openById(adminSsId);
  const studentSsId = adminSs.getSheetByName('Student responses').getRange('B1').getValue();
  const studentSs = SpreadsheetApp.openById(studentSsId);
  const adminTemplateSs = SpreadsheetApp.openById('1_AG-LWa0r8WPKwdD3ejzBB7bAAooOt_rEsv6TA8BzkY');
  const adminTemplateSheet = adminTemplateSs.getSheetByName('SAT4');
  const studentTemplateSheet = SpreadsheetApp.openById('1Zx1hxyuBsY6QzlMRn13K4LIovCSbsSJv_E6P5w94KoQ').getSheetByName('SAT4');

  for (testCode of testCodes) {
    const testNumberPosition = testCode.indexOf('SAT') + 3;
    const testType = testCode.substring(0, testNumberPosition)
    const testNumber = testCode.substring(testNumberPosition);

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


function modifyRowsAtPositions(sheet, modifications) {
  // Sort modifications in descending order of positions to avoid shifting issues
  modifications.sort((a, b) => b.position - a.position);

  // Apply each modification
  modifications.forEach(mod => {
    if (mod.rows > 0) {
      // Insert rows if `rows` is positive
      sheet.insertRows(mod.position, mod.rows);
    } else if (mod.rows < 0) {
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
    studentsFolderId = prompt.getResponseText();

    if (prompt.getSelectedButton() == ui.Button.CANCEL) {
      return;
    } else if (prompt.getResponseText().includes('/folders/')) {
      studentsFolderId = prompt.getResponseText().split('/folders/')[1].split(/[/?]/)[0];
      Logger.log(`studentsFolderId: ${studentsFolderId}`);
    } else {
    }
    clientData[7][1] = studentsFolderId;
  }
  
  let studentsStr = clientData[8][1];
  const studentsJSON = studentsStr ? JSON.parse(studentsStr) : [];

  const client = {
    'name': clientName,
    'satAdminSsId': ss.getId(),
    'satStudentSsId': ss.getSheetByName('Student responses').getRange('B1').getValue(),
    'satAdminDataId': satAdminDataId,
    'satStudentDataId': satStudentDataId,
    'revDataId': revDataId,
    'studentsFolderId': studentsFolderId,
    'studentsDataJSON': studentsJSON
  }

  const students = findStudentFileIds(client);
  studentsStr = JSON.stringify(students);
  clientData[8][1] = studentsStr;

  clientDataRange.setValues(clientData);

  return client;
}

function findStudentFileIds(
  client={
    'index': null,
    'name': null,
    'studentsFolderId': null,
    'studentsDataJSON': null
  })
  {
  const index = client.index || 1;

  Logger.log(index + '. ' + client.name + ' started');

  const studentFolders = DriveApp.getFolderById(client.studentsFolderId).getFolders();
  let students = client.studentsDataJSON;
  const studentFolderIds = [];

  while (studentFolders.hasNext()) {
    const studentFolder = studentFolders.next();
    const studentFolderId = studentFolder.getId();
    const studentFolderName = studentFolder.getName();

    studentFolderIds.push(studentFolderId);
    
    if (!studentFolderName.includes('Ξ')) {  
      const studentObj = students.find(obj => obj.folderId === studentFolderId);
      if (studentObj) {
        Logger.log(`${studentFolderName} found with folder ID ${studentFolderId}`);

        if (studentObj && studentObj.name !== studentFolderName) {
          // Update the name property
          studentObj.name = studentFolderName;
          Logger.log(`Updated name for folder ID ${studentFolderId} to ${studentFolderName}`);
        }
      }
      else {
        Logger.log(`Adding ${studentFolderName} to students data`);
        const adminFiles = studentFolder.getFiles();
        let satAdminSsId, satStudentSsId;

        while (adminFiles.hasNext()) {
          const adminFile = adminFiles.next();
          if (adminFile.getName().toLowerCase().includes('sat admin')) {
            satAdminSsId = adminFile.getId();
            break;
          }
        }

        if (satAdminSsId) {
          satStudentSsId = SpreadsheetApp.openById(satAdminSsId).getSheetByName('Student responses').getRange('B1').getValue();
        }

        students.push({
          'name': studentFolderName,
          'folderId': studentFolderId,
          'satAdminSsId': satAdminSsId,
          'satStudentSsId': satStudentSsId,
          'updateComplete': false
        })
      }
    }

    // only for clients with grouped student folders
    // const subfolders = studentFolder.getFolders();
    // while (subfolders.hasNext()) {
    //   const subfolder = subfolders.next();
    //   const subfiles = subfolder.getFiles();

    //   while (subfiles.hasNext()) {
    //     const subfile = subfiles.next();
    //     if (subfile.getName().toLowerCase().includes('sat admin')) {
    //       satAdminSsId = subfile.getId();
    //       break;
    //     }
    //   }

    //   if (satAdminSsId) {
    //     satStudentSsId = SpreadsheetApp.openById(satAdminSsId).getSheetByName('Student responses').getRange('B1').getValue();
    //   }

    //   students.push({
    //     'name': subfolder.getName(),
    //     'satAdminSsId': satAdminSsId,
    //     'satStudentSsId': satStudentSsId
    //   })
    // }
  }

  students = students.filter(student => studentFolderIds.includes(student.folderId));

  return students;
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
  const tests = getTestCodes();

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