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