function transferSatStudentData() {
  transferOldStudentData();
}

function transferOldStudentData() {
  const startTime = new Date().getTime(); // Record the start time
  let ui = SpreadsheetApp.getUi();

  let prompt = ui.prompt(
    'Old admin analysis spreadsheet URL or ID - leave blank \r\n' +
    'to update student sheet with this admin spreadsheet\'s data:',
    ui.ButtonSet.OK_CANCEL);
  let oldAdminDataUrl = prompt.getResponseText();
  if (prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }

  let oldAdminSsId = getIdFromDriveUrl(oldAdminDataUrl);
  if (!oldAdminSsId) {
    oldAdminSsId = SpreadsheetApp.getActiveSpreadsheet().getId();
  }

  syncSatStudentData(oldAdminSsId, startTime);
}

function syncSatStudentData(oldAdminSsId=SpreadsheetApp.getActiveSpreadsheet().getId(), startTime=new Date().getTime()) {
  let ui, newAdminSs;
  try {
    ui = SpreadsheetApp.getUi();
    let htmlOutput = HtmlService
      .createHtmlOutput('<p>You may close the spreadsheet or this popup. If you click "Cancel" above, you will need to restore the previous version of the new admin AND new student spreadsheets by clicking File > Version history > See version history</p><button onclick="google.script.host.close()">OK</button>')
      .setWidth(400)
      .setHeight(150);
    ui.showModalDialog(htmlOutput, 'Do not cancel this script');
    newAdminSs = SpreadsheetApp.getActiveSpreadsheet();
  }
  catch(e) {
    // function not run from spreadsheet
  }

  // temporarily relax permissions
  DriveApp.getFileById(oldAdminSsId).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  DriveApp.getFileById(newStudentSsId).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);

  let oldAdminSs, newStudentData, initialImportFunction;
  try {
    oldAdminSs = SpreadsheetApp.openById(oldAdminSsId);
    if (!newAdminSs) {
      newAdminSs = oldAdminSs;
    }

    const newAdminSsId = newAdminSs.getId()
    const newStudentSsId = newAdminSs.getSheetByName('Student responses').getRange('B1').getValue();
    const newStudentSs = SpreadsheetApp.openById(newStudentSsId);
    const maxDuration = 5.5 * 60 * 1000; // 5 minutes and 30 seconds in milliseconds
    newStudentData = newAdminSs.getSheetByName('Student responses');
    initialImportFunction = newStudentData.getRange('A3').getFormula();
    const oldRevSheet = oldAdminSs.getSheetByName('Rev sheets');
    const newRevSheet = newAdminSs.getSheetByName('Rev sheets');
    let newQbSheet = newAdminSs.getSheetByName('Question bank data');
    let timestampRange = newQbSheet.getRange(2, 11, getLastFilledRow(newQbSheet, 11));

    if (oldAdminSsId !== newAdminSsId) {
      // temporarily set old admin data imports
      newStudentData.getRange('A3').setValue('=importrange("' + oldAdminSsId + '", "Question bank data!$A$1:$G10000")');
      newStudentData.getRange('H3').setValue('=importrange("' + oldAdminSsId + '", "Question bank data!$I$1:$K10000")');

      // build reviewed column
      newStudentData.getRange('L3').setValue('=importrange("' + oldAdminSsId + '", "Practice test data!E1:M")');
      const newPracticeDataSheet = newAdminSs.getSheetByName('Practice test data');
      const reviewedLookup = '=xlookup(E2, \'Student responses\'!$L$4:$L$10000, \'Student responses\'!$T$4:$T$10000,"")';
      const reviewedStartCell = newPracticeDataSheet.getRange('M2');
      reviewedStartCell.setFormula(reviewedLookup);
      const filter = newPracticeDataSheet.getFilter();
      if(filter) {
        filter.remove();
      }
      SpreadsheetApp.flush();
      newPracticeDataSheet.getRange(1, 1, newPracticeDataSheet.getMaxRows(), 13).createFilter();
      let reviewedRange = newPracticeDataSheet.getRange('M2:M10000');
      reviewedStartCell.autoFill(reviewedRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
      reviewedRange.setValues(reviewedRange.getValues());

      // build timestamp column
      let timestampLookup = '=xlookup(A2, \'Student responses\'!$A$4:$A$10000, \'Student responses\'!$J$4:$J$10000,"")';
      let timestampStartCell = newQbSheet.getRange('K2');
      timestampStartCell.setFormula(timestampLookup);
      timestampStartCell.autoFill(timestampRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    }

    let timestampValues = timestampRange.getValues();

    for (let row = 0; row < timestampValues.length; row ++) {
      let ssRow = row + 2;
      if (timestampValues[row][0] === '') {
        timestampValues[row][0] = '=if(or(G' + ssRow + '="",I' + ssRow + '=""),"",if(K' + ssRow + ',K' + ssRow + ',if(I' + ssRow + '="","",now())))'
      }
    }
    timestampRange.setValues(timestampValues);
    timestampRange.setNumberFormat('MM/dd/yyy h:mm:ss');


    if (oldRevSheet) {
      if (oldAdminSsId !== newAdminSs.getId()) {
        let oldRevBackend = oldAdminSs.getSheetByName('Rev sheet backend');
        let oldRevDataId = oldRevBackend.getRange('U3').getValue();
        let oldStudentName = oldRevBackend.getRange('K2').getValue();
        let oldRevSs = SpreadsheetApp.openById(oldRevDataId)
        let oldRevDataStudentSheet = oldRevSs.getSheetByName(oldStudentName);
        let oldStudentRevData = oldRevDataStudentSheet.getRange(1,1,oldRevDataStudentSheet.getLastRow(), oldRevDataStudentSheet.getLastColumn()).getValues();
        let newRevBackend = newAdminSs.getSheetByName('Rev sheet backend');
        let newRevDataCell = newRevBackend.getRange('U3');
        let newRevDataId = newRevDataCell.getValue();
        let newStudentName = newRevBackend.getRange('K2').getValue();
        let newRevDataStudentSheet = SpreadsheetApp.openById(newRevDataId).getSheetByName(newStudentName);
        let rwRevSheetNumber = newAdminSs.getSheetByName('RW Rev sheet').getRange('E1').getValue();
        let mathRevSheetNumber = newAdminSs.getSheetByName('Math Rev sheet').getRange('E1').getValue();


        if ((oldRevDataId !== newRevDataId) && rwRevSheetNumber == 0 && mathRevSheetNumber == 0 && newRevDataCell && oldRevSs) {
          let prompt = ui.alert('Older and newer spreadsheet versions have different Rev Data spreadsheets. If Rev sheets were created using the older version, it is recommended that you use the older version of Rev Data. Would you like to connect the new student to the old Rev Data sheet?', ui.ButtonSet.YES_NO);

          if(prompt === ui.Button.YES) {
            const newStudentSs = SpreadsheetApp.openById(newStudentSsId);
            const newStudentRevDataCell = newStudentSs.getSheetByName('Question bank data').getRange('U3');

            if (newStudentRevDataCell) {
              newRevDataCell.setValue(oldRevDataId);
              newStudentRevDataCell.setValue(oldRevDataId);
              htmlOutput = HtmlService
                .createHtmlOutput('<p></p><button onclick="google.script.host.close()">OK</button>')
                .setWidth(400)
                .setHeight(150);
              ui.showModalDialog(htmlOutput, 'Rev Data sheet updated');
            }
            else {
              htmlOutput = HtmlService
                .createHtmlOutput('<p></p><button onclick="google.script.host.close()">OK</button>')
                .setWidth(400)
                .setHeight(150);
              ui.showModalDialog(htmlOutput, 'Rev Data sheet could not be updated');
            }

          }
          else {
            newRevDataStudentSheet.getRange(1,1,oldRevDataStudentSheet.getLastRow(), oldRevDataStudentSheet.getLastColumn()).setValues(oldStudentRevData);
          }
        }
        else {
          newRevDataStudentSheet.getRange(1,1,oldRevDataStudentSheet.getLastRow(), oldRevDataStudentSheet.getLastColumn()).setValues(oldStudentRevData);
        }
      }

      const oldRwRevResponseRange = oldRevSheet.getRange(4, 4, getLastFilledRow(oldRevSheet, 4));
      const oldMathRevResponseRange = oldRevSheet.getRange(4, 9, getLastFilledRow(oldRevSheet, 9));
      const oldRwRevResponseValues = oldRwRevResponseRange.getValues();
      const oldRwRevResponseFormulas = oldRwRevResponseRange.getFormulas();
      const oldMathRevResponseValues = oldMathRevResponseRange.getValues();
      const oldMathRevResponseFormulas = oldMathRevResponseRange.getFormulas();

      for (let i = 0; i < oldRwRevResponseValues.length; i++) {
        if (oldRwRevResponseValues[i][0] === '') {
          oldRwRevResponseValues[i] = oldRwRevResponseFormulas[i];
        }
      }

      for (let i = 0; i < oldMathRevResponseValues.length; i++) {
        if (oldMathRevResponseValues[i][0] === '') {
            oldMathRevResponseValues[i] = oldMathRevResponseFormulas[i];

        }
      }

      newRevSheet.getRange(4, 4, oldRwRevResponseValues.length).setValues(oldRwRevResponseValues);
      newRevSheet.getRange(4, 9, oldMathRevResponseValues.length).setValues(oldMathRevResponseValues);
    }

    let answerSheets = getSatTestCodes(oldAdminSs);
    let testScores = [];

    for (let s in answerSheets) {
      let sheetName = answerSheets[s];
      let oldSheet = oldAdminSs.getSheetByName(sheetName);
      if (oldSheet) {
        let subScore = oldSheet.getRange('G1:I1').getValues();
        testScores.push({
          'test': sheetName,
          'scores': subScore
        })
      }
    }

    answerSheets.push('Reading & Writing', 'Math', 'SLT Uniques');

    let allNewAdminSheetValues = [];
    let allNewStudentSheetValues = [];

    for (const sheetName of answerSheets) {
      let newAdminSheet = newAdminSs.getSheetByName(sheetName);
      let newStudentSheet = newStudentSs.getSheetByName(sheetName);

      if (newAdminSheet) {
        Logger.log('Transferring answers for ' + sheetName);
        let newAnswersLevel1, newAnswersLevel2, newAnswersLevel3;
        let newStudentLevel1, newStudentLevel2, newStudentLevel3;

        if (sheetName === 'Rev sheets') {
          newAnswersLevel1 = newAdminSheet.getRange(5, 4, getLastFilledRow(newAdminSheet, 3) - 4);
          newAnswersLevel2 = newAdminSheet.getRange(5, 9, getLastFilledRow(newAdminSheet, 8) - 4);
          newStudentLevel1 = newStudentSheet.getRange(5, 4, getLastFilledRow(newAdminSheet, 3) - 4);
          newStudentLevel2 = newStudentSheet.getRange(5, 9, getLastFilledRow(newAdminSheet, 8) - 4);
        }
        else {
          newAnswersLevel1 = newAdminSheet.getRange(5, 3, getLastFilledRow(newAdminSheet, 2) - 4);
          newAnswersLevel2 = newAdminSheet.getRange(5, 7, getLastFilledRow(newAdminSheet, 6) - 4);
          newStudentLevel1 = newStudentSheet.getRange(5, 3, getLastFilledRow(newAdminSheet, 2) - 4);
          newStudentLevel2 = newStudentSheet.getRange(5, 7, getLastFilledRow(newAdminSheet, 6) - 4);
          if (sheetName === 'SLT Uniques') {
            // newAdminSheet.getRange(5, 1, newAdminSheet.getMaxRows() - 5, 2).clearContent();
            // newAdminSheet.getRange(5, 5, newAdminSheet.getMaxRows() - 5, 2).clearContent();
            // newStudentSheet.getRange(5, 1, newStudentSheet.getMaxRows() - 5, 2).clearContent();
            // newStudentSheet.getRange(5, 5, newStudentSheet.getMaxRows() - 5, 2).clearContent();
            Logger.log('SLT Uniques logic needed');
          }
          else {
            newAnswersLevel3 = newAdminSheet.getRange(5, 11, getLastFilledRow(newAdminSheet, 10) - 4);
            newStudentLevel3 = newStudentSheet.getRange(5, 11, getLastFilledRow(newAdminSheet, 10) - 4);
          }
        }

        let newAdminRanges = [newAnswersLevel1, newAnswersLevel2, newAnswersLevel3];
        let newStudentRanges = [newStudentLevel1, newStudentLevel2, newStudentLevel3];

        for (let i = 0; i < newAdminRanges.length; i++) {
          if (newAdminRanges[i] && newStudentRanges[i]) {
            const currentTime = new Date().getTime();
            if (currentTime - startTime > maxDuration) {
              Logger.log("Exiting loop after 5 minutes and 30 seconds.");
              throw new Error("Process exceeded maximum duration of 5 minutes and 30 seconds. Cleaning up.");
            }
            let newAdminSheetValues = newAdminRanges[i].getValues();
            let newAdminSheetFormulas = newAdminRanges[i].getFormulas();
            let newStudentSheetValues = [];

            for (let row = 0; row < newAdminSheetValues.length; row++) {
              // set blank cells blank for student sheet
              if (newAdminSheetValues[row][0] === 'not found') {
                newAdminSheetValues[row][0] = '';
              }
              // save nonblank cells as values for admin sheet
              else if (newAdminSheetValues[row][0] !== '') {
                newAdminSheetFormulas[row][0] = newAdminSheetValues[row][0];
              }

              newStudentSheetValues.push(newAdminSheetValues[row]);
            }
            // Ensure the number of rows in the source and destination ranges match
            if (newStudentSheetValues.length === newStudentRanges[i].getNumRows()) {
              allNewAdminSheetValues.push({ range: newAdminRanges[i], values: newAdminSheetFormulas });
              allNewStudentSheetValues.push({ range: newStudentRanges[i], values: newStudentSheetValues });
            } //
            else {
              throw Error(`Mismatch in row count for ${sheetName} at level ${i + 1}`);
            }
          }
          else {
            Logger.log('newAdminRanges[i]=' + newAdminRanges[i] + ', newAdminRanges[i]=' + newStudentRanges[i])
          }
        }
        let testScore = testScores.find(score => score.test === sheetName);
        if (testScore) {
          newAdminSheet.getRange('G1:I1').setValues(testScore.scores);
        }
      }
    }

    allNewAdminSheetValues.forEach(item => item.range.setValues(item.values));
    allNewStudentSheetValues.forEach(item => item.range.setValues(item.values));
  }
  catch (err) {
    errorNotification(err, newAdminSsId);
  }
  finally {
    // revert student ID and SS permissions
    newStudentData.getRange('A3').setValue(initialImportFunction);
    newStudentData.getRange('H3').setValue('');
    newStudentData.getRange('T3').setValue('');
    DriveApp.getFileById(oldAdminSsId).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    DriveApp.getFileById(newStudentSsId).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  }

  try {
    htmlOutput = HtmlService.createHtmlOutput('<a href="https://docs.google.com/spreadsheets/d/' + newStudentSsId + '" target="_blank" onclick="google.script.host.close()">Student answer sheet</a>')
      .setWidth(250)
      .setHeight(50);
    ui.showModalDialog(htmlOutput, 'Data transfer complete');
  }
  catch (e) {
    Logger.log('Data transfer complete. Student answer sheet URL: https://docs.google.com/spreadsheets/d/' + newStudentSsId);
  }
}
