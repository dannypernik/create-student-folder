function createSatScoreReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const testCode = ui.prompt('Test code').getResponseText().toUpperCase();
  const testCodes = getSatTestCodes();

  if (!testCodes.includes(testCode)) {
    var htmlOutput = HtmlService.createHtmlOutput(`${testCode} is not a valid test code`)
      .setWidth(250) //optional
      .setHeight(100); //optional
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Error`);
    return;
  }
  const testSheet = ss.getSheetByName(testCode);

  if (testSheet) {
    const testHeaderValues = testSheet.getRange('A1:M2').getValues();
    const rwScore = parseInt(testHeaderValues[0][6]) || 0;
    const mScore = parseInt(testHeaderValues[0][8]) || 0;
    const dateSubmitted = testHeaderValues[1][3];

    if (rwScore && mScore) {
      testData = {
        'test': testCode,
        'rw': rwScore,
        'm': mScore,
        'total': rwScore + mScore,
        'date': dateSubmitted,
        'isNew': true
      }
    }
    else {
      ui.alert(`RW and/or Math score not present in G1 and I1 of ${testCode} sheet`);
      return;
    }
  }
  else {
    ui.alert(`${testCode} sheet not found`);
    return;
  }

  createSatScoreReportPdf(ss.getId(), testData);

}


async function createSatScoreReportPdf(adminSsId, currentTestData) {
  try {
    const adminSs = adminSsId ? SpreadsheetApp.openById(adminSsId) : SpreadsheetApp.getActiveSpreadsheet();
    adminSsId = adminSsId ? adminSsId : adminSs.getId();
    const adminSsName = adminSs.getName();
    const studentName = adminSsName.slice(adminSsName.indexOf('-') + 2);

    const scoreReportFolderId = getScoreReportFolderId(adminSsId);

    const pdfName = currentTestData.test + ' answer analysis - ' + studentName + '.pdf';
    const answerSheetId = adminSs.getSheetByName(currentTestData.test).getSheetId();
    const analysisSheetId = adminSs.getSheetByName(currentTestData.test + ' analysis').getSheetId();

    Logger.log(`Starting ${currentTestData.test} score report for ${studentName}`);

    const answerFileId = savePdfSheet(adminSsId, answerSheetId, studentName);
    const analysisFileId = savePdfSheet(adminSsId, analysisSheetId, studentName);

    const fileIdsToMerge = [analysisFileId, answerFileId];

    const mergedPdf = await mergePDFs(fileIdsToMerge, scoreReportFolderId, pdfName);
    // const mergedBlob = mergedFile.getBlob();
    // const pdfFile = DriveApp.getFolderById(scoreReportFolderId).createFile(mergedBlob).setName(pdfName);
    const pdfUrl = mergedPdf.getUrl();

    var htmlOutput = HtmlService.createHtmlOutput(`<a href="${pdfUrl}" target="_blank">${currentTestData.test} score report</a>`)
      .setWidth(250) //optional
      .setHeight(100); //optional
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, `PDF complete`);

    Logger.log(studentName + ' ' + currentTestData.test + ' score report complete');
  } catch (err) {
    Logger.log(err.stack);
    throw new Error(err.message + '\n\n' + err.stack);
  }
}


function createActScoreReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const testCode = ui.prompt('Test code').getResponseText().toUpperCase();
  const testSheet = ss.getSheetByName(testCode);

  if (testSheet) {
    const testData = getActTestData(testCode);

    if (!testData.total) {
      const response = ui.prompt(`Total score is missing from F1 of sheet ${testData.test}, suggesting that the test is incomplete. Proceed?`, ui.ButtonSet.YES_NO)

      if (response.getSelectedButton() === ui.Button.NO) {
        return;
      }

      createActScoreReportPdf(ss.getId(), testData);
    }
  }
  else {
    ui.alert(`${testCode} sheet not found`);
    return;
  }  
}


async function createActScoreReportPdf(spreadsheetId, currentTestData) {
  try {
    const spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId();
    const ssName = spreadsheet.getName();
    const studentName = ssName.slice(ssName.indexOf('-') + 2);
    const dataSheet = spreadsheet.getSheetByName('Data');
    let scoreReportFolderId;

    if (dataSheet.getRange('V1').getValue() === 'Score report folder ID:' && dataSheet.getRange('W1').getValue() !== '') {
      scoreReportFolderId = dataSheet.getRange('W1').getValue();
    } //
    else {
      scoreReportFolderId = getScoreReportFolderId(spreadsheetId);
    }

    if (dataSheet.getRange('W1').getValue() !== scoreReportFolderId) {
      dataSheet.getRange('V1:W1').setValues([['Score report folder ID:', scoreReportFolderId]]);
    }

    const pdfName = `ACT ${currentTestData.test} answer analysis - ${studentName}.pdf`;
    const answerSheetId = spreadsheet.getSheetByName(currentTestData.test).getSheetId();
    const analysisSheetName = currentTestData.test + ' analysis';
    let analysisSheet = spreadsheet.getSheetByName(analysisSheetName);

    if (!analysisSheet) {
      const testAnalysisSheet = spreadsheet.getSheetByName('Test analysis');
      analysisSheet = testAnalysisSheet.copyTo(spreadsheet).setName(analysisSheetName);
    }
    const analysisPivot = analysisSheet.getPivotTables()[0];

    if (analysisPivot) {
      const filters = analysisPivot.getFilters();
      const testCodeColumnIndex = 1;

      for (var i = 0; i < filters.length; i++) {
        var filter = filters[i];

        if (filter.getSourceDataColumn() === testCodeColumnIndex) {
          var newCriteria = SpreadsheetApp.newFilterCriteria().setVisibleValues([currentTestData.test]).build();

          filter.setFilterCriteria(newCriteria);
          break;
        }
      }
    } //
    else {
      Logger.log('No Pivot Table found at the specified range.');
    }

    const answerSheetPosition = spreadsheet.getSheetByName(currentTestData.test).getIndex();

    if (analysisSheet.getIndex() !== answerSheetPosition + 1) {
      spreadsheet.setActiveSheet(analysisSheet);
      spreadsheet.moveActiveSheet(answerSheetPosition + 1);
    }

    const analysisSheetId = analysisSheet.getSheetId();

    Logger.log(`Starting ${currentTestData.test} score report for ${studentName}`);

    const answerSheetMargins = { top: '0.3', bottom: '0.25', left: '0.35', right: '0.35' };
    const answerFileId = savePdfSheet(spreadsheetId, answerSheetId, studentName, answerSheetMargins);

    const pageBreakRow = getActPageBreakRow(analysisSheet, 3);
    const analysisSheetMargin = { top: '0.25', bottom: '0.25', left: '0.25', right: '0.25' };

    if (pageBreakRow < 80) {
      const analysisSheetWidth = 1306; // 1296px + 10px interior border padding
      const pixelsPerInch = analysisSheetWidth / 8; // (1296px + 10px) wide for 8in page width = 163.25px/inch
      const headerHeightInches = (24 * 8) / pixelsPerInch; // 24px header height at 96dpi
      const bodyHeightInches = ((pageBreakRow - 8) * 21) / pixelsPerInch; // 8 rows of header
      const marginTopInches = 0.25;
      const pageBreakHeight = headerHeightInches + bodyHeightInches + marginTopInches;
      const bottomMargin = 11 - pageBreakHeight; // 11in total height - pageBreakHeight;

      analysisSheetMargin.bottom = String(Math.floor(bottomMargin * 1000) / 1000);
    }

    Logger.log(analysisSheetMargin.bottom);
    const analysisFileId = savePdfSheet(spreadsheetId, analysisSheetId, studentName, analysisSheetMargin);

    const fileIdsToMerge = [analysisFileId, answerFileId];

    const mergedFile = await mergePDFs(fileIdsToMerge, scoreReportFolderId, pdfName);
    const mergedBlob = mergedFile.getBlob();
    const pdfFile = DriveApp.getFolderById(scoreReportFolderId).createFile(mergedBlob).setName(pdfName);
    const pdfUrl = pdfFile.getUrl();

    var htmlOutput = HtmlService.createHtmlOutput(`<a href="${pdfUrl}">ACT ${currentTestData.test} score report</a>`)
      .setWidth(250) //optional
      .setHeight(100); //optional
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, `PDF complete`);
    Logger.log(studentName + ' ' + currentTestData.test + ' score report complete');
  } catch (err) {
    Logger.log(err.stack);
    throw new Error(err.message + '\n\n' + err.stack);
  }
}


async function mergePDFs(fileIds, destinationFolderId, name = 'merged.pdf', attempt = 1) {
  const validFileIds = fileIds.filter(isValidPdf);

  if (validFileIds.length !== fileIds.length) {
    if (attempt > 5) {
      throw new Error('mergePDFs: Too many attempts, some files are still not valid PDFs.');
    }
    // Exponential backoff: wait 2^attempt * 1000 ms
    const waitMs = Math.pow(2, attempt) * 1000;
    Logger.log(`mergePDFs: Not all files are valid PDFs. Retrying in ${waitMs / 1000}s (attempt ${attempt})`);
    Utilities.sleep(waitMs);
    return await mergePDFs(fileIds, destinationFolderId, name, attempt + 1);
  }
  // Retrieve PDF data as byte arrays
  const data = fileIds.map((id) => new Uint8Array(DriveApp.getFileById(id).getBlob().getBytes()));

  // Load pdf-lib from CDN
  const cdnjs = 'https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js';
  eval(
    UrlFetchApp.fetch(cdnjs)
      .getContentText()
      .replace(/setTimeout\(.*?,.*?(\d*?)\)/g, 'Utilities.sleep($1);return t();')
  );

  // Merge PDFs
  const pdfDoc = await PDFLib.PDFDocument.create();
  for (let i = 0; i < data.length; i++) {
    const pdfData = await PDFLib.PDFDocument.load(data[i]);
    const pages = await pdfDoc.copyPages(pdfData, pdfData.getPageIndices());
    pages.forEach((page) => pdfDoc.addPage(page));
  }

  // Save merged PDF to Drive
  const bytes = await pdfDoc.save();
  const mergedBlob = Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, 'merged.pdf');
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  const mergedFile = destinationFolder.createFile(mergedBlob).setName(name);

  fileIds.forEach((id) => DriveApp.getFileById(id).setTrashed(true));

  return mergedFile;
}

function savePdfSheet(
  spreadsheetId,
  sheetId,
  studentName,
  margin = {
    top: '0.5',
    bottom: '0.5',
    left: '0.3',
    right: '0.3',
  }
) {
  try {
    var spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId();

    var url_base = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export';
    var url_ext =
      '?format=pdf' + //export as pdf
      // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
      (sheetId ? '&gid=' + sheetId : '&id=' + spreadsheetId) +
      // following parameters are optional...
      '&size=letter' + // paper size
      '&portrait=true' + // orientation, false for landscape
      '&fitw=true' + // fit to width, false for actual size
      '&fzr=true' + // repeat row headers (frozen rows) on each page
      '&top_margin=' +
      margin.top +
      '&bottom_margin=' +
      margin.bottom +
      '&left_margin=' +
      margin.left +
      '&right_margin=' +
      margin.right +
      '&printnotes=false' +
      '&sheetnames=false' +
      '&printtitle=false' +
      '&pagenumbers=false'; //hide optional headers and footers

    var options = {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      },
      muteHttpExceptions: true,
    };

    // Create PDF
    const pdfName = spreadsheet.getSheetById(sheetId).getName() + ' sheet for ' + studentName;
    const response = UrlFetchApp.fetch(url_base + url_ext, options);
    const blob = response.getBlob().setName(pdfName + '.pdf');
    const rootFolder = DriveApp.getRootFolder();
    const pdfSheet = rootFolder.createFile(blob);

    return pdfSheet.getId();
  } catch (err) {
    Logger.log(err.stack);
    throw new Error(err.message + '\n\n' + err.stack);
  }
}

function isValidPdf(fileId) {
  const blob = DriveApp.getFileById(fileId).getBlob();
  if (blob.getContentType() !== MimeType.PDF) return false;
  const bytes = blob.getBytes();
  const header = String.fromCharCode.apply(null, bytes.slice(0, 5));
  return header === '%PDF-';
}