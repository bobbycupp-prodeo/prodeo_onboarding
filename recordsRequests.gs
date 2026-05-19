/**
 * RecReq Reset Script
 *
 * Purpose:
 * When B46 is checked on any sheet whose name starts with "RecReq",
 * confirm with the user, then reset selected Column B cells based on
 * the corresponding instructions in Column D.
 *
 * Instructions:
 * - If D[row] contains exactly "Clear Value", B[row] is cleared.
 * - If D[row] contains a formula, that exact formula is placed in B[row].
 * - Otherwise, B[row] is set to the raw value from D[row].
 */

const REC_REQ_RESET_CONFIG = {
  sheetNamePrefix: 'RecReq',
  checkboxA1: 'B46',
  instructionColumn: 4, // Column D
  targetColumn: 2,      // Column B
  rowsToReset: [4, 5, 6, 9, 10, 11, 12, 27, 28, 29, 31, 32, 33],
  clearInstructionText: 'Clear Value',
};

/**
 * Installable on-edit handler for RecReq reset behavior.
 *
 * Create an installable on-edit trigger pointing to this function.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function RecReq_Reset_onEdit(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (!RecReq_Reset_shouldHandleEdit_(sheet, range, e.value)) return;

  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Confirm Reset',
    'This will reset the selected fields on this RecReq sheet. Continue?',
    ui.ButtonSet.YES_NO
  );

  

  if (response !== ui.Button.YES) {
    range.setValue(false);
    return;
  }

  RecReq_Reset_applyReset_(sheet);

  // Uncheck the reset checkbox after the reset completes.
  // Script-made edits do not cause edit triggers to run again. 
  range.setValue(false);
}

/**
 * Determines whether this edit should trigger the RecReq reset logic.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {string|undefined} editedValue
 * @returns {boolean}
 */
function RecReq_Reset_shouldHandleEdit_(sheet, range, editedValue) {
  const config = REC_REQ_RESET_CONFIG;

  if (!sheet.getName().startsWith(config.sheetNamePrefix)) return false;
  if (range.getA1Notation() !== config.checkboxA1) return false;

  // Checkboxes usually pass "TRUE" / "FALSE" as strings in the event object.
  return editedValue === 'TRUE';
}

/**
 * Applies the reset instructions from Column D to Column B.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function RecReq_Reset_applyReset_(sheet) {
  const config = REC_REQ_RESET_CONFIG;

  config.rowsToReset.forEach(function(rowNumber) {
    const targetCell = sheet.getRange(rowNumber, config.targetColumn);
    const instructionCell = sheet.getRange(rowNumber, config.instructionColumn);

    const instructionFormula = instructionCell.getFormula();
    const instructionValue = instructionCell.getValue();

    if (instructionValue === config.clearInstructionText) {
      targetCell.clearContent();
      return;
    }

    if (instructionFormula) {
      targetCell.setFormula(instructionFormula);
      return;
    }

    targetCell.setValue(instructionValue);
  });
}

/**
 * Optional helper:
 * Run this once manually to create the installable on-edit trigger.
 *
 * This removes any existing RecReq_Reset_onEdit triggers first
 * to avoid duplicate executions.
 */
function RecReq_Reset_createInstallableOnEditTrigger() {
  const ss = SpreadsheetApp.getActive();

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'RecReq_Reset_onEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('RecReq_Reset_onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

/**
 * Exports a single sheet tab as a PDF and saves it to Drive.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} pdfName
 * @param {string=} folderId Optional Drive folder ID. If omitted, saves to My Drive root.
 * @returns {GoogleAppsScript.Drive.File}
 */
function RecordsRequests_exportSheetAsPdf_(sheet, pdfName, folderId) {
  const spreadsheet = sheet.getParent();
  const spreadsheetId = spreadsheet.getId();
  const sheetId = sheet.getSheetId();

  

  const exportUrl =
    'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' +
    [
      'format=pdf',
      'gid=' + sheetId,

      // Page setup
      'size=letter',
      'portrait=true',
      // 'fitw=true',
      'scale=4',

      // Margins
      'top_margin=0.5',
      'bottom_margin=0.5',
      'left_margin=0.5',
      'right_margin=0.5',

      // Display options
      'sheetnames=false',
      'printtitle=false',
      'pagenumbers=false',
      'gridlines=false',
      'fzr=false',

      // Export behavior
      'attachment=false'
    ].join('&');

  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(exportUrl, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    muteHttpExceptions: true,
  });

  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error(
      'PDF export failed. Response code: ' +
      responseCode +
      '. Response: ' +
      response.getContentText()
    );
  }

  const pdfBlob = response
    .getBlob()
    .setName(pdfName.endsWith('.pdf') ? pdfName : pdfName + '.pdf');

  if (folderId) {
    return DriveApp.getFolderById(folderId).createFile(pdfBlob);
  }

  return DriveApp.createFile(pdfBlob);
}

/**
 * Handles the RecReq PDF export checkbox behavior.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function RecordsRequests_handleRecReqPdfExport_(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (!RecordsRequests_shouldHandleRecReqPdfExport_(sheet, range, e.value)) return;

  const spreadsheet = sheet.getParent();
  const config = RECORDS_REQUESTS_PDF_CONFIG;

  try {
    const requiredValue = sheet.getRange(config.requiredCellA1).getValue();

    if (requiredValue === '' || requiredValue === null) {
      spreadsheet.toast(
        'PDF export skipped because B5 is blank.',
        'Records Request PDF',
        5
      );

      range.setValue(false);
      return;
    }

    const rawFileName = sheet.getRange(config.filenameCellA1).getDisplayValue();
    const pdfName = RecordsRequests_sanitizePdfFileName_(rawFileName);

    if (!pdfName) {
      spreadsheet.toast(
        'PDF export skipped because C1 does not contain a valid filename.',
        'Records Request PDF',
        5
      );

      range.setValue(false);
      return;
    }

    spreadsheet.toast(
      'Preparing PDF export...',
      'Records Request PDF',
      5
    );

    const pdfFile = RecordsRequests_exportRecReqSheetAsPdf_(
      sheet,
      pdfName,
      config.destinationFolderId
    );

    spreadsheet.toast(
      'PDF saved: ' + pdfFile.getName(),
      'Records Request PDF',
      5
    );

    range.setValue(false);

    RecordsRequests_showPdfExportCompleteAlert_(
      pdfFile,
      config.destinationFolderId
    );

  } catch (error) {
    range.setValue(false);

    spreadsheet.toast(
      'PDF export failed. See Apps Script logs for details.',
      'Records Request PDF',
      8
    );

    Logger.log(error);
    throw error;
  }
}

/**
 * Determines whether the edit should trigger the RecReq PDF export.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {string|undefined} editedValue
 * @returns {boolean}
 */
function RecordsRequests_shouldHandleRecReqPdfExport_(sheet, range, editedValue) {
  const config = RECORDS_REQUESTS_PDF_CONFIG;

  if (!sheet.getName().startsWith(config.sheetNamePrefix)) return false;
  if (range.getA1Notation() !== config.checkboxA1) return false;

  return editedValue === 'TRUE';
}

/**
 * Exports a RecReq sheet as a PDF and saves it to the provided Drive folder.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} pdfName
 * @param {string} folderId
 * @returns {GoogleAppsScript.Drive.File}
 */
function RecordsRequests_exportRecReqSheetAsPdf_(sheet, pdfName, folderId) {
  const spreadsheet = sheet.getParent();
  const spreadsheetId = spreadsheet.getId();
  const sheetId = sheet.getSheetId();

  const exportUrl =
    'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' +
    [
      'format=pdf',
      'gid=' + sheetId,

      // Page setup
      'size=letter',
      'portrait=true',

      // Closest approximation of "Fit to page"
      'scale=4',

      // Margins
      'top_margin=0.5',
      'bottom_margin=0.5',
      'left_margin=0.5',
      'right_margin=0.5',

      // Display options
      'sheetnames=false',
      'printtitle=false',
      'pagenumbers=false',
      'gridlines=false',
      'fzr=false',

      // Export behavior
      'attachment=false'
    ].join('&');

  const response = UrlFetchApp.fetch(exportUrl, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
    muteHttpExceptions: true,
  });

  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error(
      'PDF export failed. Response code: ' +
      responseCode +
      '. Response: ' +
      response.getContentText()
    );
  }

  const pdfBlob = response
    .getBlob()
    .setName(pdfName.endsWith('.pdf') ? pdfName : pdfName + '.pdf');

  return DriveApp.getFolderById(folderId).createFile(pdfBlob);
}

/**
 * Sanitizes a value for use as a PDF filename.
 *
 * @param {string} rawFileName
 * @returns {string}
 */
function RecordsRequests_sanitizePdfFileName_(rawFileName) {
  if (!rawFileName) return '';

  return rawFileName
    .toString()
    .trim()
    .replace(/[\\/:*?"<>|#%{}~&]/g, '-')
    .replace(/\s+/g, ' ');
}

const RECORDS_REQUESTS_PDF_CONFIG = {
  sheetNamePrefix: 'RecReq',
  checkboxA1: 'B47',
  requiredCellA1: 'B5',
  filenameCellA1: 'C1',
  destinationFolderId: '1eOAv5hCciCS5RiEfRyXC2K3HiN3bbBZb',
};

/**
 * Installable on-edit handler for Records Requests functionality.
 *
 * Point your installable edit trigger to this function.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function RecordsRequests_onEdit(e) {
  RecordsRequests_handleRecReqReset_(e);
  RecordsRequests_handleRecReqPdfExport_(e);

  // Add future Records Requests edit handlers here.
}

/**
 * Shows a completion alert with links to the PDF and destination folder.
 *
 * @param {GoogleAppsScript.Drive.File} pdfFile
 * @param {string} folderId
 */
function RecordsRequests_showPdfExportCompleteAlert_(pdfFile, folderId) {
  const folderUrl = 'https://drive.google.com/drive/folders/' + folderId;

  const html = HtmlService
    .createHtmlOutput(
      '<p>The PDF has been created successfully.</p>' +
      '<p><a href="' + pdfFile.getUrl() + '" target="_blank">Open PDF</a></p>' +
      '<p><a href="' + folderUrl + '" target="_blank">Open destination folder</a></p>'
    )
    .setWidth(360)
    .setHeight(180);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Records Request PDF Created'
  );
}