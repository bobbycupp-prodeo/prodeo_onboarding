/***************************************
 * recordsRequests.gs
 ***************************************/

const RECORDS_REQUESTS_RESET_CONFIG = {
  sheetNamePrefix: 'RecReq',

  // Was B46 before row 10 was deleted
  checkboxA1: 'B45',

  instructionColumn: 4,
  targetColumn: 2,

  // Row 10 was deleted.
  // Original B10 is gone.
  // Rows below old row 10 moved up by 1.
  rowsToReset: [4, 5, 6, 9, 10, 11, 26, 27, 28, 30, 31, 32],

  clearInstructionText: 'Clear Value',
};

const RECORDS_REQUESTS_PDF_CONFIG = {
  sheetNamePrefix: 'RecReq',

  // Was B47 before row 10 was deleted
  checkboxA1: 'B46',

  requiredCellA1: 'B5',
  filenameCellA1: 'C1',
  destinationFolderId: '1eOAv5hCciCS5RiEfRyXC2K3HiN3bbBZb',
};

/***************************************
 * Logging helper
 ***************************************/

function RecordsRequests_log_(message, data) {
  const timestamp = new Date().toISOString();

  if (data === undefined) {
    Logger.log('[RecordsRequests] %s | %s', timestamp, message);
    return;
  }

  Logger.log(
    '[RecordsRequests] %s | %s | %s',
    timestamp,
    message,
    JSON.stringify(data)
  );
}


/***************************************
 * Main Records Requests edit dispatcher
 ***************************************/

function RecordsRequests_onEdit(e) {
  RecordsRequests_log_('onEdit started');

  if (!e || !e.range) {
    RecordsRequests_log_('onEdit exited: missing event or range');
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();

  RecordsRequests_log_('Edit detected', {
    sheetName: sheet.getName(),
    rangeA1: range.getA1Notation(),
    value: e.value,
    oldValue: e.oldValue,
  });

  RecordsRequests_handleRecReqReset_(e);
  RecordsRequests_handleRecReqPdfExport_(e);

  RecordsRequests_log_('onEdit completed');
}


/***************************************
 * Reset checkbox functions
 ***************************************/

function RecordsRequests_handleRecReqReset_(e) {
  if (!e || !e.range) {
    RecordsRequests_log_('Reset handler exited: missing event or range');
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();

  RecordsRequests_log_('Reset handler checking edit', {
    sheetName: sheet.getName(),
    rangeA1: range.getA1Notation(),
    value: e.value,
  });

  if (!RecordsRequests_shouldHandleRecReqReset_(sheet, range, e.value)) {
    RecordsRequests_log_('Reset handler skipped');
    return;
  }

  RecordsRequests_log_('Reset checkbox triggered', {
    sheetName: sheet.getName(),
  });

  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Confirm Reset',
    'This will reset the selected fields on this RecReq sheet. Continue?',
    ui.ButtonSet.YES_NO
  );

  RecordsRequests_log_('Reset confirmation response received', {
    response: response.toString(),
  });

  if (response !== ui.Button.YES) {
    RecordsRequests_log_('Reset canceled by user');
    range.setValue(false);
    return;
  }

  RecordsRequests_applyRecReqReset_(sheet);

  range.setValue(false);

  RecordsRequests_log_('Reset completed and checkbox unchecked', {
    sheetName: sheet.getName(),
  });
}


function RecordsRequests_shouldHandleRecReqReset_(sheet, range, editedValue) {
  const config = RECORDS_REQUESTS_RESET_CONFIG;

  if (!sheet.getName().startsWith(config.sheetNamePrefix)) {
    RecordsRequests_log_('Reset check failed: sheet prefix mismatch', {
      sheetName: sheet.getName(),
      expectedPrefix: config.sheetNamePrefix,
    });
    return false;
  }

  if (range.getA1Notation() !== config.checkboxA1) {
    RecordsRequests_log_('Reset check failed: edited cell mismatch', {
      editedCell: range.getA1Notation(),
      expectedCell: config.checkboxA1,
    });
    return false;
  }

  if (editedValue !== 'TRUE') {
    RecordsRequests_log_('Reset check failed: checkbox not TRUE', {
      editedValue: editedValue,
    });
    return false;
  }

  RecordsRequests_log_('Reset check passed');
  return true;
}


function RecordsRequests_applyRecReqReset_(sheet) {
  const config = RECORDS_REQUESTS_RESET_CONFIG;

  RecordsRequests_log_('Applying reset instructions', {
    sheetName: sheet.getName(),
    rowsToReset: config.rowsToReset,
  });

  config.rowsToReset.forEach(function(rowNumber) {
    const targetCell = sheet.getRange(rowNumber, config.targetColumn);
    const instructionCell = sheet.getRange(rowNumber, config.instructionColumn);

    const instructionFormula = instructionCell.getFormula();
    const instructionValue = instructionCell.getValue();

    RecordsRequests_log_('Processing reset row', {
      rowNumber: rowNumber,
      targetCell: targetCell.getA1Notation(),
      instructionCell: instructionCell.getA1Notation(),
      instructionValue: instructionValue,
      hasInstructionFormula: Boolean(instructionFormula),
    });

    if (instructionValue === config.clearInstructionText) {
      targetCell.clearContent();

      RecordsRequests_log_('Target cell cleared', {
        targetCell: targetCell.getA1Notation(),
      });

      return;
    }

    if (instructionFormula) {
      targetCell.setFormula(instructionFormula);

      RecordsRequests_log_('Formula applied to target cell', {
        targetCell: targetCell.getA1Notation(),
        formula: instructionFormula,
      });

      return;
    }

    targetCell.setValue(instructionValue);

    RecordsRequests_log_('Value applied to target cell', {
      targetCell: targetCell.getA1Notation(),
      value: instructionValue,
    });
  });

  RecordsRequests_log_('Reset instructions completed', {
    sheetName: sheet.getName(),
  });
}


/***************************************
 * PDF export checkbox functions
 ***************************************/

function RecordsRequests_handleRecReqPdfExport_(e) {
  if (!e || !e.range) {
    RecordsRequests_log_('PDF handler exited: missing event or range');
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();

  RecordsRequests_log_('PDF handler checking edit', {
    sheetName: sheet.getName(),
    rangeA1: range.getA1Notation(),
    value: e.value,
  });

  if (!RecordsRequests_shouldHandleRecReqPdfExport_(sheet, range, e.value)) {
    RecordsRequests_log_('PDF handler skipped');
    return;
  }

  const spreadsheet = sheet.getParent();
  const config = RECORDS_REQUESTS_PDF_CONFIG;

  RecordsRequests_log_('PDF export triggered', {
    sheetName: sheet.getName(),
    spreadsheetId: spreadsheet.getId(),
  });

  try {
    const requiredValue = sheet.getRange(config.requiredCellA1).getValue();

    RecordsRequests_log_('PDF required cell checked', {
      requiredCell: config.requiredCellA1,
      requiredValue: requiredValue,
    });

    if (requiredValue === '' || requiredValue === null) {
      spreadsheet.toast(
        'PDF export skipped because B5 is blank.',
        'Records Request PDF',
        5
      );

      RecordsRequests_log_('PDF export skipped: required cell blank');
      range.setValue(false);
      return;
    }

    const rawFileName = sheet.getRange(config.filenameCellA1).getDisplayValue();
    const pdfName = RecordsRequests_sanitizePdfFileName_(rawFileName);

    RecordsRequests_log_('PDF filename prepared', {
      filenameCell: config.filenameCellA1,
      rawFileName: rawFileName,
      sanitizedPdfName: pdfName,
    });

    if (!pdfName) {
      spreadsheet.toast(
        'PDF export skipped because C1 does not contain a valid filename.',
        'Records Request PDF',
        5
      );

      RecordsRequests_log_('PDF export skipped: invalid filename');
      range.setValue(false);
      return;
    }

    spreadsheet.toast(
      'Preparing PDF export...',
      'Records Request PDF',
      5
    );

    RecordsRequests_log_('Calling PDF export function', {
      sheetName: sheet.getName(),
      pdfName: pdfName,
      destinationFolderId: config.destinationFolderId,
    });

    const pdfFile = RecordsRequests_exportRecReqSheetAsPdf_(
      sheet,
      pdfName,
      config.destinationFolderId
    );

    RecordsRequests_log_('PDF file created', {
      fileName: pdfFile.getName(),
      fileUrl: pdfFile.getUrl(),
      fileId: pdfFile.getId(),
    });

    spreadsheet.toast(
      'PDF saved: ' + pdfFile.getName(),
      'Records Request PDF',
      5
    );

    range.setValue(false);

    RecordsRequests_log_('PDF checkbox unchecked');

    RecordsRequests_showPdfExportCompleteAlert_(
      pdfFile,
      config.destinationFolderId
    );

    RecordsRequests_log_('PDF completion dialog shown');

  } catch (error) {
    range.setValue(false);

    spreadsheet.toast(
      'PDF export failed. See Apps Script logs for details.',
      'Records Request PDF',
      8
    );

    RecordsRequests_log_('PDF export failed', {
      errorName: error.name,
      errorMessage: error.message,
      errorStack: error.stack,
    });

    throw error;
  }
}


function RecordsRequests_shouldHandleRecReqPdfExport_(sheet, range, editedValue) {
  const config = RECORDS_REQUESTS_PDF_CONFIG;

  if (!sheet.getName().startsWith(config.sheetNamePrefix)) {
    RecordsRequests_log_('PDF check failed: sheet prefix mismatch', {
      sheetName: sheet.getName(),
      expectedPrefix: config.sheetNamePrefix,
    });
    return false;
  }

  if (range.getA1Notation() !== config.checkboxA1) {
    RecordsRequests_log_('PDF check failed: edited cell mismatch', {
      editedCell: range.getA1Notation(),
      expectedCell: config.checkboxA1,
    });
    return false;
  }

  if (editedValue !== 'TRUE') {
    RecordsRequests_log_('PDF check failed: checkbox not TRUE', {
      editedValue: editedValue,
    });
    return false;
  }

  RecordsRequests_log_('PDF check passed');
  return true;
}


function RecordsRequests_exportRecReqSheetAsPdf_(sheet, pdfName, folderId) {
  const spreadsheet = sheet.getParent();
  const spreadsheetId = spreadsheet.getId();
  const sheetId = sheet.getSheetId();

  RecordsRequests_log_('Building PDF export URL', {
    spreadsheetId: spreadsheetId,
    sheetId: sheetId,
    sheetName: sheet.getName(),
  });

  const exportUrl =
    'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' +
    [
      'format=pdf',
      'gid=' + sheetId,

      // Page setup
      'size=letter',
      'portrait=true',

      // Export only through row 44
      'range=A1:B43',

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

  RecordsRequests_log_('Fetching PDF export URL', {
    exportUrl: exportUrl,
  });

  const response = UrlFetchApp.fetch(exportUrl, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
    muteHttpExceptions: true,
  });

  const responseCode = response.getResponseCode();

  RecordsRequests_log_('PDF export response received', {
    responseCode: responseCode,
    contentType: response.getHeaders()['Content-Type'],
  });

  if (responseCode !== 200) {
    RecordsRequests_log_('PDF export returned non-200 response', {
      responseCode: responseCode,
      responseText: response.getContentText(),
    });

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

  RecordsRequests_log_('Creating PDF file in Drive folder', {
    folderId: folderId,
    pdfBlobName: pdfBlob.getName(),
  });

  const file = DriveApp.getFolderById(folderId).createFile(pdfBlob);

  RecordsRequests_log_('Drive file created', {
    fileId: file.getId(),
    fileName: file.getName(),
    fileUrl: file.getUrl(),
  });

  return file;
}


function RecordsRequests_sanitizePdfFileName_(rawFileName) {
  if (!rawFileName) {
    RecordsRequests_log_('Filename sanitization received empty value');
    return '';
  }

  const sanitizedFileName = rawFileName
    .toString()
    .trim()
    .replace(/[\\/:*?"<>|#%{}~&]/g, '-')
    .replace(/\s+/g, ' ');

  RecordsRequests_log_('Filename sanitized', {
    rawFileName: rawFileName,
    sanitizedFileName: sanitizedFileName,
  });

  return sanitizedFileName;
}


function RecordsRequests_showPdfExportCompleteAlert_(pdfFile, folderId) {
  const folderUrl = 'https://drive.google.com/drive/folders/' + folderId;
  const pdfUrl = pdfFile.getUrl();

  // Direct download/export URL for the created PDF file.
  const downloadUrl = 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId();

  RecordsRequests_log_('Showing PDF completion alert', {
    pdfUrl: pdfUrl,
    downloadUrl: downloadUrl,
    folderUrl: folderUrl,
  });

  const html = HtmlService
    .createHtmlOutput(
      '<p>The PDF has been created successfully.</p>' +
      '<p><a href="' + pdfUrl + '" target="_blank">Open PDF</a></p>' +
      '<p><a href="' + downloadUrl + '" target="_blank">Download PDF</a></p>' +
      '<p><a href="' + folderUrl + '" target="_blank">Open destination folder</a></p>'
    )
    .setWidth(380)
    .setHeight(210);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Records Request PDF Created'
  );
}


/***************************************
 * Trigger setup
 ***************************************/

function RecordsRequests_createInstallableOnEditTrigger() {
  const ss = SpreadsheetApp.getActive();

  RecordsRequests_log_('Refreshing installable onEdit trigger');

  const recordsRequestTriggerFunctions = [
    'RecordsRequests_onEdit',

    // Old trigger function names to clean up
    'RecReq_Reset_onEdit',
  ];

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    const handlerFunction = trigger.getHandlerFunction();

    if (recordsRequestTriggerFunctions.indexOf(handlerFunction) !== -1) {
      RecordsRequests_log_('Deleting existing Records Requests trigger', {
        triggerUniqueId: trigger.getUniqueId(),
        handlerFunction: handlerFunction,
      });

      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('RecordsRequests_onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  RecordsRequests_log_('Installable onEdit trigger created', {
    spreadsheetId: ss.getId(),
    handlerFunction: 'RecordsRequests_onEdit',
  });
}