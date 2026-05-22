/***************************************
 * dataUpdates.gs
 *
 * Adds new SM Student IDs from sm_reg to
 * the first available blank rows in Incoming!A:A.
 *
 * Intended for time-driven trigger use.
 ***************************************/

const DATA_UPDATES_CONFIG = {
  sourceSheetName: 'sm_reg',
  targetSheetName: 'Incoming',
  sourceHeaderRow: 2,
  targetHeaderRow: 2,
  targetStartRow: 3,
  sourceIdHeader: 'SM Student ID',
  targetIdColumn: 1 // Column A
};

function dataUpdates_addNewStudentIdsToIncoming() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceSheet = ss.getSheetByName(DATA_UPDATES_CONFIG.sourceSheetName);
  const targetSheet = ss.getSheetByName(DATA_UPDATES_CONFIG.targetSheetName);

  if (!sourceSheet) {
    Logger.log(`Source sheet not found: ${DATA_UPDATES_CONFIG.sourceSheetName}`);
    return;
  }

  if (!targetSheet) {
    Logger.log(`Target sheet not found: ${DATA_UPDATES_CONFIG.targetSheetName}`);
    return;
  }

  Logger.log('Starting student ID update.');

  const sourceIdColumn = dataUpdates_getColumnByHeader_(
    sourceSheet,
    DATA_UPDATES_CONFIG.sourceHeaderRow,
    DATA_UPDATES_CONFIG.sourceIdHeader
  );

  if (!sourceIdColumn) {
    Logger.log(`Header not found in ${DATA_UPDATES_CONFIG.sourceSheetName}: ${DATA_UPDATES_CONFIG.sourceIdHeader}`);
    return;
  }

  const sourceIds = dataUpdates_getSourceIds_(
    sourceSheet,
    DATA_UPDATES_CONFIG.sourceHeaderRow + 1,
    sourceIdColumn
  );

  if (sourceIds.length === 0) {
    Logger.log('No source IDs found.');
    return;
  }

  Logger.log(`Found ${sourceIds.length} unique IDs in ${DATA_UPDATES_CONFIG.sourceSheetName}.`);

  const existingTargetIds = dataUpdates_getExistingTargetIds_(
    targetSheet,
    DATA_UPDATES_CONFIG.targetStartRow,
    DATA_UPDATES_CONFIG.targetIdColumn
  );

  Logger.log(`Found ${existingTargetIds.size} existing IDs in ${DATA_UPDATES_CONFIG.targetSheetName}.`);

  const newIds = sourceIds.filter(id => !existingTargetIds.has(id));

  if (newIds.length === 0) {
    Logger.log('No new IDs to add.');
    return;
  }

  Logger.log(`Found ${newIds.length} new IDs to add.`);

  const firstEmptyRow = dataUpdates_findFirstEmptyCellInColumn_(
    targetSheet,
    DATA_UPDATES_CONFIG.targetStartRow,
    DATA_UPDATES_CONFIG.targetIdColumn
  );

  Logger.log(`First empty row in Incoming column A is row ${firstEmptyRow}.`);

  const today = new Date();

  const valuesToWrite = newIds.map(id => [id, today]);

  targetSheet
    .getRange(firstEmptyRow, DATA_UPDATES_CONFIG.targetIdColumn, valuesToWrite.length, 2)
    .setValues(valuesToWrite);

  targetSheet
    .getRange(firstEmptyRow, 2, valuesToWrite.length, 1)
    .setNumberFormat('M/d/yyyy');

  Logger.log(`Added ${newIds.length} new IDs to ${DATA_UPDATES_CONFIG.targetSheetName}.`);
  Logger.log('Student ID update complete.');
}

/**
 * Finds a column number by matching a header value in a specific header row.
 */
function dataUpdates_getColumnByHeader_(sheet, headerRow, headerName) {
  const lastColumn = sheet.getLastColumn();

  const headers = sheet
    .getRange(headerRow, 1, 1, lastColumn)
    .getValues()[0];

  const normalizedHeaderName = String(headerName).trim().toLowerCase();

  for (let i = 0; i < headers.length; i++) {
    const currentHeader = String(headers[i]).trim().toLowerCase();

    if (currentHeader === normalizedHeaderName) {
      return i + 1;
    }
  }

  return null;
}


/**
 * Gets unique, nonblank IDs from the source ID column.
 * Skips students marked "Withdrawn" in column G.
 *
 * getLastRow() is acceptable here because sm_reg is the source data sheet.
 */
function dataUpdates_getSourceIds_(sheet, startRow, idColumn) {
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) {
    return [];
  }

  const numRows = lastRow - startRow + 1;
  const statusColumn = 7; // Column G

  const idValues = sheet
    .getRange(startRow, idColumn, numRows, 1)
    .getValues()
    .flat();

  const statusValues = sheet
    .getRange(startRow, statusColumn, numRows, 1)
    .getValues()
    .flat();

  const seen = new Set();
  const ids = [];

  idValues.forEach((value, index) => {
    const id = String(value).trim();
    const status = String(statusValues[index]).trim().toLowerCase();

    if (!id) {
      return;
    }

    if (status === 'withdrawn') {
      Logger.log(`Skipping withdrawn student ID: ${id}`);
      return;
    }

    if (!seen.has(id)) {
      seen.add(id);
      ids.push(id);
    }
  });

  return ids;
}

/**
 * Gets existing IDs from Incoming column A.
 *
 * This intentionally scans column A starting at row 3 and does not rely on
 * getLastRow(), because Incoming has formulas in many rows/columns.
 */
function dataUpdates_getExistingTargetIds_(sheet, startRow, idColumn) {
  const maxRows = sheet.getMaxRows();

  if (maxRows < startRow) {
    return new Set();
  }

  const values = sheet
    .getRange(startRow, idColumn, maxRows - startRow + 1, 1)
    .getValues()
    .flat();

  const ids = new Set();

  values.forEach(value => {
    const id = String(value).trim();

    if (id) {
      ids.add(id);
    }
  });

  return ids;
}

/**
 * Finds the first blank cell in a specific column, starting at startRow.
 *
 * This does not use getLastRow().
 */
function dataUpdates_findFirstEmptyCellInColumn_(sheet, startRow, column) {
  const maxRows = sheet.getMaxRows();

  if (maxRows < startRow) {
    sheet.insertRowsAfter(maxRows, startRow - maxRows);
  }

  const refreshedMaxRows = sheet.getMaxRows();

  const values = sheet
    .getRange(startRow, column, refreshedMaxRows - startRow + 1, 1)
    .getValues()
    .flat();

  for (let i = 0; i < values.length; i++) {
    const value = String(values[i]).trim();

    if (!value) {
      return startRow + i;
    }
  }

  sheet.insertRowsAfter(refreshedMaxRows, 100);

  return refreshedMaxRows + 1;
}