/***************************************
 * dataUpdates.gs
 *
 * Adds new SM Student IDs from sm_reg to
 * the first available blank rows in 
 * CH-Incoming!A:A or SP-Incoming!A:A based on Campus.
 *
 * Intended for time-driven trigger use.
 ***************************************/

const DATA_UPDATES_CONFIG = {
  sourceSheetName: 'sm_reg',
  targetSheetSuffix: '-Incoming', // Used to build 'CH-Incoming', 'SP-Incoming'
  sourceHeaderRow: 2,
  targetHeaderRow: 2,
  targetStartRow: 3,
  sourceIdHeader: 'SM Student ID',
  sourceCampusHeader: 'Campus', // New config to locate the Campus column
  targetIdColumn: 1 // Column A
};

function dataUpdates_addNewStudentIdsToIncoming() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(DATA_UPDATES_CONFIG.sourceSheetName);

  if (!sourceSheet) {
    Logger.log(`Source sheet not found: ${DATA_UPDATES_CONFIG.sourceSheetName}`);
    return;
  }

  Logger.log('Starting student ID update.');

  const sourceIdColumn = dataUpdates_getColumnByHeader_(
    sourceSheet,
    DATA_UPDATES_CONFIG.sourceHeaderRow,
    DATA_UPDATES_CONFIG.sourceIdHeader
  );

  const sourceCampusColumn = dataUpdates_getColumnByHeader_(
    sourceSheet,
    DATA_UPDATES_CONFIG.sourceHeaderRow,
    DATA_UPDATES_CONFIG.sourceCampusHeader
  );

  if (!sourceIdColumn || !sourceCampusColumn) {
    Logger.log(`Required headers not found in ${DATA_UPDATES_CONFIG.sourceSheetName}.`);
    return;
  }

  const sourceRecords = dataUpdates_getSourceRecords_(
    sourceSheet,
    DATA_UPDATES_CONFIG.sourceHeaderRow + 1,
    sourceIdColumn,
    sourceCampusColumn
  );

  if (sourceRecords.length === 0) {
    Logger.log('No valid source records found.');
    return;
  }

  // Group the retrieved IDs by campus
  const recordsByCampus = {};
  sourceRecords.forEach(record => {
    const campus = record.campus;
    if (!recordsByCampus[campus]) {
      recordsByCampus[campus] = [];
    }
    recordsByCampus[campus].push(record.id);
  });

  // Process and write new IDs for each campus individually
  for (const campus in recordsByCampus) {
    const targetSheetName = `${campus}${DATA_UPDATES_CONFIG.targetSheetSuffix}`;
    const targetSheet = ss.getSheetByName(targetSheetName);
    
    if (!targetSheet) {
      Logger.log(`Target sheet not found: ${targetSheetName}. Skipping campus ${campus}.`);
      continue;
    }

    const campusIds = recordsByCampus[campus];
    Logger.log(`Processing ${campusIds.length} unique IDs for ${targetSheetName}.`);

    const existingTargetIds = dataUpdates_getExistingTargetIds_(
      targetSheet,
      DATA_UPDATES_CONFIG.targetStartRow,
      DATA_UPDATES_CONFIG.targetIdColumn
    );

    const newIds = campusIds.filter(id => !existingTargetIds.has(id));

    if (newIds.length === 0) {
      Logger.log(`No new IDs to add for ${targetSheetName}.`);
      continue;
    }

    Logger.log(`Found ${newIds.length} new IDs to add to ${targetSheetName}.`);

    const firstEmptyRow = dataUpdates_findFirstEmptyCellInColumn_(
      targetSheet,
      DATA_UPDATES_CONFIG.targetStartRow,
      DATA_UPDATES_CONFIG.targetIdColumn
    );

    const today = new Date();
    const valuesToWrite = newIds.map(id => [id, today]);

    // Write the new IDs to the target sheet
    targetSheet
      .getRange(firstEmptyRow, DATA_UPDATES_CONFIG.targetIdColumn, valuesToWrite.length, 2)
      .setValues(valuesToWrite);

    // Format the date column
    targetSheet
      .getRange(firstEmptyRow, 2, valuesToWrite.length, 1)
      .setNumberFormat('M/d/yyyy');

    Logger.log(`Added ${newIds.length} new IDs to ${targetSheetName}.`);
  }

  Logger.log('Student ID update complete.');
}

/**
 * Finds a column number by matching a header value in a specific header row.
 */
function dataUpdates_getColumnByHeader_(sheet, headerRow, headerName) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
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
 * Gets unique, nonblank records (ID and Campus) from the source sheet.
 * Skips students marked "Withdrawn" in column G.
 */
function dataUpdates_getSourceRecords_(sheet, startRow, idColumn, campusColumn) {
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) {
    return [];
  }

  const numRows = lastRow - startRow + 1;
  const statusColumn = 7; // Column G

  const idValues = sheet.getRange(startRow, idColumn, numRows, 1).getValues().flat();
  const statusValues = sheet.getRange(startRow, statusColumn, numRows, 1).getValues().flat();
  const campusValues = sheet.getRange(startRow, campusColumn, numRows, 1).getValues().flat();

  const seen = new Set();
  const records = [];

  idValues.forEach((value, index) => {
    const id = String(value).trim();
    const status = String(statusValues[index]).trim().toLowerCase();
    const campus = String(campusValues[index]).trim().toUpperCase();

    if (!id || !campus) {
      return;
    }

    if (status === 'withdrawn') {
      Logger.log(`Skipping withdrawn student ID: ${id}`);
      return;
    }

    if (!seen.has(id)) {
      seen.add(id);
      records.push({ id: id, campus: campus });
    }
  });

  return records;
}

/**
 * Gets existing IDs from the incoming sheet's column A.
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