/***************************************
 * dataUpdates.gs
 *
 * Adds new SM Student IDs from sm_reg to
 * the first available blank rows in
 * CH-Incoming!A:B or SP-Incoming!A:B based on Campus.
 *
 * Also appends the same newly-added students
 * to marss_ids using the configured sm_reg columns.
 *
 * Intended for time-driven trigger use.
 ***************************************/

const DATA_UPDATES_CONFIG = {
  sourceSheetName: 'sm_reg',
  marssSheetName: 'marss_ids',

  targetSheetSuffix: '-Incoming', // Used to build 'CH-Incoming', 'SP-Incoming'

  sourceHeaderRow: 2,
  targetHeaderRow: 2,
  targetStartRow: 4,

  marssStartRow: 3, // Assumes marss_ids has headers in row 1

  sourceIdHeader: 'SM Student ID',
  sourceCampusHeader: 'Campus',

  targetIdColumn: 1, // Column A
  targetDateColumn: 2 // Column B
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

  const marssSheet = ss.getSheetByName(DATA_UPDATES_CONFIG.marssSheetName);

  if (!marssSheet) {
    Logger.log(`MARSS sheet not found: ${DATA_UPDATES_CONFIG.marssSheetName}. Incoming sheets will still be updated.`);
  }

  const existingMarssIds = marssSheet
    ? dataUpdates_getExistingTargetIds_(
        marssSheet,
        DATA_UPDATES_CONFIG.marssStartRow,
        1 // marss_ids column A
      )
    : new Set();

  // Group the retrieved records by campus.
  const recordsByCampus = {};

  sourceRecords.forEach(record => {
    const campus = record.campus;

    if (!recordsByCampus[campus]) {
      recordsByCampus[campus] = [];
    }

    recordsByCampus[campus].push(record);
  });

  // Process and write new IDs for each campus individually.
  for (const campus in recordsByCampus) {
    const targetSheetName = `${campus}${DATA_UPDATES_CONFIG.targetSheetSuffix}`;
    const targetSheet = ss.getSheetByName(targetSheetName);

    if (!targetSheet) {
      Logger.log(`Target sheet not found: ${targetSheetName}. Skipping campus ${campus}.`);
      continue;
    }

    const campusRecords = recordsByCampus[campus];

    Logger.log(`Processing ${campusRecords.length} unique records for ${targetSheetName}.`);

    const existingTargetIds = dataUpdates_getExistingTargetIds_(
      targetSheet,
      DATA_UPDATES_CONFIG.targetStartRow,
      DATA_UPDATES_CONFIG.targetIdColumn
    );

    const newRecords = campusRecords.filter(record => !existingTargetIds.has(record.id));

    if (newRecords.length === 0) {
      Logger.log(`No new IDs to add for ${targetSheetName}.`);
      continue;
    }

    Logger.log(`Found ${newRecords.length} new IDs to add to ${targetSheetName}.`);

    const firstEmptyRow = dataUpdates_findFirstEmptyCellInColumn_(
      targetSheet,
      DATA_UPDATES_CONFIG.targetStartRow,
      DATA_UPDATES_CONFIG.targetIdColumn
    );

    const today = new Date();

    const incomingValuesToWrite = newRecords.map(record => [
      record.id,
      today
    ]);

    targetSheet
      .getRange(
        firstEmptyRow,
        DATA_UPDATES_CONFIG.targetIdColumn,
        incomingValuesToWrite.length,
        2
      )
      .setValues(incomingValuesToWrite);

    targetSheet
      .getRange(
        firstEmptyRow,
        DATA_UPDATES_CONFIG.targetDateColumn,
        incomingValuesToWrite.length,
        1
      )
      .setNumberFormat('M/d/yyyy');

    Logger.log(`Added ${newRecords.length} new IDs to ${targetSheetName}.`);

    if (marssSheet) {
      const marssRecordsToWrite = newRecords.filter(record => !existingMarssIds.has(record.id));

      if (marssRecordsToWrite.length === 0) {
        Logger.log(`No new MARSS IDs to append for ${targetSheetName}.`);
      } else {
        dataUpdates_appendRecordsToMarssIds_(marssSheet, marssRecordsToWrite);

        marssRecordsToWrite.forEach(record => existingMarssIds.add(record.id));

        Logger.log(`Appended ${marssRecordsToWrite.length} records to ${DATA_UPDATES_CONFIG.marssSheetName}.`);
      }
    }
  }

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
 * Gets unique, nonblank records from sm_reg.
 * Skips students marked "Withdrawn" in column G.
 *
 * MARSS mapping:
 * sm_reg B  -> marss_ids A
 * sm_reg N  -> marss_ids B
 * sm_reg M  -> marss_ids C
 * sm_reg O  -> marss_ids D
 * sm_reg P  -> marss_ids E
 * sm_reg Q  -> marss_ids F
 * sm_reg R  -> marss_ids G
 * sm_reg EW -> marss_ids H
 */
function dataUpdates_getSourceRecords_(sheet, startRow, idColumn, campusColumn) {
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) {
    return [];
  }

  const numRows = lastRow - startRow + 1;

  const statusColumn = 7; // Column G

  const marssColumns = {
    smStudentId: dataUpdates_columnLetterToNumber_('B'),
    colN: dataUpdates_columnLetterToNumber_('N'),
    colM: dataUpdates_columnLetterToNumber_('M'),
    colO: dataUpdates_columnLetterToNumber_('O'),
    colP: dataUpdates_columnLetterToNumber_('P'),
    colQ: dataUpdates_columnLetterToNumber_('Q'),
    colR: dataUpdates_columnLetterToNumber_('R'),
    colEW: dataUpdates_columnLetterToNumber_('EW')
  };

  const maxNeededColumn = Math.max(
    idColumn,
    campusColumn,
    statusColumn,
    marssColumns.smStudentId,
    marssColumns.colN,
    marssColumns.colM,
    marssColumns.colO,
    marssColumns.colP,
    marssColumns.colQ,
    marssColumns.colR,
    marssColumns.colEW
  );

  const sourceValues = sheet
    .getRange(startRow, 1, numRows, maxNeededColumn)
    .getValues();

  const seen = new Set();
  const records = [];

  sourceValues.forEach(row => {
    const id = String(row[idColumn - 1]).trim();
    const status = String(row[statusColumn - 1]).trim().toLowerCase();
    const campus = String(row[campusColumn - 1]).trim().toUpperCase();

    if (!id || !campus) {
      return;
    }

    if (status === 'withdrawn') {
      Logger.log(`Skipping withdrawn student ID: ${id}`);
      return;
    }

    if (seen.has(id)) {
      return;
    }

    seen.add(id);

    records.push({
      id: id,
      campus: campus,
      marssValues: [
        row[marssColumns.smStudentId - 1], // sm_reg B  -> marss_ids A
        row[marssColumns.colN - 1],        // sm_reg N  -> marss_ids B
        row[marssColumns.colM - 1],        // sm_reg M  -> marss_ids C
        row[marssColumns.colO - 1],        // sm_reg O  -> marss_ids D
        row[marssColumns.colP - 1],        // sm_reg P  -> marss_ids E
        row[marssColumns.colQ - 1],        // sm_reg Q  -> marss_ids F
        row[marssColumns.colR - 1],        // sm_reg R  -> marss_ids G
        row[marssColumns.colEW - 1]        // sm_reg EW -> marss_ids H
      ]
    });
  });

  return records;
}

/**
 * Appends records to marss_ids columns A:H.
 */
function dataUpdates_appendRecordsToMarssIds_(marssSheet, records) {
  if (records.length === 0) {
    return;
  }

  const firstEmptyRow = dataUpdates_findFirstEmptyCellInColumn_(
    marssSheet,
    DATA_UPDATES_CONFIG.marssStartRow,
    1 // Column A
  );

  const valuesToWrite = records.map(record => record.marssValues);

  marssSheet
    .getRange(firstEmptyRow, 1, valuesToWrite.length, 8)
    .setValues(valuesToWrite);
}

/**
 * Gets existing IDs from a sheet's specified ID column.
 *
 * This intentionally scans the specified column starting at startRow
 * instead of relying on getLastRow().
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

/**
 * Converts a spreadsheet column letter to a 1-based column number.
 * Example: A -> 1, B -> 2, EW -> 153.
 */
function dataUpdates_columnLetterToNumber_(columnLetter) {
  let columnNumber = 0;
  const letters = String(columnLetter).trim().toUpperCase();

  for (let i = 0; i < letters.length; i++) {
    columnNumber = columnNumber * 26 + (letters.charCodeAt(i) - 64);
  }

  return columnNumber;
}