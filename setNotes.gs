/**
 * Daily note update for CH-Incoming and SP-Incoming.
 *
 * For each sheet:
 * - Clears notes from AT4:AU
 * - Reads [ECS] Location from BO4:BO
 * - If BO has a value, sets that value as a note on AT and AU of the same row
 *
 * Row 3 is never touched.
 */
function updateIncomingEcsLocationNotes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetNames = [
    'CH-Incoming',
    'SP-Incoming'
  ];

  const startRow = 4;
  const noteStartCol = 46; // AT
  const noteNumCols = 2;   // AT:AU
  const locationCol = 57;  // BE

  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`Sheet not found: ${sheetName}`);
      return;
    }

    const maxRows = sheet.getMaxRows();
    const numRows = maxRows - startRow + 1;

    if (numRows <= 0) {
      Logger.log(`No rows to process on ${sheetName}`);
      return;
    }

    Logger.log(`Processing ${sheetName}`);

    // Clear existing notes in AT4:AU, preserving row 3.
    const noteRange = sheet.getRange(startRow, noteStartCol, numRows, noteNumCols);
    noteRange.clearNote();

    // Read BO4:BO.
    const locationValues = sheet
      .getRange(startRow, locationCol, numRows, 1)
      .getValues();

    // Build a 2-column notes array for AT:AU.
    const notes = locationValues.map(row => {
      const location = row[0];

      if (location !== null && location !== '') {
        const noteText = String(location).trim();

        if (noteText !== '') {
          return [noteText, noteText];
        }
      }

      return ['', ''];
    });

    // Set notes in one batch operation.
    noteRange.setNotes(notes);

    const noteCount = notes.filter(row => row[0] !== '').length;
    Logger.log(`Finished ${sheetName}. Notes added to ${noteCount} rows.`);
  });

  Logger.log('Finished updateIncomingEcsLocationNotes');
}

/**
 * Adds notes to AA from BM on CH-Incoming and SP-Incoming.
 *
 * For each sheet:
 * - Only processes rows where column A is not blank
 * - Clears notes from AA4:AA
 * - Reads values from BM
 * - If BM has a value, sets that value as a note on AA of the same row
 *
 * Row 3 is never touched.
 */
function updateIncomingBmNotesToAa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetNames = [
    'CH-Incoming',
    'SP-Incoming'
  ];

  const startRow = 4;
  const keyCol = 1;       // A
  const noteCol = 27;     // AA
  const sourceCol = 65;   // BM

  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`Sheet not found: ${sheetName}`);
      return;
    }

    const maxRows = sheet.getMaxRows();
    const numRows = maxRows - startRow + 1;

    if (numRows <= 0) {
      Logger.log(`No rows to process on ${sheetName}`);
      return;
    }

    Logger.log(`Processing ${sheetName}`);

    const keyValues = sheet
      .getRange(startRow, keyCol, numRows, 1)
      .getValues();

    const sourceValues = sheet
      .getRange(startRow, sourceCol, numRows, 1)
      .getValues();

    const noteRange = sheet.getRange(startRow, noteCol, numRows, 1);

    // Clear existing notes in AA4:AA, preserving row 3.
    noteRange.clearNote();

    const notes = sourceValues.map((row, index) => {
      const keyValue = keyValues[index][0];
      const sourceValue = row[0];

      const hasKey = keyValue !== null && String(keyValue).trim() !== '';
      const hasSource = sourceValue !== null && String(sourceValue).trim() !== '';

      if (hasKey && hasSource) {
        return [String(sourceValue).trim()];
      }

      return [''];
    });

    noteRange.setNotes(notes);

    const noteCount = notes.filter(row => row[0] !== '').length;
    Logger.log(`Finished ${sheetName}. Notes added to ${noteCount} rows.`);
  });

  Logger.log('Finished updateIncomingBmNotesToAa');
}

function runDailyIncomingNoteUpdates() {
  updateIncomingEcsLocationNotes();
  updateIncomingBmNotesToAa();
}

/**
 * Run once manually to create the daily trigger.
 * Adjust the hour if needed.
 */
function installDailyIncomingNotesTrigger() {
  const functionName = 'runDailyIncomingNoteUpdates';

  // Remove existing triggers for this function to avoid duplicates.
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  Logger.log(`Daily trigger installed for ${functionName}`);
}