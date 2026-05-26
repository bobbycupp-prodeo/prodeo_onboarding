/****************************************************
 * CONFIGURATION
 ****************************************************/

const MANAGEMENT_SHEET_NAME = "sheet_slicer_management";

// Sheet management columns
const SHEET_NAME_COL = 1;  // A
const SHEET_ORDER_COL = 2; // B
const SHEET_HIDE_COL = 3;  // C
const SHEET_COLOR_COL = 4; // D

// Slicer management columns
const SLICER_NAME_COL = 9; // I
const SLICER_SHEET_COL = 10; // J
const SLICER_TARGET_CELL_COL = 11; // K

const HEADER_ROW = 1;
const START_ROW = 2;


/****************************************************
 * OPTIONAL MENU
 ****************************************************/

// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu("Sheet/Slicer Management")
//     .addItem("List Sheets", "listSheetsForManagement")
//     .addItem("Apply Sheet Order/Hide Settings", "applySheetManagement")
//     .addSeparator()
//     .addItem("Inventory Slicers", "inventoryAllSlicers")
//     .addItem("Align Slicers", "alignSlicersFromConfig")
//     .addToUi();
// }


/****************************************************
 * SHEET MANAGEMENT
 ****************************************************/

/**
 * Run once, or whenever you want to refresh the sheet list.
 *
 * Lists all sheets in the spreadsheet into:
 * Column A: Sheet Name
 * Column B: Current sheet order
 * Column C: Hide checkbox
 * Column D: Sheet tab color swatch
 *
 * Column D uses the cell background color as the control.
 * The cell text is intentionally left blank.
 */
function listSheetsForManagement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const mgmtSheet = ss.getSheetByName(MANAGEMENT_SHEET_NAME);

  if (!mgmtSheet) {
    ui.alert(
      "Missing Management Sheet",
      `Could not find a sheet named "${MANAGEMENT_SHEET_NAME}". Please create it first.`,
      ui.ButtonSet.OK
    );
    return;
  }

  const sheets = ss.getSheets();

  // Clear previous sheet-management data in A:D, preserving row 1 headers.
  const lastRow = Math.max(mgmtSheet.getLastRow(), START_ROW);

  mgmtSheet
    .getRange(START_ROW, SHEET_NAME_COL, lastRow - START_ROW + 1, 4)
    .clearContent()
    .clearDataValidations()
    .setBackground(null);

  const outputData = sheets.map((sheet, index) => [
    sheet.getName(),
    index + 1,
    sheet.isSheetHidden(),
    ""
  ]);

  if (outputData.length === 0) return;

  mgmtSheet
    .getRange(START_ROW, SHEET_NAME_COL, outputData.length, 4)
    .setValues(outputData);

  // Add checkboxes to Hide column.
  mgmtSheet
    .getRange(START_ROW, SHEET_HIDE_COL, outputData.length, 1)
    .insertCheckboxes();

  // Set Column D background colors to match sheet tab colors.
  const colorBackgrounds = sheets.map(sheet => {
    const tabColor = sheet.getTabColor();

    // If the sheet has no tab color, leave the management cell white.
    return [tabColor || "#ffffff"];
  });

  mgmtSheet
    .getRange(START_ROW, SHEET_COLOR_COL, colorBackgrounds.length, 1)
    .setBackgrounds(colorBackgrounds);

  ss.toast(
    `Listed ${outputData.length} sheet(s) for management.`,
    "Sheet List Complete",
    3
  );
}


/**
 * Reorders, hides/shows, and colors sheets based on sheet_slicer_management columns A:D.
 *
 * Sort behavior:
 * - Lower numbers in Column B move earlier.
 * - If two rows have the same order number, the row listed higher in the
 *   management sheet comes first.
 * - Blank or invalid order values are pushed to the bottom in row order.
 *
 * Hide behavior:
 * - TRUE / checked in Column C hides the sheet.
 * - FALSE / unchecked shows the sheet.
 * - Prevents hiding every sheet.
 *
 * Color behavior:
 * - Reads the cell background color from Column D.
 * - White in Column D clears the sheet tab color.
 * - Any non-white background color becomes the sheet tab color.
 */
function applySheetManagement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const mgmtSheet = ss.getSheetByName(MANAGEMENT_SHEET_NAME);

  if (!mgmtSheet) {
    ui.alert(
      "Missing Management Sheet",
      `Could not find a sheet named "${MANAGEMENT_SHEET_NAME}".`,
      ui.ButtonSet.OK
    );
    return;
  }

  const lastRow = mgmtSheet.getLastRow();

  if (lastRow < START_ROW) {
    ui.alert(
      "No Sheet Management Data",
      "No sheet management configuration found starting in Row 2.",
      ui.ButtonSet.OK
    );
    return;
  }

  const numRows = lastRow - START_ROW + 1;

  // Read values from A:D.
  const configRows = mgmtSheet
    .getRange(START_ROW, SHEET_NAME_COL, numRows, 4)
    .getValues();

  // Separately read background colors from Column D.
  // This is what controls the sheet tab color.
  const colorBackgrounds = mgmtSheet
    .getRange(START_ROW, SHEET_COLOR_COL, numRows, 1)
    .getBackgrounds();

  const managedSheets = configRows
    .map((row, index) => {
      const sheetName = String(row[0]).trim();
      const orderRaw = row[1];
      const hide = row[2] === true;
      const colorBackground = String(colorBackgrounds[index][0]).trim().toLowerCase();

      if (!sheetName) return null;

      const sheet = ss.getSheetByName(sheetName);

      if (!sheet) {
        Logger.log(`Skipping missing sheet "${sheetName}" listed on row ${index + START_ROW}.`);
        return null;
      }

      const requestedOrder = Number(orderRaw);

      return {
        sheet,
        sheetName,
        requestedOrder,
        hasValidOrder: Number.isFinite(requestedOrder),
        hide,
        tabColor: normalizeTabColor_(colorBackground),
        sourceRowIndex: index
      };
    })
    .filter(item => item !== null)
    .sort((a, b) => {
      if (a.hasValidOrder && b.hasValidOrder) {
        if (a.requestedOrder !== b.requestedOrder) {
          return a.requestedOrder - b.requestedOrder;
        }

        // Tie-breaker: keep the order from the management sheet rows.
        return a.sourceRowIndex - b.sourceRowIndex;
      }

      if (a.hasValidOrder) return -1;
      if (b.hasValidOrder) return 1;

      return a.sourceRowIndex - b.sourceRowIndex;
    });

  if (managedSheets.length === 0) {
    ui.alert(
      "No Valid Sheets Found",
      "No valid sheet names were found in Column A.",
      ui.ButtonSet.OK
    );
    return;
  }

  const sheetsToRemainVisible = managedSheets.filter(item => !item.hide);

  if (sheetsToRemainVisible.length === 0) {
    ui.alert(
      "Cannot Hide Every Sheet",
      "At least one sheet must remain visible. Uncheck Hide for at least one sheet.",
      ui.ButtonSet.OK
    );
    return;
  }

  ss.toast(
    "Applying sheet order, visibility, and color settings...",
    "Sheet Management",
    -1
  );

  // Temporarily show managed sheets because hidden sheets cannot be activated/moved.
  managedSheets.forEach(item => {
    if (item.sheet.isSheetHidden()) {
      item.sheet.showSheet();
    }
  });

  SpreadsheetApp.flush();

  // Move managed sheets into their requested sorted positions.
  managedSheets.forEach((item, index) => {
    ss.setActiveSheet(item.sheet);
    ss.moveActiveSheet(index + 1);
  });

  SpreadsheetApp.flush();

  // Apply tab colors.
  managedSheets.forEach(item => {
    item.sheet.setTabColor(item.tabColor);
  });

  SpreadsheetApp.flush();

  // Apply final visibility.
  managedSheets.forEach(item => {
    if (item.hide) {
      item.sheet.hideSheet();
    } else {
      item.sheet.showSheet();
    }
  });

  SpreadsheetApp.flush();

  ss.toast(
    `Reordered ${managedSheets.length} sheet(s), applied hide settings, and updated tab colors. Refreshing sheet list...`,
    "Sheet Management Complete",
    4
  );

  // Refresh Columns A:D so the management sheet reflects the actual final order,
  // visibility, and tab colors.
  listSheetsForManagement();
}


/**
 * Converts the management cell background into a valid sheet tab color.
 *
 * White/default cells are treated as "no tab color".
 */
function normalizeTabColor_(color) {
  if (!color) return null;

  const normalized = String(color).trim().toLowerCase();

  if (
    normalized === "#ffffff" ||
    normalized === "#fff" ||
    normalized === "white"
  ) {
    return null;
  }

  return normalized;
}


/****************************************************
 * SLICER MANAGEMENT
 ****************************************************/

/**
 * Dynamically reads target sheet names from Column J of sheet_slicer_management,
 * scans them for slicers, and refreshes the inventory list in Columns I:J.
 *
 * Column I: Slicer Name
 * Column J: Sheet Name
 * Column K: Target Cell, manually maintained
 */
function inventoryAllSlicers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const mgmtSheet = ss.getSheetByName(MANAGEMENT_SHEET_NAME);

  if (!mgmtSheet) {
    ui.alert(
      "Error",
      `Could not find the management sheet named "${MANAGEMENT_SHEET_NAME}".`,
      ui.ButtonSet.OK
    );
    return;
  }

  const lastRow = mgmtSheet.getLastRow();

  if (lastRow < START_ROW) {
    ui.alert(
      "No Sheets Listed",
      "Please list your target sheet names in Column J, starting at Row 2, before scanning.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Read existing sheet names from Column J.
  const sheetNameValues = mgmtSheet
    .getRange(START_ROW, SLICER_SHEET_COL, lastRow - START_ROW + 1, 1)
    .getValues();

  const sheetsToScan = [];

  sheetNameValues.forEach(row => {
    const sheetName = String(row[0]).trim();

    if (sheetName && sheetsToScan.indexOf(sheetName) === -1) {
      sheetsToScan.push(sheetName);
    }
  });

  if (sheetsToScan.length === 0) {
    ui.alert(
      "No Sheets Listed",
      "Column J does not contain any valid sheet names to scan.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Clear only Columns I:J. Preserve Column K target-cell assignments.
  mgmtSheet
    .getRange(START_ROW, SLICER_NAME_COL, lastRow - START_ROW + 1, 2)
    .clearContent();

  const outputData = [];

  sheetsToScan.forEach(sheetName => {
    const targetSheet = ss.getSheetByName(sheetName);

    if (!targetSheet) {
      Logger.log(`Skipping missing sheet "${sheetName}".`);
      return;
    }

    const slicers = targetSheet.getSlicers();

    slicers.forEach(slicer => {
      outputData.push([
        slicer.getTitle(),
        sheetName
      ]);
    });
  });

  if (outputData.length > 0) {
    mgmtSheet
      .getRange(START_ROW, SLICER_NAME_COL, outputData.length, 2)
      .setValues(outputData);

    ss.toast(
      `Successfully inventoried ${outputData.length} slicer(s).`,
      "Scan Complete",
      4
    );
  } else {
    ss.toast(
      "No slicers were found on the sheets listed in Column J.",
      "Scan Complete",
      3
    );
  }
}


/**
 * Reads settings from sheet_slicer_management to dynamically find and snap
 * slicers to assigned cell coordinates.
 *
 * Column I: Slicer Name
 * Column J: Sheet Name
 * Column K: Target Cell
 */
function alignSlicersFromConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const mgmtSheet = ss.getSheetByName(MANAGEMENT_SHEET_NAME);

  if (!mgmtSheet) {
    ui.alert(
      "Layout Configuration Error",
      `Could not find the sheet named "${MANAGEMENT_SHEET_NAME}".`,
      ui.ButtonSet.OK
    );
    return;
  }

  const lastRow = mgmtSheet.getLastRow();

  if (lastRow < START_ROW) {
    ui.alert(
      "Empty Configuration",
      "No slicer layout configuration found starting at Row 2.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Read Columns I:K starting from Row 2.
  const configData = mgmtSheet
    .getRange(START_ROW, SLICER_NAME_COL, lastRow - START_ROW + 1, 3)
    .getValues();

  ss.toast("Processing layout assignments...", "Slicer Alignment", -1);

  let matchCount = 0;
  const sheetCache = {};

  configData.forEach((row, index) => {
    const slicerTitle = String(row[0]).trim();
    const targetSheetName = String(row[1]).trim();
    const targetCellA1 = String(row[2]).trim();

    if (!slicerTitle || !targetSheetName || !targetCellA1) {
      return;
    }

    if (!sheetCache[targetSheetName]) {
      sheetCache[targetSheetName] = ss.getSheetByName(targetSheetName);
    }

    const targetSheet = sheetCache[targetSheetName];

    if (!targetSheet) {
      Logger.log(`Sheet "${targetSheetName}" listed on row ${index + START_ROW} does not exist.`);
      return;
    }

    const targetSlicers = targetSheet.getSlicers();
    const matchingSlicer = targetSlicers.find(
      slicer => slicer.getTitle() === slicerTitle
    );

    if (!matchingSlicer) {
      Logger.log(`No slicer titled "${slicerTitle}" found on sheet "${targetSheetName}".`);
      return;
    }

    try {
      const targetRange = targetSheet.getRange(targetCellA1);
      const targetRow = targetRange.getRow();
      const targetCol = targetRange.getColumn();

      // Step 1: Force baseline position alignment to the cell coordinates.
      matchingSlicer.setPosition(targetRow, targetCol, 0, 0);
      SpreadsheetApp.flush();

      // Step 2: Trigger 1px displacement to bypass browser coordinate rendering bugs.
      matchingSlicer.setPosition(targetRow, targetCol, 1, 1);
      SpreadsheetApp.flush();

      // Step 3: Firmly seat the graphic back into position.
      matchingSlicer.setPosition(targetRow, targetCol, 0, 0);

      matchCount++;
    } catch (err) {
      Logger.log(
        `Failed to position "${slicerTitle}" at cell "${targetCellA1}": ${err.message}`
      );
    }
  });

  SpreadsheetApp.flush();

  ss.toast(
    `Successfully aligned ${matchCount} slicer(s).`,
    "Alignment Complete",
    3
  );
}