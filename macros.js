/**
 * Moves cell values from a source range to a destination range.
 * Optionally clears the source after moving (Cut & Paste behavior).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The target sheet object.
 * @param {string} sourceA1 - Source range in A1 notation (e.g., "G2:G100" or "A1").
 * @param {string} destA1 - Destination start cell/range in A1 notation (e.g., "L2" or "B1").
 * @param {boolean} [clearSource=true] - If true, clears the source range after copying.
 */
function test() {
  Trns.trs = Trns.sht.getActiveSheet();
  Trns.Name = Trns.trs.getSheetName()  
  Logger.log('Name %s',Trns.Name); 
  row = nextcell()-1
  src = 'B'+ row + ':K'+ Trns.trs.getLastRow()
  dst = 'O' +row
  MOVE_CELL_RANGE(Trns.trs, src, dst) 
}
function MOVE_CELL_RANGE(sheet, sourceA1, destA1) {
  // --- CONFIGURATION & DEFAULTS ---
  var SHOULD_CLEAR_SOURCE = true;
  
  if (!sheet) {
    Logger.log("Error: Invalid sheet provided to MOVE_CELL_RANGE.");
    return;
  }// --- STEP 1: READ SOURCE DATA (IN-MEMORY) ---
  var sourceRange = sheet.getRange(sourceA1);
  var values = sourceRange.getValues();
  
  var numRows = values.length;
  var numCols = values[0].length;

  // --- STEP 2: WRITE TO DESTINATION ---
  var destRange = sheet.getRange(destA1);
  var destStartRow = destRange.getRow();
  var destStartCol = destRange.getColumn();

  // Target exact dimensions to prevent mismatch errors
  var targetRange = sheet.getRange(destStartRow, destStartCol, numRows, numCols);
  targetRange.setValues(values);

  // --- STEP 3: CLEAR SOURCE (CUT OPERATION) ---
  if (SHOULD_CLEAR_SOURCE) {
    sourceRange.clearContent();
  }

  Logger.log("Successfully moved range " + sourceA1 + " to " + destA1 + " (Rows: " + numRows + ", Cols: " + numCols + ")");
}

/**
 * Example Usage / Driver Test Function
 */
function TEST_MOVE_EXECUTION() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Example 1: Move derived states from G2:G100 to L2:L100 (and clear G2:G100)
  var SOURCE_RANGE = "G2:G100";
  var DESTINATION_START_CELL = "L2";
  var CLEAR_ORIGINAL = true;
  
  MOVE_CELL_RANGE(sheet, SOURCE_RANGE, DESTINATION_START_CELL, CLEAR_ORIGINAL);
}
