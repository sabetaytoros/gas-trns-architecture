




function copyG() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('SPCX'), true);
  spreadsheet.getCurrentCell().offset(-13, -2).activate();
  spreadsheet.getRange('ORCL!F7:F8').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};