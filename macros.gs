function Delete() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A3:D33300').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};