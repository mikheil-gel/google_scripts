function onEdit(e) {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getActiveSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (col === 1) {
    activeSheet.getRange(row, 2).setValue(e.user + ' ' + new Date());
  }
}
