function onEdit(e) {
  setUserAndDate(e);
}

function setUserAndDate(e) {
  //our variables
  var usersSheetName = 'Sheet1';
  var editableColumn = 1;
  var startRow = 2;

  //data from event
  var activeSheet = e.source.getActiveSheet();
  var activeSheetName = activeSheet.getName();
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (activeSheetName === usersSheetName && col === editableColumn && row >= startRow) {
    var date = new Date();
    var userColumn = editableColumn + 1;
    var dateColumn = editableColumn + 2;
    setCellValue(activeSheet, row, userColumn, e.user);
    setCellValue(activeSheet, row, dateColumn, date);
  }
}

function setCellValue(sheet, row, col, data) {
  sheet.getRange(row, col).setValue(data);
}
