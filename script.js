// built-in event listener function
function onEdit(e) {
  // run function on edit event
  setUserAndDate(e);
}

// function to set user and date after edit happened
function setUserAndDate(e) {
  // !!!
  // User should provide 4 values:
  // sheetName, columnStart, columnEnd, startRow;

  // Name of the sheet; example: ='Sheet1';
  var sheetName = 'Sheet1';

  // change detection start column letter: example: ='A';
  var columnStart = 'A';

  // change detection end column letter, if it is just one column, should be same as 'columnStart';
  var columnEnd = 'A';

  // change detection start row number, example: = 2;  (if first row is header)
  var startRow = 2;

  // !!!
  // rest of the code does not require user interaction

  // converting column letters to numbers
  var colStart = columnLetterToNumber(columnStart);
  var colEnd = columnStart === columnEnd ? colStart : columnLetterToNumber(columnEnd);

  // getting data from the event object
  var activeSheet = e.source.getActiveSheet();
  var activeSheetName = activeSheet.getName();
  var col = e.range.getColumn();
  var row = e.range.getRow();

  // checking if edit happened at the change detection range
  if (activeSheetName === sheetName && col <= colEnd && col >= colStart && row >= startRow) {
    // formatting current date
    var date = Utilities.formatDate(new Date(), 'GMT', 'yyyy/MM/dd HH:mm:ss');

    // appending user and date into next two columns
    var userColumn = colEnd + 1;
    var dateColumn = colEnd + 2;
    setCellValue(activeSheet, row, userColumn, e.user);
    setCellValue(activeSheet, row, dateColumn, date);
  }
}

// set new value to the cell
function setCellValue(sheet, row, column, data) {
  sheet.getRange(row, column).setValue(data);
}

// convert column letter to number
function columnLetterToNumber(letter) {
  var len = letter.length;
  if (len === 1) {
    return letter.charCodeAt(0) - 64;
  } else if (len > 1) {
    var letterNum = 0;
    for (var i = 0; i < len; i++) {
      var pow = len - 1 - i;
      letterNum += (letter[i].charCodeAt(0) - 64) * 26 ** pow;
    }
    return letterNum;
  }
}
