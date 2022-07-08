// built-in event listener function
function onEdit(e) {
  // run function on edit event
  setUserAndDate(e);
}

// function to set user and date after edit happened
function setUserAndDate(e) {
  // !!!
  // User should provide 6 values:
  // sheetName, startRow, startColumn, endColumn, userColumn, dateColumn;

  // Name of the sheet; example: ='Sheet1';
  const sheetName = 'Sheet1';

  // change detection start row number, example: = 2;  (if first row is header)
  const startRow = 2;

  // change detection start column letter: example: ='A';
  const startColumn = 'A';

  // change detection end column letter, if it is just one column, should be same as 'startColumn';
  const endColumn = 'A';

  // user column letter;
  const userColumn = 'B';

  // date column letter;
  const dateColumn = 'C';

  // !!!
  // rest of the code does not require user interaction

  // converting column letters to numbers
  const startCol = columnLetterToNumber(startColumn);
  const endCol = startColumn === endColumn ? startCol : columnLetterToNumber(endColumn);
  const userCol = columnLetterToNumber(userColumn);
  const dateCol = columnLetterToNumber(dateColumn);

  // getting data from the event object
  const activeSheet = e.source.getActiveSheet();
  const activeSheetName = activeSheet.getName();
  const firstCol = e.range.getColumn();
  const lastCol = e.range.getLastColumn();
  const firstRow = e.range.getRow();
  const lastRow = e.range.getLastRow();

  // checking if edit happened at the change detection range
  if (activeSheetName === sheetName && firstCol <= endCol && lastCol >= startCol && lastRow >= startRow) {
    // formatting current date
    const date = Utilities.formatDate(new Date(), 'GMT', 'yyyy/MM/dd HH:mm:ss');

    // appending user and date into columns
    let row = firstRow >= startRow ? firstRow : startRow;
    while (row <= lastRow) {
      setCellValue(activeSheet, row, userCol, e.user);
      setCellValue(activeSheet, row, dateCol, date);
      row++;
    }
  }
}

// set new value to the cell
function setCellValue(sheet, row, column, data) {
  sheet.getRange(row, column).setValue(data);
}

// convert column letter to number
function columnLetterToNumber(letter) {
  let len = letter.length;
  if (len === 1) {
    return letter.charCodeAt(0) - 64;
  } else if (len > 1) {
    let letterNum = 0;
    for (var i = 0; i < len; i++) {
      let pow = len - 1 - i;
      letterNum += (letter[i].charCodeAt(0) - 64) * 26 ** pow;
    }
    return letterNum;
  }
}
