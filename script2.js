// three built-in event listener functions
function onOpen() {
  setRemainingDays();
}

function onEdit() {
  setRemainingDays();
}

function onChange() {
  setRemainingDays();
}

// function for setting remaining days
function setRemainingDays() {
  // !!!
  // user should provide 6 values
  // sheetName, startRow, endDateColumn, remainingDaysColumnOne, checkDateColumn, remainingDaysColumnTwo;

  // name of the sheet; example: ='Sheet1';
  const sheetName = 'Sheet1';

  // number of the date column's first row, example: should be 2 if the first row is a header
  const startRow = 2;

  // end date column letter: example: ='B';
  const endDateColumn = 'B';

  // first remaining days column letter:
  const remainingDaysColumnOne = 'C';

  // check date column letter:
  const checkDateColumn = 'F';

  // second remaining days column letter:
  const remainingDaysColumnTwo = 'G';

  // !!!
  // rest of the code does not require user interaction

  // get current date
  let date = new Date();

  // remove time from the current date
  date.setHours(0, 0, 0, 0);

  // convert column letters to numbers
  const endDateCol = columnLetterToNumber(endDateColumn);
  const checkDateCol = columnLetterToNumber(checkDateColumn);
  const remainingDaysCol1 = columnLetterToNumber(remainingDaysColumnOne);
  const remainingDaysCol2 = columnLetterToNumber(remainingDaysColumnTwo);

  // run set days function
  setDays(sheetName, startRow, endDateCol, remainingDaysCol1, date);
  setDays(sheetName, startRow, checkDateCol, remainingDaysCol2, date);
}

// function to set remaining days in sheet cells
function setDays(sheetName, startRow, endColumn, remainingDays, date) {
  // get sheet by the provided name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  // get number of the last row with data
  const lastRow = sheet.getLastRow();

  // get an array of values of the column
  let valArr = getCellValues(sheet, startRow, endColumn, lastRow);

  // iterate through the array
  valArr.forEach((val, index) => {
    // check if a cell has a value
    if (val[0]) {
      // set a row
      let row = startRow + index;

      // get day difference number
      let dif = getDayDifference(date, val[0]);

      // set remaining days number to cell
      setCellValue(sheet, row, remainingDays, dif);
    }
  });
}

// get the difference between dates in days
function getDayDifference(curDate, checkDate) {
  if (checkDate - curDate <= 0) return 0;
  else return (checkDate - curDate) / (1000 * 3600 * 24);
}

// set a new value to the cell
function setCellValue(sheet, row, column, data) {
  sheet.getRange(row, column).setValue(data);
}

// get range values as array
function getCellValues(sheet, row, column, rowCount = 1, columnCount = 1) {
  return sheet.getRange(row, column, rowCount, columnCount).getValues();
}

// convert column letter to number
function columnLetterToNumber(letter) {
  let len = letter.length;
  if (len === 1) {
    return letter.charCodeAt(0) - 64;
  } else if (len > 1) {
    let letterNum = 0;
    for (let i = 0; i < len; i++) {
      let pow = len - 1 - i;
      letterNum += (letter[i].charCodeAt(0) - 64) * 26 ** pow;
    }
    return letterNum;
  }
}
