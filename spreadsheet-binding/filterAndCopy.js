// !!!
// script uses advanced service: Sheets API
// it should be added manually:
// Apps Script (Editor)> Services > Google Sheets API > Add;

// !!!
// 5 values should be provided:
// localSheetName, headerRows, targetSpreadsheetId,runOnEveryDays,setFilterFromScript

// local sheet name
const localSheetName = 'Sheet1';

// header rows count
const headerRows = 2;

// target spreadsheet id
const targetSpreadsheetId = '1BeF-2ax7VS7boKktnlt3NI6Y0RYWkwOzaHcjUa9AoMQ';

// number of days function should run (1 - once a day, 7 - once a week)
const runOnEveryDays = 1;

// option to set filters from the script (boolean value)
const setFilterFromScript = false;
// !!!
// if filter should be set from the script (setFilterFromScript is true),
// array/s should be added in filterArray:

// fillterArray should contain array/s with two values:
// first value: column's (capital)letter/s;
// second value: criteria bulider with method
// example: ['A', SpreadsheetApp.newFilterCriteria().whenTextNotEqualTo('text')]
// criteria full list can be found on:
// https://developers.google.com/apps-script/reference/spreadsheet/filter-criteria-builder

const filterArray = [
  ['C', SpreadsheetApp.newFilterCriteria().whenTextEqualTo('vue')],
  ['D', SpreadsheetApp.newFilterCriteria().whenNumberGreaterThan(1)],
];

// !!!
// this function should run manually from the Apps Script, when:
// it is first run or runOnEveryDays variable was changed
// creates time based trigger function
// copies current data to target sheet
// while code should work correctly, I faced some unexpected behavior when changing trigger option;
// would recommend deleting existing trigger manually, before running a function again;
function createTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  // check if trigger exists
  let timeTriggerIndex = null;
  triggers.forEach((item, index) => {
    if (item.getHandlerFunction() === 'timeBasedEvent') timeTriggerIndex = index;
  });

  // delete trigger if it exists
  if (timeTriggerIndex !== null) {
    ScriptApp.deleteTrigger(triggers[timeTriggerIndex]);
  }

  // create time based trigger
  ScriptApp.newTrigger('timeBasedEvent').timeBased().everyDays(runOnEveryDays).create();

  // copy data to target spreadsheet
  copyData();
}

// trgger function runs once in specified (runOnEveryDays) days
function timeBasedEvent() {
  copyData();
}

// function for coping values to target spreadsheet
function copyData() {
  // get local sheet data
  const app = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = app.getSheetByName(localSheetName);

  // get last Row
  const lastRow = sheet.getLastRow();
  // get last Column
  const lastColumn = sheet.getLastColumn();
  // get range with data
  const dataRange = sheet.getDataRange();

  // get range address
  const rangeNotation = dataRange.getA1Notation();

  // get filter Range
  let filter = sheet.getFilter();
  let originalCriteria = [];

  // set filter from the script
  if (setFilterFromScript) {
    if (!filter) {
      // create new filter
      const filterRangeNotation = rangeNotation.replace('A1', `A${headerRows}`);
      filter = sheet.getRange(filterRangeNotation).createFilter();
    } else {
      for (let i = 1; i <= lastColumn; i++) {
        if (filter.getColumnFilterCriteria(i)) {
          // save original filter criteria
          originalCriteria.push(filter.getColumnFilterCriteria(i).copy());
          // clear criteria
          filter.removeColumnFilterCriteria(i);
        } else {
          originalCriteria.push(false);
        }
      }
    }
    // set new filter criteria
    filterArray.forEach((item) => {
      filter.setColumnFilterCriteria(columnLetterToNumber(item[0]), item[1].build());
    });
  }

  // get hidden rows
  let hiddenRowsIndexes = getHiddenRowsinGoogleSheets(app.getId(), sheet.getSheetId());

  // reset to original filter criteria
  if (setFilterFromScript) {
    for (let i = 1; i <= lastColumn; i++) {
      if (filter.getColumnFilterCriteria(i)) filter.removeColumnFilterCriteria(i);
      if (originalCriteria[i - 1]) filter.setColumnFilterCriteria(i, originalCriteria[i - 1].build());
    }
  }

  // get range values
  let rangeValues = dataRange.getValues().filter((item, index) => !hiddenRowsIndexes.includes(index));

  // get target spreadsheet
  const targetApp = SpreadsheetApp.openById(targetSpreadsheetId);

  // get current time
  const date = Utilities.formatDate(new Date(), 'GMT', 'yyyy/MM/dd HH:mm:ss');
  // create new sheet
  let targetSheet = targetApp.insertSheet();
  // set time as a sheet name
  targetSheet.setName(date);
  // get target range
  const targetRange = targetSheet.getRange(
    ['A1', rangeNotation.split(':')[1].replace(lastRow, lastRow - hiddenRowsIndexes.length)].join(':')
  );

  // set target sheet values
  targetRange.setValues(rangeValues);
}

// function to convert column letters to numbers
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

// function to get hidden rows;
const getHiddenRowsinGoogleSheets = (spreadsheetId, sheetId) => {
  const fields = 'sheets(data(rowMetadata(hiddenByFilter)),properties/sheetId)';
  const { sheets } = Sheets.Spreadsheets.get(spreadsheetId, { fields });
  const [sheet] = sheets.filter(({ properties }) => {
    return String(properties.sheetId) === String(sheetId);
  });

  const { data: [{ rowMetadata = [] }] = {} } = sheet;

  const hiddenRows = rowMetadata
    .map(({ hiddenByFilter }, index) => {
      return hiddenByFilter ? index : -1;
    })
    .filter((rowId) => rowId !== -1);

  return hiddenRows;
};
