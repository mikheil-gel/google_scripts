// !!!
// script uses advanced service: Sheets API
// it should be added manually:
// Apps Script (Editor)> Services > Google Sheets API > Add;

// !!!
// 11 values should be provided:
// localSpreadsheetId, dataSpreadsheetsArray, dataHeaderRows, copyStartColumn, copyEndColumn, columnsToIgnore, runOnEveryDays, enableWeekday, weekday, runAtHour, setFilterFromScript

// local spreadsheet id
const localSpreadsheetId = '1S_gzGL2yRDcvqB6V6iZOSkUkD9jbhQ45FCz3TtTg11c';

// data spreadsheet array should contain array/s with two values:
// data spreadsheet Id
// sheet name with data
// example: const dataSpreadsheetsArray = [ ['19D-XDwswxFfl--hTlc-Jc7eF_G_CfXdLmfTfYddzziE','Sheet1'], ... ]
const dataSpreadsheetsArray = [
  ['142qUNhUTq3KUYIgTl9fmZKOCyemoDlYfH4_ZRXRMS7w', 'Sheet1'],
  ['1VmMqIaMoVwVTC4PhTKG6Xfc8ziHH9sH_xxCjmkRAa30', 'Sheet1'],
];

// data header rows count
const dataHeaderRows = 2;

// copy range start column letter
const copyStartColumn = 'A';

// copy range end column letter
const copyEndColumn = 'F';

// column letter/s that should be ignored
// example: const columnsToIgnore = ['B','C']
// array should be empty if all columns are needed
const columnsToIgnore = [];

// number of days function should run (1 - once a day, 7 - once a week)
// is used when 'enableWeekday' is false
const runOnEveryDays = 1;

// option to use 'weekday' instead of 'runOnEveryDays'
const enableWeekday = false;

// name of the day function should run: ScriptApp.WeekDay.MONDAY
// is used when 'enableWeekday' is true
const weekday = ScriptApp.WeekDay.MONDAY;

// time (hour) function should run (integers from 0 to 23)
const runAtHour = 0;

// option to set filters from the script (boolean value)
const setFilterFromScript = true;
// !!!
// if filter should be set from the script (setFilterFromScript is true),
// array/s should be added in filterArray:

// fillterArray should contain array/s with two values:
// first value: column's (capital)letter/s;
// second value: criteria bulider with method
// example: ['A', SpreadsheetApp.newFilterCriteria().whenTextNotEqualTo('text')]
// criteria full list can be found on:
// mine: https://github.com/mikheil-gel/google_scripts/blob/main/filterAndCopy/filterCriteria.txt
// official: https://developers.google.com/apps-script/reference/spreadsheet/filter-criteria-builder

const filterArray = [
  ['D', SpreadsheetApp.newFilterCriteria().whenTextContains('web')],
  // ['E', SpreadsheetApp.newFilterCriteria().whenNumberGreaterThan(1)],
];

// !!!
// this function should run manually from the Apps Script, when:
// it is first run or runOnEveryDays/runAtHour/weekday variables were changed
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
  if (enableWeekday) {
    ScriptApp.newTrigger('timeBasedEvent').timeBased().onWeekDay(weekday).atHour(runAtHour).create();
  } else {
    ScriptApp.newTrigger('timeBasedEvent').timeBased().everyDays(runOnEveryDays).atHour(runAtHour).create();
  }

  // copy data to target spreadsheet
  copyData();
}

// trigger function runs once in specified (runOnEveryDays/weekday) days
function timeBasedEvent() {
  copyData();
}

// function for coping values to target spreadsheet
function copyData() {
  // get target spreadsheet
  const targetSpreadsheet = SpreadsheetApp.openById(localSpreadsheetId);

  // get spreadsheet's time zone
  let timeZone = targetSpreadsheet.getSpreadsheetTimeZone();
  // get current time
  const date = Utilities.formatDate(new Date(), timeZone, 'yyyy/MM/dd HH:mm:ss');
  // create new sheet
  let targetSheet = targetSpreadsheet.insertSheet();
  // set time as a sheet name
  targetSheet.setName(date);

  dataSpreadsheetsArray.forEach((spreadsheetsData, index) => {
    // get spreadsheet id and sheet name
    const [dataSpreadsheetId, dataSheetName] = spreadsheetsData;
    // get data sheet
    const spreadsheet = SpreadsheetApp.openById(dataSpreadsheetId);
    const sheet = spreadsheet.getSheetByName(dataSheetName);

    const firstSheet = index === 0;

    const firstDataRow = firstSheet ? 1 : dataHeaderRows + 1;

    // get last Row
    const lastRow = sheet.getLastRow();
    // get last Column
    const lastColumn = sheet.getLastColumn();
    // get range with data
    const dataRange = sheet.getDataRange();
    // get header range
    const headerRange = sheet.getRange(`${copyStartColumn}1:${copyEndColumn}${dataHeaderRows}`);
    // get copy range
    const copyRange = sheet.getRange(`${copyStartColumn}${firstDataRow}:${copyEndColumn}${lastRow}`);

    // get range address
    const rangeNotation = dataRange.getA1Notation();

    // get filter Range
    let filter = sheet.getFilter();
    const filterPresent = !!filter;
    let originalFilterRange = filterPresent ? filter.getRange().getA1Notation() : '';
    let originalCriteria = [];

    // set filter from the script
    if (setFilterFromScript) {
      if (filterPresent) {
        for (let i = filter.getRange().getColumn(), n = filter.getRange().getLastColumn(); i <= n; i++) {
          if (filter.getColumnFilterCriteria(i)) {
            // save original filter criteria
            originalCriteria.push({ column: i, criteria: filter.getColumnFilterCriteria(i).copy() });
          }
        }
        // remove present filter
        filter.remove();
      }
      // create new filter
      const filterRangeNotation = rangeNotation.replace('A1', `A${dataHeaderRows}`);
      filter = sheet.getRange(filterRangeNotation).createFilter();

      // set new filter criteria
      filterArray.forEach((item) => {
        filter.setColumnFilterCriteria(columnLetterToNumber(item[0]), item[1].build());
      });
    }

    // get hidden rows
    let hiddenRowsIndexes = getHiddenRowsinGoogleSheets(dataSpreadsheetId, sheet.getSheetId());

    // reset to original filter criteria
    if (setFilterFromScript) {
      filter.remove();
      if (filterPresent) {
        filter = sheet.getRange(originalFilterRange).createFilter();
        if (originalCriteria.length) {
          originalCriteria.forEach((data) => {
            filter.setColumnFilterCriteria(data.column, data.criteria.build());
          });
        }
      }
    }

    const rowCount = firstSheet
      ? lastRow - hiddenRowsIndexes.length
      : lastRow - hiddenRowsIndexes.length - dataHeaderRows;

    // check if filtered data exists
    if (rowCount >= 1) {
      const rowIndexCheck = firstSheet ? 0 : dataHeaderRows;

      // get range values
      let rangeValues = copyRange
        .getValues()
        .filter((item, index) => !hiddenRowsIndexes.includes(index + rowIndexCheck));

      // get links
      let linksArr = [];

      copyRange
        .getRichTextValues()
        .filter((item, index) => !hiddenRowsIndexes.includes(index + rowIndexCheck))
        .forEach((rt, rowIndex) => {
          rt.forEach((ct, columnIndex) => {
            let links = [];
            let indexes = [];
            ct.getRuns().forEach((rr) => {
              let link = rr.getLinkUrl();
              if (link) {
                links.push(link);
                indexes.push([rr.getStartIndex(), rr.getEndIndex()]);
              }
            });

            if (links.length) linksArr.push({ links, indexes, row: rowIndex, column: columnIndex + 1 });
          });
        });

      const firstTargetRow = firstSheet ? 1 : targetSheet.getLastRow() + 1;

      // get target range
      const columnCount = copyRange.getLastColumn() - copyRange.getColumn() + 1;
      const targetRange = targetSheet.getRange(firstTargetRow, 1, rowCount, columnCount);

      targetRange.setValues(rangeValues);

      if (firstSheet) {
        // get header colors
        const headerBackground = headerRange.getBackgrounds();
        const headerFontColor = headerRange.getFontColors();

        const targetHeaderRange = targetSheet.getRange(1, 1, dataHeaderRows, columnCount);

        // set header colors
        targetHeaderRange.setFontColors(headerFontColor);
        targetHeaderRange.setBackgrounds(headerBackground);
      }

      // set links
      if (linksArr.length) {
        linksArr.forEach((data) => {
          let cell = targetSheet.getRange(data.row + firstTargetRow, data.column);
          let cellValue = cell.getValue();
          let length = data.links.length;

          let richText = SpreadsheetApp.newRichTextValue().setText(cellValue);
          for (let i = 0; i < length; i++) {
            richText.setLinkUrl(...data.indexes[i], data.links[i]);
          }
          richText = richText.build();

          cell.setRichTextValue(richText);
        });
      }
    }
  });

  // delete to be ignored columns
  if (columnsToIgnore.length) {
    let startColumn = columnLetterToNumber(copyStartColumn);
    let columnsToDelete = columnsToIgnore
      .map((letter) => columnLetterToNumber(letter))
      .filter((colNum) => colNum > startColumn)
      .map((midColNum) => midColNum - startColumn + 1)
      .sort((a, b) => b - a);
    columnsToDelete.forEach((columnNumber) => {
      targetSheet.deleteColumn(columnNumber);
    });
  }
}

// function to convert column letters to numbers
function columnLetterToNumber(letter) {
  let len = letter.length;
  if (len === 1) {
    return letter.charCodeAt(0) - 64;
  } else if (len > 1) {
    let letterNum = 0;
    for (let i = 0; i < len; i++) {
      let pow = len - 1 - i;
      letterNum += (letter.charCodeAt(i) - 64) * 26 ** pow;
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
