// !!!
// after providing all the necessary values below,
// user should choose and run function 'createTriggers' from the menu bar above

// parent (local) spreadsheet spreadsheetId
const parentSpreadsheetId = '11Rpx0aY-AaZqTlLAfu9mYl_S2AzziX5jhzA9EqaP3B4';

// parent (local) sheet name, that should be synced
const parentSheetName = 'Sheet1';

// header rows count on the parent (local) sheet
const headerRows = 2;

// column letter with unique values
const uniqueValuesColumn = 'A';

// array of columns letters, where data should merge
const columnsToMerge = [];

// !!!
// user should add array/s in childrenSpreadsheetsArray;
// arrays should have 3 values:
// target (child) spreadsheet spreadsheetId;
// target (child) sheet name
// name of the trigger function (should be created manually)
// example: ['1qXSJDoqDEewRMRHkwQdLVoptm89zBKuKDbY8ZwKdo5U', 'Sheet1', 'event1' ]

const childrenSpreadsheetsArray = [['1NldYfCGJdgW2L7oDb_uhLEQYNID2_YNoCZLwdjB_lIw', 'Sheet1', 'event1']];

// !!!
// functions should be created manually with name, provided as 4th value of childern arrays above:
// e - event object should be passed as a parameter
// updateSpreadsheet(e) function should be called;
// example: function event1(e){ updateSpreadsheet(e) };

function event1(e) {
  updateSpreadsheet(e);
}

// !!!
// this function should be run manually from the Apps Script
// removes older triggers (would recommend to do that manually to be more safe);
// copies data from the parent spreadsheet to children;
// creates change event trigger for every spreadsheet
// permissions should be granted
function createTriggers() {
  // create an array of trigger function names
  let triggerFunctions = childrenSpreadsheetsArray.map((it) => it[2]);
  triggerFunctions.push('changeEvent');

  // check if trigger already exists
  const triggers = ScriptApp.getProjectTriggers();
  let triggerIndexes = [];
  triggers.forEach((item, index) => {
    if (triggerFunctions.includes(item.getHandlerFunction())) triggerIndexes.push(index);
  });

  // delete triggers if they exists
  if (triggerIndexes.length) {
    triggerIndexes.forEach((triggerIndex) => {
      ScriptApp.deleteTrigger(triggers[triggerIndex]);
    });
  }

  // create new triggers
  // triggers for the children spreadsheets
  childrenSpreadsheetsArray.forEach((arr) => {
    // create triggers for children spreadsheets
    ScriptApp.newTrigger(arr[2]).forSpreadsheet(arr[0]).onEdit().create();
  });

  // create trigger for the parent spreadsheet
  ScriptApp.newTrigger('changeEvent').forSpreadsheet(parentSpreadsheetId).onEdit().create();
}

// trigger function for the parent spreadsheet
function changeEvent(e) {
  updateSpreadsheet(e);
}

// function for updating target spreadsheets
function updateSpreadsheet(e) {
  // get edited spreadsheet's id
  const spreadsheetId = e.source.getId();
  // get edited sheet name
  const activeSheetName = e.source.getActiveSheet().getName();
  // check if edit happened on the parent spreadsheet
  const localChange = parentSpreadsheetId === spreadsheetId;

  // create variable to store spreadsheets data
  let spreadsheetsDataArray = [];
  // fill 'spreadsheetsDataArray' with data
  if (localChange) {
    if (parentSheetName === activeSheetName) {
      spreadsheetsDataArray = childrenSpreadsheetsArray.map((item) => [
        item[0],
        item[1],
        parentSpreadsheetId,
        parentSheetName,
      ]);
    }
  } else {
    let childSpreadsheetData = childrenSpreadsheetsArray.find(
      (arr) => arr[0] === spreadsheetId && arr[1] === activeSheetName
    );
    if (childSpreadsheetData.length) {
      spreadsheetsDataArray = [
        [parentSpreadsheetId, parentSheetName, childSpreadsheetData[0], childSpreadsheetData[1]],
      ];
    }
  }

  // check if edit happened on the valid sheet
  if (spreadsheetsDataArray.length) {
    // loop through spreadsheets data
    for (let dataArray of spreadsheetsDataArray) {
      // get data
      let [targetId, targetSheetName, localId, localSheetName] = dataArray;
      // get spreadsheets
      const [changedSpreadsheet, targetSpreadsheet] = [
        SpreadsheetApp.openById(localId),
        SpreadsheetApp.openById(targetId),
      ];
      // get changed sheet
      const sheet = changedSpreadsheet.getSheetByName(localSheetName);

      // get unique column number
      const uniqueColumn = columnLetterToNumber(uniqueValuesColumn);

      // get edited range address
      const editedRowStart = e.range.rowStart > headerRows ? e.range.rowStart : headerRows + 1;
      const editedRowEnd = e.range.rowEnd;
      const editedColumnStart = e.range.columnStart;
      const editedColumnEnd = e.range.columnEnd;

      let editedRowCount = editedRowEnd - editedRowStart + 1;
      let editedColumnCount = editedColumnEnd - editedColumnStart + 1;

      // create array for edited rows
      let editedRowsRange = [];
      // fill 'editedRowsRange' with data
      if (editedRowStart === editedRowEnd) {
        editedRowsRange.push(editedRowStart);
      } else if (editedRowStart < editedRowEnd) {
        for (let i = editedRowStart; i <= editedRowEnd; i++) {
          editedRowsRange.push(i);
        }
      }

      // get target sheet
      let targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

      // copy data if changed and target sheets exist and edit didn't happen in header
      if (sheet && targetSheet && editedRowsRange.length) {
        // get target sheet's last row with data
        const targetLastRow = targetSheet.getLastRow();

        // get edited sheet's unique colummn's data
        const localColumnData = sheet.getRange(editedRowStart, uniqueColumn, editedRowCount).getValues();
        // get target sheet's unique colummn's data
        const targetColumnData = targetSheet
          .getRange(headerRows + 1, uniqueColumn, targetLastRow - headerRows)
          .getValues();

        // create array for local and destination cell addresses
        let dataToCopy = [];

        localColumnData.forEach((localValue, localRowIndex) => {
          if (localValue[0].toString().length) {
            for (let [targetRowIndex, targetValue] of targetColumnData.entries()) {
              if (targetValue[0].toString().length) {
                if (localValue[0] === targetValue[0]) {
                  dataToCopy.push({
                    localRow: editedRowsRange[localRowIndex],
                    targetRow: targetRowIndex + headerRows + 1,
                  });
                  break;
                }
              }
            }
          }
        });

        // check if data in columns should merge
        if (columnsToMerge.length) {
          let editedColumnsRange = [];
          if (editedColumnStart === editedColumnEnd) {
            editedColumnsRange.push(editedColumnStart);
          } else {
            for (let i = editedColumnStart; i <= editedColumnEnd; i++) {
              editedColumnsRange.push(i);
            }
          }
          let editedColumnsToMerge = columnsToMerge
            .map((letter) => columnLetterToNumber(letter))
            .filter((column) => editedColumnsRange.includes(column));
          if (editedColumnsToMerge.length) {
            if (e.oldValue) {
              if (!e.value.toString().includes(e.oldValue)) {
                sheet.getRange(editedRowStart, editedColumnStart).setValue(e.oldValue + ' ' + e.value);
              }
            } else if (dataToCopy.length) {
              editedColumnsToMerge.forEach((column) => {
                dataToCopy.forEach((data) => {
                  let localRange = sheet.getRange(data.localRow, column);
                  let localValue = localRange.getValue();
                  let targetValue = targetSheet.getRange(data.targetRow, column).getValue();
                  if (targetValue.toString().length && !localValue.toString().includes(targetValue)) {
                    localRange.setValue(targetValue + ' ' + localValue);
                  }
                });
              });
            }
          }
        }

        if (dataToCopy.length) {
          // copy edited sheet to target spreadsheet
          let copiedSheet = sheet.copyTo(targetSpreadsheet);
          for (let rowData of dataToCopy) {
            // copy changed cells to target sheet
            let copiedDataRange = copiedSheet.getRange(rowData.localRow, editedColumnStart, 1, editedColumnCount);
            let copyToDataRange = targetSheet.getRange(rowData.targetRow, editedColumnStart, 1, editedColumnCount);
            copiedDataRange.copyTo(copyToDataRange);
          }
          // delete copied sheet
          targetSpreadsheet.deleteSheet(copiedSheet);
        }
      }
    }
  }
}

// function to convert column letters to numbers
function columnLetterToNumber(letter) {
  letter = letter.toUpperCase();
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
