// !!!
// after providing all the necessary values below,
// user should run function 'createTriggers' from the menu bar above

// parent (local) spreadsheet spreadsheetId
const parentSpreadsheetId = '19sKN5j_jgc0ZR1I0_JwP_kVA_J6Sb6mlFJirLayzieU';

// parent (local) sheet name, that should be synced
const parentSheetName = 'Sheet1';

// header rows count on the parent (local) sheet
const headerRows = 1;

// array of columns letters, where data should merge
const columnsToMerge = ['F', 'G'];

// !!!
// user should add array/s in childrenSpreadsheetsArray;
// arrays should have 3 values:
// target (child) spreadsheet id;
// target (child) sheet name
// name of the trigger function (should be created manually)
// example: ['1qXSJDoqDEewRMRHkwQdLVoptm89zBKuKDbY8ZwKdo5U', 'Sheet1', 'event1' ]

const childrenSpreadsheetsArray = [
  ['14aTdO_E1aCRu-KN_olsfhXkZORiV1SMyxlHEdv_nG_8', 'Sheet2', 'event1'],
  // ['1MZKji9aGJlGlpoTgg77-wv2GPmIZHo2MytuAIaXFE4Y', 'Sheet3', 'event2'],
];

// !!!
// functions should be created manually with name, provided as 4th value of childern arrays above:
// e - event object should be passed as a parameter
// updateSpreadsheet(e) function should be called;
// example: function event1(e){ updateSpreadsheet(e) };

function event1(e) {
  updateSpreadsheet(e);
}

// function event2(e) {
//   updateSpreadsheet(e);
// }

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
    ScriptApp.newTrigger(arr[2]).forSpreadsheet(arr[0]).onChange().create();
  });

  // create trigger for the parent spreadsheet
  ScriptApp.newTrigger('changeEvent').forSpreadsheet(parentSpreadsheetId).onChange().create();
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
    let childSpreadsheetData = childrenSpreadsheetsArray.filter(
      (arr) => arr[0] === spreadsheetId && arr[1] === activeSheetName
    );
    if (childSpreadsheetData.length) {
      spreadsheetsDataArray = [
        [parentSpreadsheetId, parentSheetName, childrenSpreadsheetsArray[0][0], childrenSpreadsheetsArray[0][1]],
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
      // get last column with data
      const lastColumn = sheet.getLastColumn();
      // get last row with data
      const lastRow = sheet.getLastRow();
      // get target sheet
      let targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
      // create target sheet in the child spreadsheet if not present
      if (localChange && sheet && !targetSheet) {
        targetSheet = targetSpreadsheet.insertSheet().setName(targetSheetName);
      }
      // copy data if both changed and target sheets exist
      if (sheet && targetSheet) {
        // check if data should be copied

        const localColumnData = sheet.getRange(headerRows + 1, 1, lastRow - headerRows).getValues();
        const targetColumnData = targetSheet.getRange(headerRows + 1, 1, lastRow - headerRows).getValues();

        let dataToCopy = [];

        localColumnData.forEach((localValue, localRowIndex) => {
          if (localValue[0].toString().length) {
            targetColumnData.forEach((targetValue, targetRowIndex) => {
              if (targetValue[0].toString().length) {
                if (localValue[0] === targetValue[0]) {
                  dataToCopy.push({ localRow: localRowIndex + 1, targetRow: targetRowIndex + 1 });
                }
              }
            });
          }
        });
        // console.log({dataToCopy})
        // check if data in columns should merge
        if (columnsToMerge.length) {
          dataToCopy.forEach((rowData) => {
            let dataToMerge = [];
            // create array to store data
            // get values of cells in 'to be merged' columns in the target sheet
            columnsToMerge.forEach((columnLetter) => {
              let columnNumber = columnLetterToNumber(columnLetter);
              let columnValue = targetSheet.getRange(rowData.targetRow, columnNumber, 1).getValue();
              if (columnValue.toString().length) {
                dataToMerge.push({
                  localRow: rowData.localRow,
                  column: columnNumber,
                  value: columnValue,
                });
              }
            });
            // compare values of target sheet with local (changed) sheet and merge
            dataToMerge.forEach((data) => {
              // get local values
              let localValue = sheet.getRange(data.localRow, data.localRow.column).getValue();
              // if changed values doesn't include old value, then merge
              if (!localValue.toString().includes(data.value)) {
                dataRange.setValue(data.value + ' ' + localValue);
              }
            });
          });
        }

        if (dataToCopy.length) {
          let copiedSheet = sheet.copyTo(targetSpreadsheet);
          for (let rowData of dataToCopy) {
            // copy changed sheet to target spreadsheet
            // get range data
            let copiedDataRange = copiedSheet.getRange(rowData.localRow, 1, 1, lastColumn);
            let copyToDataRange = targetSheet.getRange(rowData.targetRow, 1, 1, lastColumn);

            // check change event type
            if (e.changeType !== 'INSERT_GRID' && e.changeType !== 'REMOVE_GRID' && e.changeType !== 'OTHER') {
              // copy table data to target sheet
              copiedDataRange.copyTo(copyToDataRange);
              // only copy notes if change type is 'OTHER'
            } else if (e.changeType === 'OTHER') {
              let localNotes = copiedDataRange.getNotes();
              copyToDataRange.setNotes(localNotes);
            }
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
