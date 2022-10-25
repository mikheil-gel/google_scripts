// !!!
// after providing all the necessary values below,
// user should run function 'createTriggers' from the menu bar above

// !!!
// user should provide local (parent) spreadsheet id
const localSpreadsheetId = '1fBs3N-zh7QmhLqG3lqzfNiuArRapOfI5ri58rZB-EBs';

// !!!
// user should add array/s in childrenSpreadsheetsArr;
// arrays should have 4 values:
// target (child) spreadsheet id;
// target (child) sheet name
// local (parent) sheet name
// name of the trigger function (should be created manually)
// expample: ['1qXSJDoqDEewRMRHkwQdLVoptm89zBKuKDbY8ZwKdo5U', 'Sheet1','Sheet1', 'event1' ]

const childrenSpreadsheetsArr = [
  ['14aTdO_E1aCRu-KN_olsfhXkZORiV1SMyxlHEdv_nG_8', 'Sheet1', 'Sheet1', 'event1'],
  // ['1MZKji9aGJlGlpoTgg77-wv2GPmIZHo2MytuAIaXFE4Y', 'Sheet1', 'Sheet2', 'event2'],
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
// copies data from the children spreadsheets to parent one;
// creates change event trigger for every spreadsheet
// permissions should be granted
function createTriggers() {
  // create an array of trigger function names
  let triggerFunctions = childrenSpreadsheetsArr.map((it) => it[3]);
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
  childrenSpreadsheetsArr.forEach((arr) => {
    // copy data from children spreadsheets to parent one;
    updateSpreadsheet({ data: arr });
    // create triggers for children spreadsheets
    ScriptApp.newTrigger(arr[3]).forSpreadsheet(arr[0]).onChange().create();
  });

  // create trigger for the parent spreadsheet
  ScriptApp.newTrigger('changeEvent').forSpreadsheet(localSpreadsheetId).onChange().create();
}

// trigger function for the parent spreadsheet
function changeEvent(e) {
  updateSpreadsheet(e);
}

// function for updating target spreadsheets
function updateSpreadsheet(e) {
  // get edited spreadsheet id
  const id = e.data ? e.data[0] : e.source.getId();
  // get edited sheet name
  const activeSheetName = e.data ? e.data[1] : e.source.getActiveSheet().getName();
  // check if edit happened on the parent spreadsheet
  const localChange = localSpreadsheetId === id;
  //create variables and asign them values, according to change event origin
  let targetId, targetSheetName, localSheetName;
  if (localChange) {
    [targetId, targetSheetName, localSheetName] =
      childrenSpreadsheetsArr.filter((arr) => arr[2] === activeSheetName)[0] || [];
  } else {
    [targetId, localSheetName, targetSheetName] =
      childrenSpreadsheetsArr.filter((arr) => arr[0] === id && arr[1] === activeSheetName)[0] || [];
  }

  // check if edit happened in the valid spreadsheet and sheet
  if (targetId) {
    // get spreadsheets
    const [changedSpreadsheet, targetSpreadsheet] = localChange
      ? [SpreadsheetApp.openById(localSpreadsheetId), SpreadsheetApp.openById(targetId)]
      : [SpreadsheetApp.openById(targetId), SpreadsheetApp.openById(localSpreadsheetId)];
    const sheet = changedSpreadsheet.getSheetByName(localSheetName);
    // get target sheet
    let targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

    // copy sheet to the parent spreadsheet if it's not present;
    if (!localChange && !targetSheet) {
      sheet.copyTo(targetSpreadsheet).setName(targetSheetName);
    } else {
      // copy data if both changed and target sheets exist
      if (sheet && targetSheet) {
        // check change event type
        if (e.changeType !== 'INSERT_GRID' && e.changeType !== 'REMOVE_GRID' && e.changeType !== 'OTHER') {
          // copy changed sheet to synced spreadsheet
          let copiedSheet = sheet.copyTo(targetSpreadsheet);
          // get copied range adrress
          let rangeNotation = copiedSheet.getDataRange().getA1Notation();
          const localLastRow = sheet.getLastRow();
          const localLastColumn = sheet.getLastColumn();

          // get target range adrress
          const targetLastRow = targetSheet.getLastRow();
          const targetLastColumn = targetSheet.getLastColumn();
          // clear tartget sheet if it has more rows/columns with data
          if (localLastRow < targetLastRow || localLastColumn < targetLastColumn) {
            targetSheet.clear();
            targetSheet.clearNotes();
          }
          // copy data to target sheet
          copiedSheet.getDataRange().copyTo(targetSheet.getRange(rangeNotation));
          // delete copied sheet
          targetSpreadsheet.deleteSheet(copiedSheet);
          // copy notes if changeType is 'OTHER'
        } else if (e.changeType === 'OTHER') {
          let localRange = sheet.getDataRange();
          let rangeNotation = localRange.getA1Notation();
          let localNotes = localRange.getNotes();

          targetSheet.getRange(rangeNotation).setNotes(localNotes);
        }
      }
    }
  }
}
