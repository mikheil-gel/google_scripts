// !!!
// user should provide local (parent) spreadsheet id
const localSpreadsheetId = '1NnRkEQgg4kDj6-XXpv5G1dxFDbbUWZdA8pfy1dPzdUQ';

// !!!
// user should add array/s in childrenSpreadsheetsArr;
// arrays should have 4 values:
// target (child) spreadsheet id;
// target (child) sheet name
// local (parent) sheet name
// name of the trigger function (should be created manually)
// expample: ['1qXSJDoqDEewRMRHkwQdLVoptm89zBKuKDbY8ZwKdo5U', 'Sheet1','Sheet1', 'event1' ]

const childrenSpreadsheetsArr = [
  ['1qXSJDoqDEewRMRHkwQdLVoptm89zBKuKDbY8ZwKdo5U', 'Sheet1', 'Sheet1', 'event1'],
  ['1kTSOZXLTs7SaqknVjrqH3qlbNUtVdy8WjZ-KOn2LvZE', 'Sheet1', 'Sheet2', 'event2'],
];

// !!!
// functions should be created manually with name provided as 4th value of childern arrays above:
// e - event object should be passed as a parameter
// updateSpreadsheet(e) function should be called;
// example: function event1(e){ updateSpreadsheet(e) };

function event1(e) {
  updateSpreadsheet(e);
}
function event2(e) {
  updateSpreadsheet(e);
}

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
  // triggers for children spreadsheets
  childrenSpreadsheetsArr.forEach((arr) => {
    // copy data from children spreadsheets to parent one;
    copyData(
      false,
      SpreadsheetApp.openById(arr[0]).getSheetByName(arr[1]),
      SpreadsheetApp.openById(localSpreadsheetId),
      arr[2]
    );
    // create triggers for children spreadsheets
    ScriptApp.newTrigger(arr[3]).forSpreadsheet(arr[0]).onChange().create();
  });

  // create trigger for parent spreadsheet
  ScriptApp.newTrigger('changeEvent').forSpreadsheet(localSpreadsheetId).onChange().create();
}

// trigger function for the parent spreadsheet
function changeEvent(e) {
  updateSpreadsheet(e);
}

// function for updating target spreadsheets
function updateSpreadsheet(e) {
  // get edited spreadsheet id
  const id = e.source.getId();
  // get edited sheet name
  const activeSheetName = e.source.getActiveSheet().getName();
  // check if edit happened on the parent spreadsheet
  const localChange = localSpreadsheetId === id;
  //create variables and asign them values, according to change event origin
  let targetId, targetSheetName, localSheetName;
  if (localChange) {
    let arrVal = childrenSpreadsheetsArr.filter((arr) => arr[2] === activeSheetName);
    if (arrVal.length) [targetId, targetSheetName, localSheetName] = arrVal[0];
  } else {
    let arrVal = childrenSpreadsheetsArr.filter((arr) => arr[0] === id && arr[1] === activeSheetName);
    if (arrVal.length) [targetId, localSheetName, targetSheetName] = arrVal[0];
  }

  // check if edit happened on valid spreadsheet and sheet
  if (targetId) {
    // get spreadsheets
    const [app, targetApp] = localChange
      ? [SpreadsheetApp.openById(localSpreadsheetId), SpreadsheetApp.openById(targetId)]
      : [SpreadsheetApp.openById(targetId), SpreadsheetApp.openById(localSpreadsheetId)];
    const sheet = app.getSheetByName(localSheetName);
    // run function for copying data
    copyData(localChange, sheet, targetApp, targetSheetName);
  }
}

// function for copying data
function copyData(localChange, sheet, targetApp, targetSheetName) {
  // get target sheet
  let targetSheet = targetApp.getSheetByName(targetSheetName);

  // create sheet in the parent spreadsheet if it is not present;
  if (!localChange && !targetSheet) {
    targetSheet = targetApp.insertSheet();
    targetSheet.setName(targetSheetName);
  }
  // copy data if both changed and target sheets exist
  if (sheet && targetSheet) {
    // get range with data
    const dataRange = sheet.getDataRange();

    // get range address
    const rangeNotation = dataRange.getA1Notation();

    // get range values
    const rangeValues = dataRange.getValues();

    // get range formulas
    const rangeFormulas = dataRange.getFormulas();

    // create an array and save formulas addresses
    let formulasAddresses = [];
    rangeFormulas.forEach((val, rowIndex) => {
      val.forEach((cell, columnIndex) => {
        if (cell) {
          formulasAddresses.push({ row: rowIndex + 1, column: columnIndex + 1 });
        }
      });
    });

    // get formatting
    const background = dataRange.getBackgrounds();
    const fontColor = dataRange.getFontColors();
    const fontFamily = dataRange.getFontFamilies();
    const fontLine = dataRange.getFontLines();
    const fontSize = dataRange.getFontSizes();
    const fontStyle = dataRange.getFontStyles();
    const fontWeight = dataRange.getFontWeights();
    const textStyle = dataRange.getTextStyles();
    const horAlign = dataRange.getHorizontalAlignments();
    const vertAlign = dataRange.getVerticalAlignments();
    const bandings = dataRange.getBandings();
    const mergedRanges = dataRange.getMergedRanges();
    const notes = dataRange.getNotes();

    // get target range
    const targetRange = targetSheet.getRange(rangeNotation);

    // update target sheet values
    targetRange.setValues(rangeValues);
    // check if  the current sheet has formulas
    if (formulasAddresses.length) {
      // overwrite values with formulas in the target sheet
      formulasAddresses.forEach((cor) => {
        const formula = sheet.getRange(cor.row, cor.column).getFormula();
        targetSheet.getRange(cor.row, cor.column).setFormula(formula);
      });
    }
    //update formatting
    targetRange.setBackgrounds(background);
    targetRange.setFontColors(fontColor);
    targetRange.setFontFamilies(fontFamily);
    targetRange.setFontLines(fontLine);
    targetRange.setFontSizes(fontSize);
    targetRange.setFontStyles(fontStyle);
    targetRange.setFontWeights(fontWeight);
    targetRange.setTextStyles(textStyle);
    targetRange.setHorizontalAlignments(horAlign);
    targetRange.setVerticalAlignments(vertAlign);
    targetRange.setNotes(notes);

    for (let i in bandings) {
      let srcBandA1 = bandings[i].getRange().getA1Notation();
      let destBandRange = targetSheet.getRange(srcBandA1);

      destBandRange
        .applyRowBanding()
        .setFirstRowColor(bandings[i].getFirstRowColor())
        .setSecondRowColor(bandings[i].getSecondRowColor())
        .setHeaderRowColor(bandings[i].getHeaderRowColor())
        .setFooterRowColor(bandings[i].getFooterRowColor());
    }

    for (let i = 0; i < mergedRanges.length; i++) {
      targetSheet.getRange(mergedRanges[i].getA1Notation()).merge();
    }

    for (let i = 1; i <= dataRange.getWidth(); i++) {
      let width = sheet.getColumnWidth(i);
      targetSheet.setColumnWidth(i, width);
    }

    for (let i = 1; i <= dataRange.getHeight(); i++) {
      let height = sheet.getRowHeight(i);
      targetSheet.setRowHeight(i, height);
    }
  }
}
