// !!!
// this function should run manually from the Apps Script once
// creates change event trigger for the spreadsheet
// permissions should be granted
function createTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // check if trigger already exist
  const triggers = ScriptApp.getUserTriggers(ss);
  let triggerExists = false;
  triggers.forEach((item) => {
    if (item.getHandlerFunction() === 'changeEvent') triggerExists = true;
  });
  // create trigger if it's first run
  if (!triggerExists) {
    ScriptApp.newTrigger('changeEvent').forSpreadsheet(ss).onChange().create();
  }
}

function changeEvent(e) {
  // !!!
  // function takes sheet name and target spreadsheet Id
  // more information below
  updateTargetSpreadsheet(e, 'Sheet1', '1qXSJDoqDEewRMRHkwQdLVoptm89zBKuKDbY8ZwKdo5U');
  updateTargetSpreadsheet(e, 'Sheet2', '1kTSOZXLTs7SaqknVjrqH3qlbNUtVdy8WjZ-KOn2LvZE', 'Sheet1');
}

// function for updating other spreadsheets
// takes 4 parameters:
// event object,
// local sheet name,
// target spreadsheet Id,
// target sheet name (optional: not needed if it's the same as the local name)
function updateTargetSpreadsheet(e, sheetName, spreadsheetId, targetSheetName = sheetName) {
  // get edited sheet name
  const activeSheetName = e.source.getActiveSheet().getName();

  // check if edit happened on the current sheet
  if (sheetName === activeSheetName) {
    // get current sheet
    const app = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = app.getSheetByName(sheetName);

    // get range with data
    const dataRange = sheet.getDataRange();

    // get range address
    const rangeNotation = dataRange.getA1Notation();

    // get range values
    const rangeValues = dataRange.getValues();

    // get range formulass
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

    // get target sheet
    const targetApp = SpreadsheetApp.openById(spreadsheetId);
    const targetSheet = targetApp.getSheetByName(targetSheetName);
    const tergetRange = targetSheet.getRange(rangeNotation);

    // update target sheet values
    tergetRange.setValues(rangeValues);
    // check if  the current sheet has formulas
    if (formulasAddresses.length) {
      // overwrite values with formulas in the target sheet
      formulasAddresses.forEach((cor) => {
        const formula = sheet.getRange(cor.row, cor.column).getFormula();
        targetSheet.getRange(cor.row, cor.column).setFormula(formula);
      });
    }
    //update formatting
    tergetRange.setBackgrounds(background);
    tergetRange.setFontColors(fontColor);
    tergetRange.setFontFamilies(fontFamily);
    tergetRange.setFontLines(fontLine);
    tergetRange.setFontSizes(fontSize);
    tergetRange.setFontStyles(fontStyle);
    tergetRange.setFontWeights(fontWeight);
    tergetRange.setTextStyles(textStyle);
    tergetRange.setHorizontalAlignments(horAlign);
    tergetRange.setVerticalAlignments(vertAlign);

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

    for (let i = 1; i <= dataRange.getHeight(); i++) {
      let width = sheet.getColumnWidth(i);
      targetSheet.setColumnWidth(i, width);
    }

    for (let i = 1; i <= dataRange.getWidth(); i++) {
      let height = sheet.getRowHeight(i);
      targetSheet.setRowHeight(i, height);
    }
  }
}
