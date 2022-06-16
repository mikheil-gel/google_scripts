// custom trigger function
// !!!
// should be added manually:
// Apps Scipt > Triggers (clock icon) > Add Trigger
// > Choose which function to run: editEvent, Select event type: On edit > Save
// pop up will show up for granting permissions
function editEvent(e) {
  // !!!
  // function takes sheet name and target spreadsheet ID
  // more information below
  updateTargetSpreadsheet(e, 'Sheet1', '1qXSJDoqDEewRMRHkwQdLVoptm89zBKuKDbY8ZwKdo5U');
  updateTargetSpreadsheet(e, 'Sheet2', '1kTSOZXLTs7SaqknVjrqH3qlbNUtVdy8WjZ-KOn2LvZE', 'Sheet1');
}

// function for updating other spreadsheets
// takes 4 parameters:
// event object,
// local sheet name,
// target spreadsheet ID,
// target sheet name (optional: not needed if it's the same as the local name)
function updateTargetSpreadsheet(e, sheetName, ID, targetSheetName = sheetName) {
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

    // get target sheet
    const targetApp = SpreadsheetApp.openById(ID);
    const targetSheet = targetApp.getSheetByName(targetSheetName);

    // update target sheet values
    targetSheet.getRange(rangeNotation).setValues(rangeValues);
    // check if  the current sheet has formulas
    if (formulasAddresses.length) {
      // overwrite values with formulas in the target sheet
      formulasAddresses.forEach((cor) => {
        const formula = sheet.getRange(cor.row, cor.column).getFormula();
        targetSheet.getRange(cor.row, cor.column).setFormula(formula);
      });
    }
  }
}
