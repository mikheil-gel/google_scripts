# Instructions:

## The script should be copied to the parent spreadsheet's Apps Script editor;

## 6 values should be provided:

## 1) parent spreadsheet's id:

const parentSpreadsheetId = '11Rpx0aY-AaZqTlLAfu9mYl_S2AzziX5jhzA9EqaP3B4';

## 2) sheet name with data (in parent spreadsheet):

const parentSheetName = 'Sheet1';

## 3) number of the header rows:

const headerRows = 2;

## 4) column letter with unique values (e.g. employee name):

const uniqueValuesColumn = 'A';

#### columns were data should merge, can stay empty for now:

#### const columnsToMerge = [ ];

## 5) - 6) child spreadsheet's id and sheet name with filtered data:

#### first and second items of child array:

const childrenSpreadsheetsArray = [ [ '1NldYfCGJdgW2L7oDb_uhLEQYNID2_YNoCZLwdjB_lIw', 'Sheet1', 'event1' ] ];

function with name 'event1' is already created, so no more changes are required.

---

## The final step is to run 'createTriggers' to bind spreadsheets:

### on the editors top bar, choose createTriggers from the dropdown (event1 will be as default) and click run.
