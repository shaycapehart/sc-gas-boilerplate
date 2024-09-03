/**
 * Custom formula
 * @customfunction
 */
function GETSHEETS(cmd) {
  return SpreadsheetApp.getActive()
    .getSheets()
    .map((sheet) => [sheet[`get${cmd}`]()]);
}

/**
 * Custom formula
 * @customfunction
 */
function SETSHEET(cmd, param) {
  return param
    ? SpreadsheetApp.getActiveSheet()[cmd](param)
    : SpreadsheetApp.getActiveSheet()[cmd]();
}
